from typing import *
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from reportlab.pdfgen import canvas
from reportlab.lib import utils, colors
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.lib import utils, colors
from reportlab.lib.colors import HexColor
import fitz
import pandas as pd
from copy import copy
from collections import OrderedDict
import math
import shutil
import _pickle as pickle
import _thread
import time

from pathlib import Path
from datetime import datetime
from PIL import Image
import os


import doc_maker,dc
import id_card_maker
import report_card_maker
import util

global Input

import os

root = tk.Tk()
import time
# create root window

if not os.path.isdir("finalcovers"):
   os.makedirs("finalcovers")
   
   
folder = ""
file = ""
id_card_file = ""
report_card_file = ""


if os.path.isdir("STICKERS")==False:
   os.makedirs("STICKERS")

#DOCDRIVER = DOC["DriverData"]
print_jobs = []

SEEKPOS = 9938
Input=(420*2.83465,(290*2.83465))

school_vars = []
schools_inner_frame = None
school_canvas = None
kid_index_entry = None
status_label = None
size_info_label = None
scale_info_label = None


def set_status_message(message: str, color: str = "green") -> None:
   """Safely update the status label from any thread."""
   if status_label is None:
      return

   def _update() -> None:
      status_label.configure(text=message, fg=color)

   try:
      status_label.after(0, _update)
   except RuntimeError:
      # The widget may have been destroyed during shutdown.
      pass


class SchoolDocuments(NamedTuple):
   school_id: Optional[str]
   label: str
   pdfs: List[Path]


def _sanitize_school_label(value: str) -> str:
   value = (value or "").strip()
   fallback = "School"
   if not value:
      value = fallback
   if hasattr(dc, "_sanitize_for_path"):
      sanitized = dc._sanitize_for_path(value, fallback)
   else:
      sanitized = "".join(ch if ch.isalnum() else "_" for ch in value)
      sanitized = "_".join(part for part in sanitized.split("_") if part)
      sanitized = sanitized or fallback
   return sanitized


def _merge_pdf_files(pdf_paths: Sequence[Path], output_path: Path) -> bool:
   combined = fitz.open()
   try:
      for pdf_path in pdf_paths:
         try:
            with fitz.open(str(pdf_path)) as src:
               combined.insert_pdf(src)
         except Exception as exc:
            print(f"Failed to include {pdf_path}: {exc}")
      if combined.page_count:
         output_path.parent.mkdir(parents=True, exist_ok=True)
         combined.save(str(output_path))
         return True
      return False
   finally:
      combined.close()


def _create_verification_pdf(pdf_paths: Sequence[Path], output_path: Path) -> bool:
   verification_doc = fitz.open()
   matrix = fitz.Matrix(100 / 72, 100 / 72)
   try:
      for pdf_path in pdf_paths:
         try:
            with fitz.open(str(pdf_path)) as src:
               for page in src:
                  pix = page.get_pixmap(matrix=matrix, alpha=False)
                  width_pt = pix.width * 72 / 100
                  height_pt = pix.height * 72 / 100
                  new_page = verification_doc.new_page(width=width_pt, height=height_pt)
                  new_page.insert_image(new_page.rect, stream=pix.tobytes("png"))
         except Exception as exc:
            print(f"Failed to render {pdf_path}: {exc}")
      if verification_doc.page_count:
         output_path.parent.mkdir(parents=True, exist_ok=True)
         verification_doc.save(str(output_path))
         return True
      return False
   finally:
      verification_doc.close()


def _collect_child_pdf_paths(school_dir: Path) -> List[Path]:
   pdf_paths: List[Path] = []
   child_dirs = [
      entry
      for entry in school_dir.iterdir()
      if entry.is_dir() and entry.name.lower() not in {"working"}
   ]
   for child_dir in sorted(child_dirs, key=lambda path: path.name.lower()):
      front_path: Optional[Path] = None
      back_path: Optional[Path] = None
      for pdf_path in sorted(child_dir.glob("*.pdf"), key=lambda path: path.name.lower()):
         upper_name = pdf_path.name.upper()
         if upper_name.endswith("_FRONT.PDF"):
            front_path = pdf_path
         elif upper_name.endswith("_BACK.PDF"):
            back_path = pdf_path
      if front_path is not None:
         pdf_paths.append(front_path)
      if back_path is not None:
         pdf_paths.append(back_path)
   return pdf_paths


def _collect_cover_documents(cover_root: Path) -> List[SchoolDocuments]:
   documents: List[SchoolDocuments] = []
   if not cover_root.is_dir():
      return documents
   for entry in sorted(cover_root.iterdir(), key=lambda path: path.name.lower()):
      if not entry.is_dir():
         continue
      pdfs = sorted(entry.glob("*.pdf"), key=lambda path: path.name.lower())
      if not pdfs:
         continue
      school_id: Optional[str] = None
      label_source = entry.name
      parts = entry.name.split("_", 1)
      if len(parts) == 2:
         potential_id, remainder = parts
         school_id = potential_id or None
         label_source = remainder
      documents.append(
         SchoolDocuments(school_id, _sanitize_school_label(label_source), pdfs)
      )
   return documents


def _collect_report_card_documents(report_root: Path) -> List[SchoolDocuments]:
   documents: List[SchoolDocuments] = []
   if not report_root.is_dir():
      return documents
   for school_dir in sorted(report_root.iterdir(), key=lambda path: path.name.lower()):
      if not school_dir.is_dir():
         continue
      pdfs = _collect_child_pdf_paths(school_dir)
      if not pdfs:
         continue
      school_id: Optional[str] = None
      label_source = school_dir.name
      parts = school_dir.name.split("_", 1)
      if len(parts) == 2:
         potential_id, remainder = parts
         school_id = potential_id or None
         label_source = remainder
      documents.append(
         SchoolDocuments(school_id, _sanitize_school_label(label_source), pdfs)
      )
   return documents


def _collect_id_card_documents(
   id_root: Path, name_overrides: Dict[str, str]
) -> List[SchoolDocuments]:
   documents: List[SchoolDocuments] = []
   if not id_root.is_dir():
      return documents
   for school_dir in sorted(id_root.iterdir(), key=lambda path: path.name.lower()):
      if not school_dir.is_dir():
         continue
      pdfs = _collect_child_pdf_paths(school_dir)
      if not pdfs:
         continue
      school_id = school_dir.name
      label = name_overrides.get(school_id) or _sanitize_school_label(school_dir.name)
      documents.append(SchoolDocuments(school_id, label, pdfs))
   return documents


def _build_print_label(documents: Sequence[SchoolDocuments]) -> str:
   labels = sorted({doc.label for doc in documents})
   if not labels:
      return "School"
   if len(labels) == 1:
      return labels[0]
   return "_".join(labels)


def _load_tabular_file(path: str, **kwargs):
   """Load a spreadsheet or CSV file based on its extension."""
   suffix = Path(path).suffix.lower()
   if suffix == ".csv":
      return pd.read_csv(path, **kwargs)
   return pd.read_excel(path, **kwargs)

def storeDocs2(subject):
   
   
   combined = fitz.open()
   totpgs = 0
   binderNum = 0
   #DRIVER = pickle.loads(bytes(info["driver"]))
   # for file in os.listdir("PDFS"  + "/" + subject):
     
   #    for pdf in os.listdir("PDFS"  + "/" + subject+"/" +file):
     
   #       totpgs += 1
   #       with fitz.open(os.path.join("PDFS"  + "/" + subject, file)) as outfile :
            
   #          combined.insert_pdf(outfile)
        
         
            
   #    if totpgs > 300:
         
   #       combined.save(r"STICKERS" + '\\' +  subject + str(binderNum) +  '.pdf')
   #       binderNum += 1
   #       totpgs = 0
   #       combined = fitz.open()
   for directory in os.listdir("PDFS"  + "/" + subject):
         
      for pdf in os.listdir("PDFS"  + "/" + subject+"/" +directory):
            
            totpgs += 1
            with fitz.open(os.path.join("PDFS"  + "/" + subject+"/" +directory,pdf)) as outfile :
            
               combined.insert_pdf(outfile)
        
                 
         
            
      if totpgs > 300:
         
         combined.save(r"STICKERS" + '\\' +  subject + str(binderNum) +  '.pdf')
         binderNum += 1
         totpgs = 0
         combined = fitz.open()
   

   if totpgs != 0:
      
      combined.save(r"STICKERS" + '\\' +  subject + str(binderNum) +  '.pdf')
            
   else:
      raise Exception("pages are zero")

def storeDocs(subject,info, scholname = ''):
   
   global Input
   combined = fitz.open()
   totpgs = 0
   binderNum = 0
   #DRIVER = pickle.loads(bytes(info["driver"]))
   for file in os.listdir("PDFS"  + "/" + subject):
    
      k=util.build_doc("PDFS" + "/" + subject + "/" + file, "", file, info["num"],Input,combined,page_scale, 0, True, scholname)
     
      totpgs += info["num"]
      
      
      if totpgs > 300:
         
         combined.save(r"FINAL BINDERS" + '\\' + scholname + '_' +  subject + str(binderNum) +  '.pdf')

         print_jobs.append({"path" : r"FINAL BINDERS", "name" : subject + str(binderNum) +  '.pdf', "driver": "", "checkFORM" : False})
         binderNum += 1
         totpgs = 0
         combined = fitz.open()
   

   if totpgs != 0:
      
      combined.save(r"FINAL BINDERS" + '\\'  + scholname + '_' +  subject + str(binderNum) +  '.pdf')

      print_jobs.append({"path" : r"FINAL BINDERS", "name" : subject + str(binderNum) +  '.pdf', "driver": "", "checkFORM" : False})      


   
   
   # for n in dir(properties["pDevMode"]):
   #    if n in DRIVER.keys():
   #       try:   
   #          setattr(properties["pDevMode"], n, DRIVER[n])
   #       except:
   #          print(n + '->')
   #          pass
   
   # win32print.SetPrinter(pHandle,2,properties,0)
   # win32api.ShellExecute(0, "print", subject + '.pdf', None,  "./FINAL BINDERS",  0)   
   
   



def storePS (tuple, prev, subDict):
   
   subject = str(tuple["inner_code"]).zfill(7)

   if subject != prev:
        if os.path.isdir("PDFS"  + "/" + subject):
            shutil.rmtree("PDFS"  + "/" + subject)
   
   if subject != prev and prev != "": 
      if prev.endswith("s"):
         storeDocs2(prev)
      else:
        storeDocs(prev, subDict[prev], tuple["school_name"])
   
   if subject.endswith("s"):
      doc_maker.personalize(tuple, subject,  CheckVar2.get(),True)
   else:      
      try:
         subDict[subject]
      except KeyError:   
         subjectFolder = r"\\pixartnas\home\INTERNAL_PROCESSING\ALL BOOKS FORM\NONP"+"\\"+subject

         
         pdf_files=[f for f in os.listdir(subjectFolder) if f.endswith('.pdf')]
         numpages = len(pdf_files)

         if os.path.isfile(r"\\pixartnas\home\INTERNAL_PROCESSING\FORM"+"\\" + subject + '.pdf') == False:
            
            formPath = util.build_doc(subjectFolder, r"\\pixartnas\home\INTERNAL_PROCESSING\FORM", subject, numpages,Input,None,page_scale)
            
            
         
         
         # for n in dir(properties["pDevMode"]):
         #    if n in FORM.keys():
         #       try:   
         #          setattr(properties["pDevMode"], n, FORM[n])
         #       except:
         #          print(n + '->')
         #          pass
         
         # win32print.SetPrinter(pHandle,2,properties,0)

         # win32api.ShellExecute(0, "print", subject + '.pdf', None,  "./FORM",  0)
         
         
         # print("checkVar " + str(CheckVar1))
         
         # if CheckVar1.get() == 0:
         #    print_jobs.append({"path" : "./FORM", "name" : subject + '.pdf', "driver": FORM, "checkFORM" : True})

         #print(getattr(properties["pDevMode"], "DriverData"))
         
         # bytes_from_list = bytearray(DOC)
         

         # for pos in range(SEEKPOS, SEEKPOS + 13, 2):
         #    bytes_from_list[pos] = ord(subject[int((pos - SEEKPOS)/2)])
         
         subDict[subject] = {"processed": True, "driver": 0, "num": numpages, "ref": "", "school": tuple["school_name"]}
      

      
      doc_maker.personalize(tuple, subject, CheckVar2.get(),False)
   # storeDocs(subject, str(tuple["BOOKID"]).zfill(3), subDict[subject])



def populate_school_checkboxes(sheet_path: Optional[str] = None):
   global school_vars, schools_inner_frame, school_canvas, status_label, file, id_card_file, report_card_file

   if schools_inner_frame is None:
      return

   for child in schools_inner_frame.winfo_children():
      child.destroy()

   school_vars = []

   paths_to_scan = []
   if sheet_path:
      paths_to_scan.append(sheet_path)
   if file and file not in paths_to_scan:
      paths_to_scan.append(file)
   if id_card_file and id_card_file not in paths_to_scan:
      paths_to_scan.append(id_card_file)
   if report_card_file and report_card_file not in paths_to_scan:
      paths_to_scan.append(report_card_file)

   if not paths_to_scan:
      return

   schools = set()
   for path in paths_to_scan:
      try:
         df = _load_tabular_file(path, header=0)
      except Exception as exc:
         print(f"Failed to read '{path}': {exc}")
         continue

      if "school_name" not in df.columns:
         continue

      names = {
         str(name).strip()
         for name in df["school_name"].dropna()
         if str(name).strip() != ""
      }
      schools.update(names)

   if not schools:
      if status_label is not None:
         status_label.configure(fg="red", text="No schools found in selected sheets.")
      return

   for name in sorted(schools):
      var = tk.IntVar(value=0)
      chk = tk.Checkbutton(schools_inner_frame, text=name, variable=var, anchor="w", justify="left")
      chk.pack(fill="x", anchor="w")
      school_vars.append((var, name))

   if school_canvas is not None:
      school_canvas.yview_moveto(0)

   if status_label is not None:
      status_label.configure(fg="green", text=f"Loaded {len(schools)} school(s). Select to process.")


def open_win_diag():
   # Create a dialog box
   global file
   file=filedialog.askopenfilename(initialdir="E:/winapp")

   if file:
      populate_school_checkboxes(file)


def open_id_card_diag():
   global id_card_file
   id_card_file = filedialog.askopenfilename(initialdir="E:/winapp")

   if id_card_file:
      populate_school_checkboxes(id_card_file)


def open_report_card_diag():
   global report_card_file
   report_card_file = filedialog.askopenfilename(initialdir="E:/winapp")

   if report_card_file:
      populate_school_checkboxes(report_card_file)



def windowDialog():
   _thread.start_new_thread(make, (0,))


def convert_cmyk_to_rgb(input_path, output_path):
    with Image.open(input_path) as img:
        if img.mode == "CMYK":
            img = img.convert("RGB")
        img.save(output_path)

def make(dum):
   spine=0
   global file,folder, school_vars, kid_index_entry, status_label, selected_size, id_card_file, report_card_file

   processing_books = checkVar3.get()==1 or checkVar4.get()==1
   processing_id_cards = checkVar5.get()==1
   processing_report_cards = checkVar6.get()==1

   if not processing_books and not processing_id_cards and not processing_report_cards:
      if status_label is not None:
         status_label.configure(fg="red", text="Select at least one processing option.")
      return

   if processing_books and file == "":
      if status_label is not None:
         status_label.configure(fg="red", text = "Select the book Excel file before processing.")
      return

   if processing_id_cards and id_card_file == "":
      if status_label is not None:
         status_label.configure(fg="red", text = "Select the ID card Excel file before processing.")
      return

   if processing_report_cards and report_card_file == "":
      if status_label is not None:
         status_label.configure(fg="red", text = "Select the report card Excel file before processing.")
      return

   current_size = ""
   if 'selected_size' in globals() and selected_size is not None:
      current_size = selected_size.get()

   if processing_books and current_size == "":
      if status_label is not None:
         status_label.configure(fg="red", text="ERROR SELECT A BOOKLET SIZE ")
      return

   if processing_books and os.path.isdir("FINAL BINDERS") == False:
      os.makedirs("FINAL BINDERS", exist_ok=True)

   kid_idx = ""
   if kid_index_entry is not None:
      kid_idx = kid_index_entry.get().strip()

   kid_idx = kid_idx if kid_idx is not None else ""

   selected_school_names = {name for var, name in school_vars if var.get() == 1}

   process_single_kid = False
   target_user_id = None
   target_book_id = None

   if kid_idx:
      process_single_kid = True
      if kid_idx.upper().startswith('U'):
         target_user_id = kid_idx[1:].strip()
         if target_user_id == "":
            if status_label is not None:
               status_label.configure(fg="red", text="Enter a valid user id after 'U'.")
            return
      else:
         try:
            target_book_id = int(kid_idx)
         except ValueError:
            if status_label is not None:
               status_label.configure(fg="red", text="Kid index must be numeric or start with 'U'.")
            return
   elif not selected_school_names:
      if status_label is not None:
         status_label.configure(fg="red", text="Select at least one school or enter a kid index.")
      return


   def matches_selection(item: Dict[str, Any]) -> bool:
      if process_single_kid:
         if target_user_id is not None:
            return str(item.get("user_id", "")).strip() == target_user_id
         if target_book_id is not None:
            try:
               return int(item.get("book_id")) == target_book_id
            except (TypeError, ValueError):
               return False
         return False

      school_value = item.get("school_name", "")
      if pd.isna(school_value):
         school_value = ""
      else:
         school_value = str(school_value).strip()
      return school_value in selected_school_names


   if status_label is not None:
      status_label.configure(text="", fg="red")


   data: Dict[Any, Dict[str, Any]] = {}
   sheet_has_multiple_schools = False
   subjectIDX: Dict[str, Dict[str, Any]] = {}
   prev = ""

   colorcodes = ['#ff0000', '#00ff00', '#8F7C00' ,'#0000ff', '#993F00'  ,'#ffff00', '#00ffff', '#ff00ff', '#000000', '#888888',
'#234567']

   if processing_books:
      df = _load_tabular_file(file, header=0)
      if "school_name" in df.columns:
         cleaned_names = {
            str(name).strip()
            for name in df["school_name"].dropna()
            if str(name).strip() != ""
         }
         sheet_has_multiple_schools = len(cleaned_names) > 1
      data = df.to_dict('index')

   cover_pages_created = 0

   if checkVar3.get()==1:
      
      if not os.path.isdir("store"):
         os.makedirs("store")
 
      for key,item in data.items():
         if pd.isna(item["last_name"]):
            item["last_name"] = ''
            print(item["book_id"])
         if not matches_selection(item):
            continue
         
         if str(item["outer_code"]).endswith("s") or str(item["outer_code"]).strip()=="":
            continue
         
         item["outer_code"] = str(item["outer_code"]).zfill(7)
         
         full_path = os.path.join(r"\\pixartnas\home\INTERNAL_PROCESSING\SCHOOLCOVERS"+"\\" + str(item["outer_code"])[:3]  , str(item["outer_code"])+".svg")
         
         if not os.path.isfile(full_path):
            continue
         
         if os.path.isdir("Temp"+"/"+str(item["outer_code"])[0:3])==False:
            os.makedirs("Temp"+"/"+str(item["outer_code"])[0:3])
            school_covers_folder=r"\\pixartnas\home\INTERNAL_PROCESSING\SCHOOLCOVERS"+"\\"+str(item["outer_code"])[0:3]
           
            for file in os.listdir(school_covers_folder):
               name, ext = os.path.splitext(file)
               if file == "Assets" and os.path.isdir(os.path.join(school_covers_folder, file)):
                  if not os.path.isdir("Temp" + "/" + str(item["outer_code"])[:3]+"/" + file):
                     os.makedirs("Temp" + "/" + str(item["outer_code"])[:3]+"/" + file)
                  for sub_file in os.listdir(school_covers_folder+"/"+ file):
                     sub_name, sub_ext = os.path.splitext(sub_file)
                     if not sub_ext.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')) :
                        shutil.copyfile(school_covers_folder+"/"+ file + "/" + sub_file, "Temp" + "/" + str(item["outer_code"])[:3]+"/" + file + "/" + sub_file) 
                     else:
                        convert_cmyk_to_rgb (school_covers_folder+"/"+ file + "/" + sub_file, "Temp" + "/" + str(item["outer_code"])[:3]+"/" + file + "/" + sub_file)
               
               elif os.path.isfile(os.path.join(school_covers_folder, file)) and  ext!= '.svg':
                  if not ext.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')) :
                     shutil.copyfile(school_covers_folder+"/"+ file, "Temp" + "/" + str(item["outer_code"])[:3]+"/"+file)
                  else:
                     convert_cmyk_to_rgb (school_covers_folder+"/"+ file, "Temp" + "/" + str(item["outer_code"])[:3]+"/"+file)
                  continue
         
         
         school_color_code1=colorcodes[int(int(item["school_color_id"])/11)]
         school_color_code2=colorcodes[int((item["school_color_id"]))%11]
         
         grade_colour_code=colorcodes[int(int(item["class_color_id"])%11)]

         
         kid_color_code1=colorcodes[int(int((item["user_color_id"]))/11)]
         kid_color_code2=colorcodes[int((item["user_color_id"])%11)]
         print(kid_color_code2)
         
         photo_name = item["user_id"]
         
         dc.personalize(item["outer_code"], photo_name , item["school_id"],school_color_code1,school_color_code2,grade_colour_code,kid_color_code1,kid_color_code2,item["first_name"]+" "+item["last_name"],str(item["book_id"]).zfill(3),item,sheet_has_multiple_schools)
         cover_pages_created += 1
   
   if checkVar4.get()==1:
      for key,item in data.items():
         if not matches_selection(item):
            continue
         if str(item["inner_code"]).endswith("b") or str(item["inner_code"]).strip()=="":
            continue
         full_path = os.path.join(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL BOOKS FORM\NONP", str(item["inner_code"]).zfill(7) )
         if not os.path.exists(full_path) and  not str(item["inner_code"]).endswith("s"):
            continue
         
         storePS(item, prev, subjectIDX)
         prev = str(item["inner_code"]).zfill(7)
      if prev.endswith("s"):
         storeDocs2(prev)
      else:
         storeDocs(prev, subjectIDX[prev], item["school_name"])

   id_cards_created = 0
   if processing_id_cards:
      try:
         id_df = _load_tabular_file(id_card_file, header=0)
      except Exception as exc:
         if status_label is not None:
            status_label.configure(fg="red", text=f"Failed to load ID card sheet: {exc}")
         return

      for record in id_df.to_dict("records"):
         if not matches_selection(record):
            continue

         try:
            if id_card_maker.personalize_id_card(record):
               id_cards_created += 1
         except id_card_maker.TemplateNotFoundError as exc:
            print(exc)
         except Exception as exc:
            print(f"Failed to generate ID card for {record.get('user_id')}: {exc}")

   report_cards_created = 0
   if processing_report_cards:
      try:
         report_df = _load_tabular_file(report_card_file, header=0)
      except Exception as exc:
         if status_label is not None:
            status_label.configure(fg="red", text=f"Failed to load report card sheet: {exc}")
         return

      for record in report_df.to_dict("records"):
         if not matches_selection(record):
            continue

         try:
            if report_card_maker.personalize_report_card(record):
               report_cards_created += 1
         except report_card_maker.TemplateNotFoundError as exc:
            print(exc)
         except Exception as exc:
            print(f"Failed to generate report card for {record.get('user_id')}: {exc}")

   status_messages = []
   status_color = "green"

   if checkVar3.get()==1:
      if cover_pages_created:
         status_messages.append("Cover pages generated.")
      else:
         status_messages.append("No cover pages generated.")
         status_color = "red"

   if processing_id_cards:
      if id_cards_created:
         status_messages.append(f"Generated {id_cards_created} ID card(s).")
      else:
         status_messages.append("No ID cards generated.")
         status_color = "red"

   if processing_report_cards:
      if report_cards_created:
         status_messages.append(f"Generated {report_cards_created} report card(s).")
      else:
         status_messages.append("No report cards generated.")
         status_color = "red"

   if status_messages:
      set_status_message(" ".join(status_messages), status_color)


def _merge_cover_pages_worker() -> None:
   set_status_message("Merging documents...", "blue")

   date_prefix = datetime.now().strftime("%Y%m%d")
   print_root = Path("Binders for Print")
   verification_root = Path("Binders for Verification")
   print_root.mkdir(parents=True, exist_ok=True)
   verification_root.mkdir(parents=True, exist_ok=True)

   cover_documents = _collect_cover_documents(Path("finalcovers"))
   report_documents = _collect_report_card_documents(Path("Report cards"))

   name_overrides: Dict[str, str] = {}
   for doc in cover_documents + report_documents:
      if doc.school_id:
         name_overrides[doc.school_id] = doc.label

   id_documents = _collect_id_card_documents(Path("ID Cards"), name_overrides)

   messages: List[str] = []
   any_success = False

   if cover_documents:
      cover_print_label = _build_print_label(cover_documents)
      cover_print_name = (
         f"{date_prefix}_{cover_print_label}_COVER_Own_Sheet_SS_1 copy.pdf"
      )
      cover_print_path = print_root / cover_print_name
      cover_print_created = _merge_pdf_files(
         [pdf for doc in cover_documents for pdf in doc.pdfs], cover_print_path
      )
      cover_verification_created = 0
      for doc in cover_documents:
         verification_name = (
            f"{date_prefix}_{doc.label}_COVER_verify.pdf"
         )
         if _create_verification_pdf(doc.pdfs, verification_root / verification_name):
            cover_verification_created += 1
      if cover_print_created or cover_verification_created:
         any_success = True
         messages.append(
            f"Merged cover pages for {len(cover_documents)} school(s)."
         )
      else:
         messages.append("No cover pages found to merge.")
   else:
      messages.append("No cover pages found to merge.")

   if report_documents:
      report_print_label = _build_print_label(report_documents)
      report_print_name = (
         f"{date_prefix}_{report_print_label}_report_card_Own_Sheet_BB_1 copy.pdf"
      )
      report_print_path = print_root / report_print_name
      report_print_created = _merge_pdf_files(
         [pdf for doc in report_documents for pdf in doc.pdfs], report_print_path
      )
      report_verification_created = 0
      for doc in report_documents:
         verification_name = (
            f"{date_prefix}_{doc.label}_report_card_verify.pdf"
         )
         if _create_verification_pdf(doc.pdfs, verification_root / verification_name):
            report_verification_created += 1
      if report_print_created or report_verification_created:
         any_success = True
         messages.append(
            f"Merged report cards for {len(report_documents)} school(s)."
         )
      else:
         messages.append("No report cards found to merge.")
   else:
      messages.append("No report cards found to merge.")

   if id_documents:
      id_verification_created = 0
      for doc in id_documents:
         verification_name = (
            f"{date_prefix}_{doc.label}_ID_Card_verify.pdf"
         )
         if _create_verification_pdf(doc.pdfs, verification_root / verification_name):
            id_verification_created += 1
      if id_verification_created:
         any_success = True
         messages.append(
            f"Merged ID cards for {len(id_documents)} school(s)."
         )
      else:
         messages.append("No ID cards found to merge.")
   else:
      messages.append("No ID cards found to merge.")

   status_color = "green" if any_success else "red"
   set_status_message(" ".join(messages), status_color)


def merge_cover_pages() -> None:
   _thread.start_new_thread(_merge_cover_pages_worker, tuple())


# root window title and dimension
root.title("Processing UI")
# Set geometry(widthxheight)
root.geometry('520x680')

root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(6, weight=1)


#_thread.start_new_thread(printOut, (0,))
# adding Entry Field
button=tk.Button(root, text="Open Excel Sheet", command=open_win_diag)
button.grid(column=0, row=0, columnspan=2, pady=(10,5))

id_button=tk.Button(root, text="Open ID Card Sheet", command=open_id_card_diag)
id_button.grid(column=0, row=1, columnspan=2, pady=(0,5))

report_button=tk.Button(root, text="Open Report Card Sheet", command=open_report_card_diag)
report_button.grid(column=0, row=2, columnspan=2, pady=(0,5))

kid_label = tk.Label(root, text = "Kid Index (optional):")
kid_label.grid(column=0, row=3, sticky="w", padx=5)
kid_index_entry = tk.Entry(root, width=20)
kid_index_entry.grid(column=1, row=3, sticky="ew", padx=5)

status_label = tk.Label(root, fg = "red", text = "")
status_label.grid(column=0, row=4, columnspan=2, sticky="w", padx=5, pady=(5,0))

school_label = tk.Label(root, text="Select schools to process:")
school_label.grid(column=0, row=5, columnspan=2, sticky="w", padx=5, pady=(10,0))

school_frame = tk.Frame(root, bd=1, relief="sunken")
school_frame.grid(column=0, row=6, columnspan=2, sticky="nsew", padx=5, pady=(0,10))

school_canvas = tk.Canvas(school_frame, highlightthickness=0)
school_canvas.pack(side="left", fill="both", expand=True)
school_scrollbar = tk.Scrollbar(school_frame, orient="vertical", command=school_canvas.yview)
school_scrollbar.pack(side="right", fill="y")
school_canvas.configure(yscrollcommand=school_scrollbar.set)

schools_inner_frame = tk.Frame(school_canvas)
school_canvas.create_window((0,0), window=schools_inner_frame, anchor="nw")

def _configure_schools_frame(event):
    school_canvas.configure(scrollregion=school_canvas.bbox("all"))

schools_inner_frame.bind("<Configure>", _configure_schools_frame)

options = ["440X290", "380X255", "420X290", "420X297", "297X210"]
selected_size = tk.StringVar()

size_label = tk.Label(root, text="Select a size from dropdown")
size_label.grid(column=0, row=7, sticky="w", padx=5)

size_dropdown = tk.OptionMenu(root, selected_size, *options)
size_dropdown.grid(column=1, row=7, sticky="ew", padx=5)

size_info_label = tk.Label(root, fg = "blue", text = "")
size_info_label.grid(column=0, row=8, columnspan=2, sticky="w", padx=5)


def display_size(*args):
    size = selected_size.get()
    if size_info_label is not None:
        size_info_label.configure(text=f"Selected Size: {size} mm")
    width, height = map(int, size.split('X'))
    global Input
    Input=(width*2.83465,height*2.83465)


selected_size.set("420X290")
selected_size.trace_add("write", display_size)


def display_scale(*args):

   scale=selected_scale.get()
   global page_scale
   page_scale=int(scale)
   if scale_info_label is not None:
      scale_info_label.configure(text=f"Selected scale is {scale}%")


scale_label=tk.Label(root,text="Select a scale from dropdown in %")
scale_label.grid(column=0, row=9, sticky="w", padx=5)
scale_options=[50,60,70,80,90,100]
selected_scale=tk.StringVar()
selected_scale.trace_add("write", display_scale)
selected_scale.set(100)

scale_dropdown = tk.OptionMenu(root, selected_scale, *scale_options)
scale_dropdown.grid(column=1, row=9, sticky="ew", padx=5)

scale_info_label=tk.Label(root,fg="red",text="")
scale_info_label.grid(column=0,row=10,columnspan=2,sticky="w",padx=5)

display_scale()


button3=Button(root, text="DONE", command=windowDialog)
button3.grid(column=0, row=11, columnspan=2, pady=(10,5))


CheckVar1 = IntVar()
C1 = tk.Checkbutton(root, text = "SKIP FORM", variable = CheckVar1, \
                 onvalue = 1, offvalue = 0, height=2, \
                 width = 20)
C1.grid(column =0, row =12, columnspan=2, sticky="w", padx=5)


CheckVar2 = IntVar()
C2 = tk.Checkbutton(root, text = "OLD FORM", variable = CheckVar2, \
                 onvalue = 1, offvalue = 0, height=2, \
                 width = 20)
C2.grid(column =0, row =13, columnspan=2, sticky="w", padx=5)


checkVar3=tk.IntVar(value=0)

cv_page_button=tk.Checkbutton(root,var=checkVar3,text="Cover Page",height=2)
cv_page_button.grid(row=14, column=0, columnspan=2, sticky="w", padx=5)
checkVar4=tk.IntVar(value=0)



Ip_button=tk.Checkbutton(root,var=checkVar4,text="Inner Pages",height=2)
Ip_button.grid(row=15, column=0, columnspan=2, sticky="w", padx=5)

checkVar5=tk.IntVar(value=0)

id_cards_button=tk.Checkbutton(root, var=checkVar5, text="ID Cards", height=2)
id_cards_button.grid(row=16, column=0, columnspan=2, sticky="w", padx=5)

checkVar6=tk.IntVar(value=0)

report_cards_button=tk.Checkbutton(root, var=checkVar6, text="Report Cards", height=2)
report_cards_button.grid(row=17, column=0, columnspan=2, sticky="w", padx=5)

merge_button = tk.Button(root, text="Merge PDF", command=merge_cover_pages)
merge_button.grid(column=0, row=18, columnspan=2, pady=(10,5))

if __name__=="__main__":

   root.mainloop()

