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

from PIL import Image
import os


import doc_maker,dc
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


if os.path.isdir("STICKERS")==False:
   os.makedirs("STICKERS")

#DOCDRIVER = DOC["DriverData"]
print_jobs = []

SEEKPOS = 9938
Input=(420*2.83465,(290*2.83465))

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
   
   

def open_win_diag():
   # Create a dialog box
   global file
   file=filedialog.askopenfilename(initialdir="E:/winapp")



def windowDialog():
   _thread.start_new_thread(make, (0,))


def convert_cmyk_to_rgb(input_path, output_path):
    with Image.open(input_path) as img:
        if img.mode == "CMYK":
            img = img.convert("RGB")
        img.save(output_path)

def make(dum):
   spine=0
   global file,folder
   if file == "" :
      lbl2.configure(text = "ERROR SELECT EXCEL \nFILE AND FOLDER \nWTH PHOTOS" )
      return
   elif size=="":
      lbl2.configure(text="ERROR SELECT A BOOKLET SIZE ")
      return
         
   
   if os.path.isdir("FINAL BINDERS") == False:
      os.makedirs("FINAL BINDERS", exist_ok=True)
   idx = txt.get()

   if idx .startswith('U'):
      
      idx1=idx[1:].split(',')
      list_idx=[ f.strip() for f in idx1 ]
      
   else:
      idx1 = idx.split(",")
      fromIdx = int(idx1[0].strip())
      toIdx = int(idx1[1].strip())
   
  
   
   df = pd.read_excel(file, header=0)
   data = df.to_dict('index')
   subjectIDX = {}
   prev = ""
   
   
   colorcodes = ['#ff0000', '#00ff00', '#8F7C00' ,'#0000ff', '#993F00'  ,'#ffff00', '#00ffff', '#ff00ff', '#000000', '#888888', '#234567']
        
         
   
   if checkVar3.get()==1:
      
      if not os.path.isdir("store"):
         os.makedirs("store")
 
      for key,item in data.items():
         if pd.isna(item["last_name"]):
            item["last_name"] = ''
            print(item["book_id"])
         if idx.startswith('U'):
            if item["user_id"] not in list_idx:
               continue
         
         elif int(item["book_id"]) < fromIdx or int(item["book_id"]) > toIdx:
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
         
         dc.personalize(item["outer_code"], photo_name , item["school_id"],school_color_code1,school_color_code2,grade_colour_code,kid_color_code1,kid_color_code2,item["first_name"]+" "+item["last_name"],str(item["book_id"]).zfill(3),item)      
   
   if checkVar4.get()==1:
      for key,item in data.items():
         if idx.startswith('U'):
            if item["user_id"] not in list_idx:
               continue
         elif int(item["book_id"]) < fromIdx or int(item["book_id"]) > toIdx:
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

 
# root window title and dimension
root.title("Processing UI")
# Set geometry(widthxheight) 
root.geometry('350x350')



#_thread.start_new_thread(printOut, (0,))
# adding Entry Field
button=tk.Button(root, text="Open Excel Sheet", command=open_win_diag)
button.grid(column =1, row =0)

lbl = tk.Label(root, text = "Book Index From,TO:")
lbl.grid(column =0, row =2)
txt = tk.Entry(root, width=10)
txt.grid(column =1, row =2)
size_label = tk.Label(root, text="Select a size from dropdown")
size_label.grid(column =0, row =4)
options = ["440X290", "380X255", "420X290", "420X297", "297X210"]
selected_size = tk.StringVar()
lbl2 = tk.Label(root, fg = "red", text = "")
lbl2.grid(column =1, row =5)
def display_size(*args):
    size = selected_size.get()
    lbl2.configure(text=f"Selected Size: {size} mm")
    width, height = map(int, size.split('X'))
    global Input
    Input=(width*2.83465,height*2.83465) 
selected_size.set("420X290") 
selected_size.trace_add("write", display_size)
def display_scale(*args):
   
   scale=selected_scale.get()
   global page_scale
   page_scale=int(scale)
   lbl3.configure(text=f"selected scale is {scale}")

 
size_dropdown = tk.OptionMenu(root, selected_size, *options)
size_dropdown.grid(column =1, row =4)
size=selected_size.get()
lbl3=tk.Label(root,fg="red",text="")
lbl3.grid(column=1,row=7)

scale=tk.Label(root,text="Select a scale from dropdown in %")
scale.grid(column =0, row =6)
scale_options=[50,60,70,80,90,100]
selected_scale=tk.StringVar()
selected_scale.trace_add("write", display_scale)
selected_scale .set(100)  

scale_dropdown = tk.OptionMenu(root, selected_scale, *scale_options)
scale_dropdown.grid(column =1, row =6)

button3=Button(root, text="DONE", command=windowDialog)
button3.grid(column =1, row =3,pady=10)



checkVar3=tk.IntVar(value=0)

cv_page_button=tk.Checkbutton(root,var=checkVar3,text="Cover Page",height=5)
cv_page_button.grid(row=10, column=1)
checkVar4=tk.IntVar(value=0)



Ip_button=tk.Checkbutton(root,var=checkVar4,text="Inner Pages",height=5)
Ip_button.grid(row=11, column=1)




CheckVar1 = IntVar()
C1 = tk.Checkbutton(root, text = "SKIP FORM", variable = CheckVar1, \
                 onvalue = 1, offvalue = 0, height=5, \
                 width = 20)
C1.grid(column =1, row =8)


CheckVar2 = IntVar()
C2 = tk.Checkbutton(root, text = "OLD FORM", variable = CheckVar2, \
                 onvalue = 1, offvalue = 0, height=5, \
                 width = 20)
C2.grid(column =1, row =9)

if __name__=="__main__":
   
   root.mainloop()

