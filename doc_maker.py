from xml.dom.minidom import parse as p
import os,shutil
import subprocess
from pathlib import Path
import pandas as pd
#parse your XML-document

def callInkscape(infile, outfile, timeout = 10, counter = 1,old = 0):
    try:
        # Do not override the DPI when exporting PDFs so the output canvas
        # matches the source SVG document size. Using ``--export-area-page``
        # ensures the exported PDF keeps the original SVG canvas dimensions,
        # which is important for report card front/back alignment.
        command = [
            'inkscape',
            infile,
            '--export-type=pdf',
            '--export-filename=' + outfile,
            '--export-area-page',
        ]
        # ``old`` is retained for backwards compatibility. We always
        # preserve the full page bounds so the canvas size in the exported
        # PDF matches the SVG template.
        p=subprocess.run(command, timeout = timeout)
        if p.returncode != 0:
            raise subprocess.CalledProcessError(p.returncode, command)
    except Exception as e:
        print(e)

        counter-=1

        if counter == 0:
            raise ValueError('TIMEOUT EXIRATION')
        
        callInkscape(infile, outfile, 20, counter - 1, old)

def personalize(tuple, subject, old = 0,sticker=False):
    print(tuple["first_name"])

    if os.path.exists(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL BOOKS FORM\PER"+"\\"+ subject) == False:
        pdfFolder = "PDFS" + "/" + subject + "/"  + str(tuple["book_id"]).zfill(3)+ "/"
        os.makedirs(pdfFolder, exist_ok=True)
        return

    svgFolder = "SVGS" + "/" + subject + "/"  + str(tuple["book_id"]).zfill(3)+ "/"
    pdfFolder = "PDFS"  + "/" + subject + "/"  + str(tuple["book_id"]).zfill(3)+ "/"
    store="store"
    
    if os.path.isdir(svgFolder):
        shutil.rmtree(svgFolder)

    if os.path.isdir(pdfFolder):
        shutil.rmtree(pdfFolder)
        

    os.makedirs(svgFolder, exist_ok=True)
    os.makedirs(pdfFolder, exist_ok=True)
    os.makedirs(store,exist_ok=True)
    print(svgFolder)
    print(pdfFolder)

    fname = ''
    fphoto = ''
    mname = ''
    mphoto = ''

    
    if (pd.isna(tuple["guardian_1_type"]) == False) and int(tuple['guardian_1_type']) == 0:
        fname = tuple['guardian_1_name']
        if (pd.isna(tuple["guardian_1_image"]) == False):
            fphoto = tuple["guardian_1_id"]
    elif (pd.isna(tuple["guardian_1_type"]) == False) and  int(tuple['guardian_1_type']) == 1:
        mname = tuple['guardian_1_name']
        if (pd.isna(tuple["guardian_1_image"]) == False):
            mphoto = tuple["guardian_1_id"]
    
    if (pd.isna(tuple["guardian_2_type"]) == False) and int(tuple['guardian_2_type']) == 0:
        fname = tuple['guardian_2_name']
        if (pd.isna(tuple["guardian_2_image"]) == False):
            fphoto = tuple["guardian_2_id"]
    elif (pd.isna(tuple["guardian_2_type"]) == False) and int(tuple['guardian_2_type']) == 1:
        mname = tuple['guardian_2_name']
        if (pd.isna(tuple["guardian_2_image"]) == False):
            mphoto = tuple["guardian_2_id"]
    
    photoFolder=r"\\pixartnas\home\INTERNAL_PROCESSING\ALL_PHOTOS"+"\\"+str(tuple["school_id"])

    shutil.copyfile(photoFolder + "\\"+"FULL"+"\\" + str(tuple["user_id"]) + '.png', 'store/' + str(tuple["user_id"]) + '.png')
    
    if mphoto != '':
        shutil.copyfile(photoFolder + "\\"+"PARTIAL"+"\\" + str(mphoto) + '.png', 'store/' + str(mphoto) + '.png')
    if fphoto != '':
        shutil.copyfile(photoFolder + "\\"+"PARTIAL"+"\\" + str(fphoto) + '.png', 'store/' + str(fphoto) + '.png')
    
    # if tuple["guardian_1_image"]!=None and tuple["guardian_1_image"]!="":
    #     shutil.copyfile(photoFolder + "\\"+"PARTIAL"+"\\" + str(tuple["guardian_1_id"]) + '.png', 'store/' + str(tuple["guardian_1_id"]) + '.png')
    # if tuple["guardian_2_image"]!=None and tuple["guardian_2_image"]!="":
    #      shutil.copyfile(photoFolder + "\\"+"PARTIAL"+"\\" + str(tuple["guardian_2_id"])  + '.png', 'store/' + str(tuple["guardian_2_id"]) + '.png')
        
    for file in os.listdir(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL BOOKS FORM\PER"+"\\" + subject):
        
        name, ext = os.path.splitext(file)
        
        if ext != '.svg':
            shutil.copyfile(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL BOOKS FORM\PER"+"\\" + subject + "\\" + file, svgFolder + "/" + file)

    for file in os.listdir(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL BOOKS FORM\PER"+"\\" + subject):
        
        name, ext = os.path.splitext(file)
        
        if ext != '.svg':
            continue
        
        cmmn_doc = p(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL BOOKS FORM\PER"+"\\" + subject + "\\" + file)

        notelist = cmmn_doc.getElementsByTagName("g")
        
        def find_element(id):
            definitions = {}
            i=0
            for i in range(len(notelist)):
                if notelist[i].getAttribute("id").lower() in id:
                    definitions[notelist[i].getAttribute("id").lower()] = notelist[i] #(or whatever you want to do)

            return definitions

        allLayers = find_element(["head", "name", "grade", "gender", "dob", "fphoto", "mphoto", "mname", "fname"])

        try:
            allLayers["head"]
            heads = allLayers["head"].getElementsByTagName("image")

            for head in heads:
                if head.getAttribute('xlink:href') != None:
                    print("changingHead")
                    head.setAttribute("xlink:href", '../../../store/' + str(tuple["user_id"]) + '.png' )

        except KeyError:
            pass
        except Exception as e:
            print(e)

        if fphoto != '':
            try:
                allLayers["fphoto"]
                heads = allLayers["fphoto"].getElementsByTagName("image")

                for head in heads:
                    if head.getAttribute('xlink:href') != None:
                        print("changingHead")
                        head.setAttribute("xlink:href", '../../../store/' + str(fphoto) + '.png' )

            except KeyError:
                pass
            except Exception as e:
                print(e)

        if mphoto != '':
            try:
                allLayers["mphoto"]
                heads = allLayers["mphoto"].getElementsByTagName("image")

                for head in heads:
                    if head.getAttribute('xlink:href') != None:
                        print("changingHead")
                        head.setAttribute("xlink:href", '../../../store/' + str(mphoto) + '.png' )

            except KeyError:
                pass
            except Exception as e:
                print(e)


        try:
            txts = allLayers["name"].getElementsByTagName("text")
            for txt in txts:
                
                
                txt.firstChild.data = (tuple["first_name"]+" "+tuple["last_name"]).title()
                
        except KeyError:
            pass
        except Exception as e:
            print(e)

        if mname != '':
            try:
                txts = allLayers["mname"].getElementsByTagName("text")
                for txt in txts:
                    txt.firstChild.data = mname.title()
                    
            except KeyError:
                pass
            except Exception as e:
                print(e)

        if fname != '':
            try:
                txts = allLayers["fname"].getElementsByTagName("text")
                for txt in txts:
                    txt.firstChild.data = fname.title()
                    
            except KeyError:
                pass
            except Exception as e:
                print(e)
         
        try:
            txts = allLayers["gender"].getElementsByTagName("text")

            for txt in txts:
                txt.firstChild.data = "boy" if tuple["gender"]=="male" else "girl"
                
        except KeyError:
            pass
        except Exception as e:
            print(e)

        try:
            txts = allLayers["grade"].getElementsByTagName("text")

            for txt in txts:
                txt.firstChild.data = tuple["class_name"]
                
        except KeyError:
            pass
        except Exception as e:
            print(e)


        try:
            txts = allLayers["dob"].getElementsByTagName("text")

            for txt in txts:
                txt.firstChild.data = tuple["date_of_birth"]
                
        except KeyError:
            pass
        except Exception as e:
            print(e)
       
            
        if sticker==True:
            
            svg = cmmn_doc.getElementsByTagName("svg")[0]  # Ensure the first <svg> element is accessed
            view_box = svg.getAttribute("viewBox")
            svg.setAttribute("width","410mm")
            svg.setAttribute("height","265mm")
            values = view_box.split()
            view_box_width = values[2]
            view_box_height = values[3]
            canvas_width=view_box_width
            canvas_height=view_box_height
            book_id_x=float(canvas_width)*0.02
            book_id_y=float(canvas_height)*0.02

            school_id_x=float(canvas_width)*0.4
            school_id_y=float(canvas_height)*0.02
            school_name=cmmn_doc.createElement("text")
            school_name.setAttribute("font-size", "15") 
            school_name.setAttribute("fill", "black")  # Ensure the text color is visible
            school_name.setAttribute("x", str(school_id_x))  # Center horizontally
            school_name.setAttribute("y", str(school_id_y))     
            school_name.setAttribute("text-anchor", "end")
            school_name.setAttribute("dominant-baseline", "middle")
            school_name.setAttribute("font-family", "Arial")
            school_name.appendChild(cmmn_doc.createTextNode(str(tuple["school_name"]).zfill(3)))
            cmmn_doc.documentElement.appendChild(school_name)

            new_text = cmmn_doc.createElement("text")
            new_text.setAttribute("font-size", "15") 
            new_text.setAttribute("fill", "black")  # Ensure the text color is visible
            new_text.setAttribute("x", str(book_id_x))  # Center horizontally
            new_text.setAttribute("y", str(book_id_y))     
            new_text.setAttribute("text-anchor", "middle")
            new_text.setAttribute("dominant-baseline", "middle")
            new_text.setAttribute("font-family", "Arial")
            
            new_text.appendChild(cmmn_doc.createTextNode(str(tuple["book_id"]).zfill(3)))
            cmmn_doc.documentElement.appendChild(new_text)

        open(svgFolder + file, "w", encoding="utf-8").write(cmmn_doc.toprettyxml())
        print(svgFolder)
        callInkscape (svgFolder + file, pdfFolder + file, 10, 4, old)