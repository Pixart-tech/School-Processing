import math
from xml.dom.minidom import parse as p
from PIL import ImageFont
import re
import subprocess
sticker_path=None
cmmn_doc=0

import os,shutil,sys


seperator_rect=None
def set_font_family(text_element, font_family,font_weight):
    """Updates the font size in the style attribute of a <text> element."""
    style = text_element.getAttribute("style")
    new_style = ""

    # Modify the existing font-family,fontweight or add if not present
    font_family_found = False
    font_weight_found = False
    
    for rule in style.split(";"):
        if "font-family" in rule:
            new_style += f"font-family:{font_family}; "  # Update font size
            font_family_found = True
        elif "font-weight" in rule:
            new_style += f"font-weight:{font_weight}; "  # Update font size
            font_weight_found = True
        else:   
            new_style += rule + "; " if rule.strip() else ""

    # If font-size wasn't found, add it
    if not font_family_found:
        new_style += f"font-size:{font_family};"
    if not font_weight_found:
        new_style += f"font-weight:{font_weight};"
        
    text_element.setAttribute("style", new_style.strip())


def set_font_size(text_element, new_size):
    """Updates the font size in the style attribute of a <text> element."""
    style = text_element.getAttribute("style")
    new_style = ""
    
    # Modify the existing font-size or add if not present
    font_size_found = False
    for rule in style.split(";"):
        if "font-size" in rule:
            new_style += f"font-size:{new_size}px; "  # Update font size
            font_size_found = True
        else:
            new_style += rule + "; " if rule.strip() else ""

    # If font-size wasn't found, add it
    new_style += f"font-size:{new_size}px;"
        
    text_element.setAttribute("style", new_style.strip())

def fontSize (elem):
    style = elem.getAttribute("style")
    # Find the font-size inside the style attribute
    font_size = 38
    for rule in style.split(";"):
        if "font-size" in rule:
            font_size = rule.split(":")[-1].strip().replace("px", "")  # Extract number and remove 'px'
            break
    return font_size

def fontFam (elem):
    style = elem.getAttribute("style")
    # Find the font-size inside the style attribute
    font_size = ''
    for rule in style.split(";"):
        if "font-family" in rule:
            font_size = rule.split(":")[-1].strip()  # Extract number and remove 'px'
            break
    return font_size


#parse your XML-document
def get_next_element_sibling(node):
    """Return the next sibling that is an element node (skips text, comments, etc.)."""
    sibling = node.nextSibling
    while sibling and sibling.nodeType != sibling.ELEMENT_NODE:
        sibling = sibling.nextSibling
    return sibling  

def get_prev_element_sibling(node):
    """Return the next sibling that is an element node (skips text, comments, etc.)."""
    sibling = node.previousSibling
    while sibling and sibling.nodeType != sibling.ELEMENT_NODE:
        sibling = sibling.previousSibling
    return sibling


def get_sibling_rect(element):
    """
    Starting from the given element, find the next sibling that is a <rect> element.
    Returns the <rect> element if found, else None.
    """
    sibling = get_next_element_sibling(element)
    while sibling:
        if sibling.tagName == "rect":
            return sibling
        sibling = get_next_element_sibling(sibling)
        
    sibling = get_prev_element_sibling(element)
    while sibling:
        if sibling.tagName == "rect":
            return sibling
        sibling = get_prev_element_sibling(sibling)
        
        
    return None



   
   
def callInkscape(infile, outfile, timeout = 10, counter = 1,old = 0):
    try:
        if old == 0:
            print("Iam running")
            p=subprocess.run(['inkscape',infile,'--export-type=pdf','--export-filename=' + outfile, "--export-dpi=400", "--export-area-page"], timeout = timeout)
        else:    
            print("Iam also running")
            p=subprocess.run(['inkscape',infile,'--export-type=pdf','--export-filename=' + outfile, "--export-dpi=400", "--export-area-drawing"], timeout = timeout)
    except Exception as e:
        print(e)
        
        counter-=1
        
        if counter == 0:
            raise ValueError('TIMEOUT EXIRATION')
        
        callInkscape(infile, outfile, 20, counter - 1, old)
# Ensure 'infile' and 'outfile' are defined before calling this function
def convert_to_mm(value):
    match = re.match(r'([0-9.]+)([a-z%]*)', value.strip())
    if not match:
        return None

    number, unit = match.groups()
    number = float(number)

    unit = unit.lower()
    if unit in ('pt', ''):
        return number/2.835
    elif unit == 'px':
        return number * 0.2645833333   # Assuming 96 DPI
    elif unit == 'in':
        return number * 25.4
    elif unit == 'cm':
        return number * 10
    elif unit == 'mm':
        return number
    else:
        return None  # Unknown or unsupported unit
    
def convert_to_points(value):
    match = re.match(r'([0-9.]+)([a-z%]*)', value.strip())
    if not match:
        return None

    number, unit = match.groups()
    number = float(number)

    unit = unit.lower()
    if unit in ('pt', ''):
        return number
    elif unit == 'px':
        return number * 0.75  # Assuming 96 DPI
    elif unit == 'in':
        return number * 72
    elif unit == 'cm':
        return number * 28.3465
    elif unit == 'mm':
        return number * 2.83465
    else:
        return None  # Unknown or unsupported unit


def callInkscape_png(infile, outfile, timeout = 10, counter = 1,old = 0):
    try:
        if old == 0:
            p=subprocess.run(['inkscape',infile,'--export-filename=' + outfile, "--export-dpi=100", "--export-area-page"], timeout = timeout)
        else:    
            p=subprocess.run(['inkscape',infile,'--export-filename=' + outfile, "--export-dpi=100", "--export-area-drawing"], timeout = timeout)
    except Exception as e:
        print(e)
        
        counter-=1
        
        if counter == 0:
            raise ValueError('TIMEOUT EXIRATION')
        
        callInkscape_png(infile, outfile, 20, counter - 1, old)


def personalize(outer_code,photoFolder ,id,school_color_code1,school_color_code2,grade_colour_code,kid_color_code1,kid_color_code2,name,bookid,tuple):
       
        cmmn_doc =p( r"\\pixartnas\home\INTERNAL_PROCESSING\SCHOOLCOVERS"+"\\"+str(outer_code)[0:3]+"\\"+str(outer_code)+".svg")
        
        shutil.copyfile(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL_PHOTOS"+"\\"+str(id)+"\\"+"FULL"+"\\"+photoFolder+".png", 'store/' + str(photoFolder)+".png" )
        
        
        
        notelist = cmmn_doc.getElementsByTagName("g")                                                       
    
        def find_element(id):
            definitions = {}
            i=0
            for i in range(len(notelist)):
                if notelist[i].getAttribute("id").lower() in id:
                    definitions[notelist[i].getAttribute("id").lower()] = notelist[i] 
            
            return definitions
        
        allLayers = find_element(["head", "name","rectLayer"])
        
       
        try:
            allLayers["head"]
            heads = allLayers["head"].getElementsByTagName("image")

            for head in heads:
                if head.getAttribute('xlink:href') != None:
                    print("changingHead")
                    head.setAttribute("xlink:href",  '../../store/' + photoFolder + '.png')

        except KeyError:
            pass
        except Exception as e:
            print(e)
        try:
            
            txts = allLayers["name"].getElementsByTagName("text")
            rects = allLayers["name"].getElementsByTagName("rect")
            idx = 0
            
            for txt in txts:
                
                try:
                    rect=rects[idx]
                except IndexError:
                    rect = None

                idx = idx + 1

                if rect != None:
                    rx = float(rect.getAttribute('x'))
                    ry = float(rect.getAttribute('y'))
                    width = float(rect.getAttribute('width'))
                    height = float(rect.getAttribute('height'))
                    # print(width)
                    # Calculate the center of the rect
                    center_x = rx + width / 2
                    center_y = ry + height / 2

                    # Set text attributes for centering
                    txt.setAttribute('text-anchor', 'middle')
                    txt.setAttribute('x', str(center_x))
                    txt.setAttribute('y', str(center_y))
                    txt.setAttribute('dominant-baseline', 'middle')
                    
                    txt.setAttribute('transform', '')
                    
                txt.firstChild.data = name.title()
                #set_font_family(txt,"Playpen Sans",500)
                
                if len(name)>12:
                    font_size = math.floor(float(fontSize(txt)))
                    font_fam = fontFam(txt)
                    
                    font_file = "PlaypenSans-Medium.ttf"

                    if 'Marvin' in font_fam:
                         font_file = "Marvin.ttf"

                    if font_size==None:
                        font_size=38
                    
                    font = ImageFont.truetype(font_file, int(font_size))
                    x,y,x1,y1 = font.getbbox(name.title())
                    text_width = x1-x
                    while int(text_width) > int(width) and int(font_size) > 1:
                        font_size -= 0.2
                        font = ImageFont.truetype(font_file, font_size)
                        x,y,x1,y1 = font.getbbox(name.title())
                        text_width = x1-x
                    print(font_size)
                    set_font_size(txt, font_size)
             
        except KeyError:
            pass
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

        svg = cmmn_doc.getElementsByTagName("svg")[0]  # Ensure the first <svg> element is accessed
        view_box = svg.getAttribute("viewBox")
        values = view_box.split()
        view_box_width = values[2]
        view_box_height = values[3]

        if not str(outer_code).endswith("b"): 
                   
            rect_width="20"

            rect_height=float(view_box_height)
            rect_y=15
            rect_x=float(view_box_width)/2-float(rect_width)/2
            
            diff = 25
          
            # text_x=float(rect_x)+float(rect_width)*0.5
            text_x=float(rect_x)-20
            text_y=float(rect_y)+rect_height*0.5
            
            school_rect1=cmmn_doc.createElement("rect")
            school_rect1.setAttribute("x", str(rect_x))
            school_rect1.setAttribute("y", str(rect_y))  
            school_rect1.setAttribute("width",str(20))
            school_rect1.setAttribute("height",str(30)) 
            school_rect1.setAttribute("fill", str(school_color_code1)) 
            cmmn_doc.documentElement.appendChild(school_rect1)
            
            school_rect2=cmmn_doc.createElement("rect")
            school_rect2.setAttribute("x", str(rect_x))  
            school_rect2.setAttribute("y", str(rect_y+60))  
            school_rect2.setAttribute("width",str(20))
            school_rect2.setAttribute("height",str(30))
            school_rect2.setAttribute("fill", str(school_color_code2)) 
            cmmn_doc.documentElement.appendChild(school_rect2)
            
            # grade_rect1=cmmn_doc.createElement("rect")
            # grade_rect1.setAttribute("x", str(rect_x))
            # grade_rect1.setAttribute("y", str(rect_y+diff*2)  ) 
            # grade_rect1.setAttribute("width",str(20))
            # grade_rect1.setAttribute("height",str(30))
            # grade_rect1.setAttribute("fill", grade_colour_code) 
            # cmmn_doc.documentElement.appendChild(grade_rect1)
            
           
            
            kid_color_rect1=cmmn_doc.createElement("rect")
            kid_color_rect1.setAttribute("x", str(rect_x))
            kid_color_rect1.setAttribute("y", str(rect_height-97))   
            kid_color_rect1.setAttribute("width",str(20))
            kid_color_rect1.setAttribute("height",str(30))
            kid_color_rect1.setAttribute("fill", kid_color_code1) 
            
            cmmn_doc.documentElement.appendChild(kid_color_rect1)
            
            kid_color_rect2=cmmn_doc.createElement("rect")
            kid_color_rect2.setAttribute("x", str(rect_x))
            kid_color_rect2.setAttribute("y", str(rect_height -45))  
            kid_color_rect2.setAttribute("width",str(20))
            kid_color_rect2.setAttribute("height",str(30))
            kid_color_rect2.setAttribute("fill", kid_color_code2) 
            cmmn_doc.documentElement.appendChild(kid_color_rect2)  
            
      
            grade_start_bound=0.3*rect_height
            grade_end_bound=rect_height  - 115
            
                             
            top_line=cmmn_doc.createElement("line")
            top_line.setAttribute("x1",str(rect_x-5))
            top_line.setAttribute("x2",str(rect_x-5+float(rect_width)+10))
            top_line.setAttribute("stroke-width",str(1))
            top_line.setAttribute("stroke", "black")
            top_line.setAttribute("y1",str(grade_start_bound))
            top_line.setAttribute("y2",str(grade_start_bound))
            cmmn_doc.documentElement.appendChild(top_line)
            
            bottom_line=cmmn_doc.createElement("line")
            bottom_line.setAttribute("x1",str(rect_x-5))
            bottom_line.setAttribute("x2",str(rect_x-5+float(rect_width)+10))
            bottom_line.setAttribute("y1",str(grade_end_bound))
            bottom_line.setAttribute("y2",str(grade_end_bound))
            bottom_line.setAttribute("x2",str(rect_x+25))
            bottom_line.setAttribute("stroke-width",str(1))
            bottom_line.setAttribute("stroke", "black")
            cmmn_doc.documentElement.appendChild(bottom_line)
             
            
            
            
            
            
            ht = (grade_end_bound - grade_start_bound)/int(tuple['numsub'])
            y_pos = grade_start_bound + (ht * (int(tuple['subidx']) -1))
            
            
            grade_rect=cmmn_doc.createElement("rect")
            grade_rect.setAttribute("x",str(rect_x))
                
            grade_rect.setAttribute("y",str(y_pos))
            
            grade_rect.setAttribute("width",str(rect_width))
            grade_rect.setAttribute("height",str(ht))
            grade_rect.setAttribute("fill",grade_colour_code)
            cmmn_doc.documentElement.appendChild(grade_rect)
            
             
            new_text = cmmn_doc.createElement("text")
            new_text.setAttribute("font-size", "11") 
            new_text.setAttribute("fill", "black")  # Ensure the text color is visiblee
            new_text.setAttribute("x", str(text_x))  # Center horizontally
            new_text.setAttribute("y", str(text_y))     
            new_text.setAttribute("text-anchor", "middle")
            new_text.setAttribute("dominant-baseline", "middle")
            new_text.setAttribute("font-family", "Arial")
            
            new_text.setAttribute("transform", f"rotate(-90 {text_x} {text_y})") 
            
            
            new_text.appendChild(cmmn_doc.createTextNode(bookid))
            
            cmmn_doc.documentElement.appendChild(new_text)

            svg = cmmn_doc.getElementsByTagName("svg")[0]
                   
            svg.setAttribute("width", str(math.floor(convert_to_mm(view_box_width))) + "mm")
            svg.setAttribute("height", str(math.floor(convert_to_mm(view_box_height))) + "mm")
            
            open("Temp"+"/"+ str(outer_code)[:3]+"/"+bookid+".svg", "w", encoding="utf-8").write(cmmn_doc.toprettyxml())
            callInkscape("Temp"+"/"+ str(outer_code)[:3]+"/"+bookid+".svg","finalcovers"+"/"+bookid+".pdf",10,10,0)
            
            if int(tuple['subidx']) == 1:
                png_folder = os.path.join("Temp", str(outer_code)[:3], "PNG")
                if not os.path.isdir(png_folder):
                    os.makedirs(png_folder)
                callInkscape_png("Temp"+"/"+str(outer_code)[:3]+"/"+bookid+".svg","Temp"+"/"+str(outer_code)[:3]+"/"+"PNG"+"/"+bookid+".png",10,3,0)
        # callInkscape(school_code+"/"+in_code+code+".svg",school_code+"/"+"PDF"+"/"+in_code+code+".pdf")f")
        else:
            svg.setAttribute("width","300mm")
            svg.setAttribute("height","220mm")
            canvas_width=view_box_width
            canvas_height=view_box_height
            school_id_x=float(canvas_width)*0.35
            school_id_y=float(canvas_height)*0.97
            school_name=cmmn_doc.createElement("text")
            school_name.setAttribute("font-size", "15") 
            school_name.setAttribute("fill", "black")  # Ensure the text color is visible
            school_name.setAttribute("x", str(school_id_x))  # Center horizontally
            school_name.setAttribute("y", str(school_id_y))     
            school_name.setAttribute("text-anchor", "end")
            school_name.setAttribute("dominant-baseline", "middle")
            school_name.setAttribute("font-family", "Arial")
            school_name.appendChild(cmmn_doc.createTextNode(tuple["school_name"]))
            cmmn_doc.documentElement.appendChild(school_name)
            
            open("Temp"+"/"+ str(outer_code)[:3]+"/"+bookid+".svg", "w", encoding="utf-8").write(cmmn_doc.toprettyxml()) 
            if  not os.path.isdir("Temp"+"/"+ str(outer_code)[:3]+"/"+"Boxstickers"):
                os.makedirs("Temp"+"/"+ str(outer_code)[:3]+"/"+"Boxstickers")
                
            if str(outer_code).endswith("b"):
                callInkscape("Temp"+"/"+ str(outer_code)[:3]+"/"+bookid+".svg","Temp"+"/"+ str(outer_code)[:3]+"/"+"Boxstickers"+"/"+bookid+".pdf",10,4,0)
            
       
    
            

    

                            
    
    

        


                    
        
    
            
    
    # if cover_page==False: 
        
    #     svg = cmmn_doc.getElementsByTagName("svg")[0]
    #     svg.setAttribute("width", "410mm")
    #     svg.setAttribute("height", "265mm")
    #     open(school_code+"/"+in_code+code+".svg", "w", encoding="utf-8").write(cmmn_doc.toprettyxml())
    #     callInkscape(school_code+"/"+in_code+code+".svg",school_code+"/"+"sticker PDF"+"/"+in_code+code+".pdf",10,3,0)



