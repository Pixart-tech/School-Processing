import math
from xml.dom.minidom import parse as p
from PIL import ImageFont
import re
import subprocess
sticker_path=None
cmmn_doc=0

import os,shutil,sys


def _sanitize_for_path(value, fallback):
    """Return a filesystem-friendly string for folder/file names."""
    if value is None:
        value = ""
    value = str(value).strip()
    if not value:
        return fallback

    sanitized = re.sub(r"[^A-Za-z0-9_-]+", "_", value)
    sanitized = re.sub(r"_+", "_", sanitized).strip("_-")
    return sanitized or fallback


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


MIN_FONT_SIZE = 9.0


def _format_float(value):
    return ("{:.4f}".format(value)).rstrip("0").rstrip(".")


def _clear_text(element):
    while element.firstChild:
        element.removeChild(element.firstChild)


def _set_text_content(element, text):
    _clear_text(element)
    element.appendChild(element.ownerDocument.createTextNode(text))


def _resolve_font_file(element):
    font_family = fontFam(element)
    if font_family and "marvin" in font_family.lower():
        return "Marvin.ttf"
    return "PlaypenSans-Medium.ttf"


def _load_font(font_file, size):
    try:
        return ImageFont.truetype(font_file, max(1, int(math.floor(size))))
    except OSError:
        return None


def _measure_text_width(font, text):
    if font is None or not text:
        return 0.0
    left, _, right, _ = font.getbbox(text)
    return float(right - left)


def _split_text_into_two_lines(text):
    cleaned = text.strip()
    if not cleaned:
        return [""]

    words = cleaned.split()
    if len(words) <= 1:
        midpoint = max(1, len(cleaned) // 2)
        return [cleaned[:midpoint].strip(), cleaned[midpoint:].strip()]

    best_index = 1
    best_diff = float("inf")
    for index in range(1, len(words)):
        first_line = " ".join(words[:index])
        second_line = " ".join(words[index:])
        diff = abs(len(first_line) - len(second_line))
        if diff < best_diff:
            best_diff = diff
            best_index = index

    first = " ".join(words[:best_index]).strip()
    second = " ".join(words[best_index:]).strip()
    return [line for line in (first, second) if line]


def _set_two_line_centered_text(element, first_line, second_line, center_x, line_height=1.1):
    _clear_text(element)
    document = element.ownerDocument
    x_value = _format_float(center_x)
    half_step = line_height / 2.0

    first_tspan = document.createElement("tspan")
    first_tspan.setAttribute("x", x_value)
    first_tspan.setAttribute("dy", f"-{_format_float(half_step)}em")
    first_tspan.appendChild(document.createTextNode(first_line))
    element.appendChild(first_tspan)

    second_tspan = document.createElement("tspan")
    second_tspan.setAttribute("x", x_value)
    second_tspan.setAttribute("dy", f"{_format_float(line_height)}em")
    second_tspan.appendChild(document.createTextNode(second_line))
    element.appendChild(second_tspan)


def _fit_text_within_rect(element, rect, text, base_size=None, min_scale=0.7, max_scale=1.1):
    if rect is None:
        _set_text_content(element, text)
        return {
            "fits": True,
            "rect_width": None,
            "rect_height": None,
            "center_x": None,
            "center_y": None,
            "font_file": _resolve_font_file(element),
            "current_size": base_size,
            "base_size": base_size,
            "min_size": None,
            "max_size": None,
        }

    try:
        rect_x = float(rect.getAttribute('x')) if rect.hasAttribute('x') else 0.0
    except (TypeError, ValueError):
        rect_x = 0.0
    try:
        rect_y = float(rect.getAttribute('y')) if rect.hasAttribute('y') else 0.0
    except (TypeError, ValueError):
        rect_y = 0.0
    try:
        rect_width = float(rect.getAttribute('width')) if rect.hasAttribute('width') else 0.0
    except (TypeError, ValueError):
        rect_width = 0.0
    try:
        rect_height = float(rect.getAttribute('height')) if rect.hasAttribute('height') else 0.0
    except (TypeError, ValueError):
        rect_height = 0.0

    if rect_width <= 0:
        _set_text_content(element, text)
        return {
            "fits": True,
            "rect_width": rect_width,
            "rect_height": rect_height,
            "center_x": None,
            "center_y": None,
            "font_file": _resolve_font_file(element),
            "current_size": base_size,
            "base_size": base_size,
            "min_size": None,
            "max_size": None,
        }

    center_x = rect_x + rect_width / 2.0
    center_y = rect_y + (rect_height / 2.0 if rect_height else 0.0)

    element.setAttribute('text-anchor', 'middle')
    element.setAttribute('dominant-baseline', 'middle')
    element.setAttribute('x', _format_float(center_x))
    element.setAttribute('y', _format_float(center_y))
    if element.hasAttribute('transform'):
        element.removeAttribute('transform')

    style_size = fontSize(element)
    try:
        style_size = float(style_size)
    except (TypeError, ValueError):
        style_size = 38.0

    if base_size is None:
        base_size = style_size if style_size else 38.0

    if base_size <= 0:
        base_size = 1.0

    font_file = _resolve_font_file(element)

    scaled_min = base_size * min_scale
    if base_size >= MIN_FONT_SIZE:
        min_allowed = max(MIN_FONT_SIZE, scaled_min)
    else:
        min_allowed = scaled_min
    max_allowed = base_size * max_scale
    if max_allowed < base_size:
        max_allowed = base_size

    current_size = min(float(base_size), max_allowed)

    font = _load_font(font_file, current_size)
    if font is None:
        set_font_size(element, round(current_size, 2))
        _set_text_content(element, text)
        return {
            "fits": False,
            "rect_width": rect_width,
            "rect_height": rect_height,
            "center_x": center_x,
            "center_y": center_y,
            "font_file": font_file,
            "current_size": current_size,
            "base_size": base_size,
            "min_size": min_allowed,
            "max_size": max_allowed,
        }

    text_width = _measure_text_width(font, text)
    fits_width = text_width <= rect_width

    if text_width <= rect_width * 0.95:
        set_font_size(element, round(current_size, 2))
        _set_text_content(element, text)
        return {
            "fits": fits_width,
            "rect_width": rect_width,
            "rect_height": rect_height,
            "center_x": center_x,
            "center_y": center_y,
            "font_file": font_file,
            "current_size": current_size,
            "base_size": base_size,
            "min_size": min_allowed,
            "max_size": max_allowed,
        }

    while text_width > rect_width and current_size > min_allowed:
        new_size = max(current_size - 0.2, min_allowed)
        if math.isclose(new_size, current_size, rel_tol=1e-3, abs_tol=1e-3):
            current_size = new_size
            break
        current_size = new_size
        font = _load_font(font_file, current_size)
        if font is None:
            break
        text_width = _measure_text_width(font, text)

    set_font_size(element, round(current_size, 2))
    _set_text_content(element, text)

    return {
        "fits": text_width <= rect_width,
        "rect_width": rect_width,
        "rect_height": rect_height,
        "center_x": center_x,
        "center_y": center_y,
        "font_file": font_file,
        "current_size": current_size,
        "base_size": base_size,
        "min_size": min_allowed,
        "max_size": max_allowed,
    }


def _apply_two_line_layout(element, text, fit_result):
    rect_width = fit_result.get("rect_width")
    rect_height = fit_result.get("rect_height")
    center_x = fit_result.get("center_x")
    center_y = fit_result.get("center_y")
    font_file = fit_result.get("font_file")

    if not rect_width or rect_width <= 0 or center_x is None or center_y is None:
        return

    lines = _split_text_into_two_lines(text)
    if len(lines) < 2:
        return

    base_size = fit_result.get("base_size") or fit_result.get("current_size") or 38.0
    min_allowed = fit_result.get("min_size")
    max_allowed = fit_result.get("max_size")

    if min_allowed is None:
        scaled_min = base_size * 0.7
        if base_size >= MIN_FONT_SIZE:
            min_allowed = max(MIN_FONT_SIZE, scaled_min)
        else:
            min_allowed = scaled_min

    if max_allowed is None:
        max_allowed = max(base_size, base_size * 1.1)

    effective_size = min(max(base_size, min_allowed), max_allowed)
    line_spacing = 1.1
    vertical_limit = None
    if rect_height:
        vertical_limit = rect_height / (1.0 + line_spacing)
        effective_size = min(effective_size, max(min_allowed, vertical_limit))

    font = _load_font(font_file, effective_size) if font_file else None
    if font is not None:
        max_width = max(_measure_text_width(font, line) for line in lines)
        while (
            (rect_width and max_width > rect_width)
            or (vertical_limit is not None and effective_size > vertical_limit)
        ) and effective_size > min_allowed:
            new_size = max(effective_size - 0.2, min_allowed)
            if math.isclose(new_size, effective_size, rel_tol=1e-3, abs_tol=1e-3):
                effective_size = new_size
                break
            effective_size = new_size
            font = _load_font(font_file, effective_size) if font_file else None
            if font is None:
                break
            max_width = max(_measure_text_width(font, line) for line in lines)

    effective_size = min(max(effective_size, min_allowed), max_allowed)
    set_font_size(element, round(effective_size, 2))

    element.setAttribute('text-anchor', 'middle')
    element.setAttribute('dominant-baseline', 'middle')
    element.setAttribute('x', _format_float(center_x))
    element.setAttribute('y', _format_float(center_y))
    if element.hasAttribute('transform'):
        element.removeAttribute('transform')

    _set_two_line_centered_text(element, lines[0], lines[1], center_x, line_height=line_spacing)

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


def personalize(outer_code,photoFolder ,id,school_color_code1,school_color_code2,grade_colour_code,kid_color_code1,kid_color_code2,name,bookid,tuple,multiple_schools=False):
       
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
                    rect = rects[idx]
                except IndexError:
                    rect = None

                idx = idx + 1

                display_name = name.title()
                fit_result = _fit_text_within_rect(txt, rect, display_name)

                if not fit_result.get("fits"):
                    _apply_two_line_layout(txt, display_name, fit_result)

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

            school_name_raw = tuple.get("school_name", "")
            subject_name_raw = (
                tuple.get("subject_name")
                or tuple.get("subject")
                or tuple.get("subjectname")
                or ""
            )
            first_name_raw = tuple.get("first_name", "")
            last_name_raw = tuple.get("last_name", "")

            school_name_clean = _sanitize_for_path(school_name_raw, "School")
            subject_name_clean = _sanitize_for_path(subject_name_raw, "Subject")
            first_name_clean = _sanitize_for_path(first_name_raw, "First")
            last_name_clean = _sanitize_for_path(last_name_raw, "Last")

            output_dir = os.path.join("finalcovers", f"{str(outer_code)[:3]}_{school_name_clean}")
            os.makedirs(output_dir, exist_ok=True)

            output_name = (
                f"{school_name_clean}_{outer_code}_{subject_name_clean}_{first_name_clean}_{last_name_clean}.pdf"
            )

            output_path = os.path.join(output_dir, output_name)

            callInkscape("Temp"+"/"+ str(outer_code)[:3]+"/"+bookid+".svg",output_path,10,10,0)
            
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



