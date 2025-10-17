



import fitz  # PyMuPDF
from  reportlab import *
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import os
import sys
import math
import logging
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import os
import sys
import math
import logging
# def addBorder(sizeArt, page, mark="", endSticksLev=0, diff=False):
#     if mark == "" and endSticksLev == 0 and not diff:
#         return
    
    
#     c = canvas.Canvas("watermark.pdf")
#     c.setPageSize((sizeArt[0], sizeArt[1]))
    
#     if mark:
#         c.drawString(int(sizeArt[0] / 2), int(sizeArt[1] / 2), mark)
    
#     if diff:
#         c.setFillColor(colors.black)
#         c.rect(0, sizeArt[1] - 15, 70, 10, fill=1)
    
#     if endSticksLev > 0:
#         c.setFillColor(colors.black)
#         c.rect(0, sizeArt[1] + 40 - (endSticksLev * 50), 10, 10, fill=1)
    
#     c.save()
    
#     watermark = fitz.open("watermark.pdf")
#     page.show_pdf_page(page.rect, watermark, 0)
def iter_pages(pages, aspect):
    refinedaspect = 0.95
    for output_page in range(pages // 4):
        first = output_page * 2
        last = pages - 1 - first
        yield first, -90, 1 - ((1 - refinedaspect) / 2), 0.5
        yield last, -90, 1 - ((1 - refinedaspect) / 2), 0.5 - (refinedaspect) / 2
        yield first + 1, 90, (1 - refinedaspect) / 2, 0.5 + (refinedaspect) / 2
        yield last - 1, 90, (1 - refinedaspect) / 2, 0.5

def build_doc(in_file, out_file, name, pages, i,combined=None,page_scale=100, scale=0, addMarker=False, school = ''):
    print(page_scale)
    if pages % 4 != 0 or not os.path.exists(in_file):
        print(f"ERROR IN {name}")
        return
    
    sizchange = i
    
    if combined == None:
        out = fitz.open()
    else:
        out = combined
    
    for i, (p, r, x, y) in enumerate(iter_pages(pages, 1)):
        
        if i%2==0 :
            new_page = out.new_page(width=sizchange[0], height=sizchange[1])
            new_page.insert_text((sizchange[0]/2,sizchange[1]/4), school, rotate=90,fontsize=12, fontname="helvetica", color=(0, 0, 0))     
        
        if (i == 1 or (i == pages-1)) and addMarker:
            book_id=name
            new_page.insert_text((sizchange[0]/2,sizchange[1]/2), book_id, rotate=90,fontsize=20, fontname="helvetica", color=(0, 0, 0))
        
        start_x=2*(sizchange[0]/3)
        start_y=0
        height=20
        width=sizchange[0]/9
        end_y=start_y+height
        rect_end_x=0
        if i==pages-5:
            rect_start_x=start_x
            rect_end_x=rect_start_x+width
            rect = fitz.Rect(rect_start_x, start_y,  rect_end_x, height)
            new_page.draw_rect(rect, color=(0, 0, 0), width=1)  
            
        elif i==pages-3:
            rect_start_x=start_x+width
            rect_end_x=rect_start_x+width
            print("rect of last 2nd ",rect_start_x)
            rect = fitz.Rect(rect_start_x, start_y, rect_end_x, end_y)
            new_page.draw_rect(rect, color=(0, 0, 0), width=1) 
        
        elif i==pages-1:
            rect_start_x=start_x+2*width
             
            rect_end_x=rect_start_x+width
            print("rect of last page",rect_end_x)
            rect = fitz.Rect(rect_start_x, start_y, rect_end_x, end_y)
            new_page.draw_rect(rect, color=(0, 0, 0), width=1) 
            
            
        
        """
        if addMarker:
            addBorder(sizchange, new_page, name if i == 0 else "", i // 4 if i % 4 == 0 else 0, True)
        """
        if p < pages:
            file_path = os.path.join(in_file, f"{p + 1:02}.pdf")
            if os.path.exists(file_path):
                src_doc = fitz.open(file_path)
                src_page = src_doc[0]
                src_doc.close()
            else:
                continue  # Skip missing pages
        
        file_path = os.path.join(in_file, f"{p + 1:02}.pdf")    
        src_doc = fitz.open(file_path)
        
        src_page = src_doc[0]
       
        if page_scale<=100:
            page_scalenew = page_scale/100
          
            max_width=sizchange[0]/2*page_scalenew
            max_height=sizchange[1]*page_scalenew 
          
            
            input_ratio=(src_page.rect.width/src_page.rect.height)
           
            output_ratio=max_width/max_height
           
            
            if input_ratio > output_ratio:
                page_width=max_width
                page_height=src_page.rect.height*(max_width/src_page.rect.width)
            else:
                page_height=max_height
                page_width=src_page.rect.width*(max_height/src_page.rect.height)  
            
        else:
            print("scale cant be greater than zero")
            
        print(max_width, max_height)
        print(src_page.rect.width, src_page.rect.height)
        print(page_width, page_height)
        
        start_x=(sizchange[0]/2-page_width)/2
        start_y=(sizchange[1]-page_height)/2    
                    # Place the  page (page1) on the right side
        if i%4==0 or i%4== 3:
            newRec = fitz.Rect(int(sizchange[0]/2 + start_x),int(start_y) ,int(sizchange[0]/2 + page_width + start_x) ,int(page_height+start_y))       
            
            new_page.show_pdf_page(newRec, src_doc, 0)
            
            
        else:
            newRec = fitz.Rect(int(start_x),int(start_y) ,int(page_width + start_x) ,int(page_height+start_y))            
            
            new_page.show_pdf_page(newRec, src_doc, 0)  
            
            # Place the right page (page1) on the right side
                
        src_doc.close()
        
            
    output_path = os.path.join(out_file, f"{name}.pdf")
    
    
    
    if combined == None:
        out.save(output_path)
    
    return output_path