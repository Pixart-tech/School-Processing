import fitz  # PyMuPDF
import os

MATCH_CODES = {"0101", "0102", "0103", "0104", "0105", "0106", "0107", "0208", "0308", "0408"}
BORDER_IMAGE = "border2.png"  # Ensure this path is correct

def add_border_and_scale_content(input_path, border_image_path, scale=0.85):
    original_doc = fitz.open(input_path)
    border_img = fitz.open(border_image_path)
    output_doc = fitz.open()
    
    for page in original_doc:
        rect = page.rect
        new_page = output_doc.new_page(width=rect.width, height=rect.height)
        
        # Insert border image
        new_page.insert_image(rect, filename=border_image_path)

        # Create a pixmap of the page (rasterize content)
        pix = page.get_pixmap(dpi=150)  # Higher dpi = better quality
        img_rect = fitz.Rect(0, 0, pix.width, pix.height)

        # Calculate scaled position
        target_width = rect.width * scale
        target_height = rect.height * scale
        x_offset = (rect.width - target_width) / 2
        y_offset = (rect.height - target_height) / 2
        scaled_rect = fitz.Rect(x_offset, y_offset, x_offset + target_width, y_offset + target_height)

        # Insert image (scaled content) over border
        img_stream = pix.tobytes("png")
        new_page.insert_image(scaled_rect, stream=img_stream)

    original_doc.close()
    border_img.close()
    return output_doc

def combine_pdfs_in_subfolders(root_folder):
    for subdir, dirs, files in os.walk(root_folder):
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        if pdf_files:
            combined_pdf = fitz.open()
            pdf_files.sort()

            for pdf_file in pdf_files:
                code = pdf_file[3:7]
                if code not in MATCH_CODES:
                    continue
                pdf_path = os.path.join(subdir, pdf_file)
                try:
                    print(f"Processing with border: {pdf_file}")
                    doc = add_border_and_scale_content(pdf_path, BORDER_IMAGE)
                    combined_pdf.insert_pdf(doc)
                    doc.close()
                except Exception as e:
                    print(f"Failed to process {pdf_path}: {e}")

            subfolder_name = os.path.basename(subdir)
            output_path = os.path.join(subdir, f"{subfolder_name}.pdf")
            combined_pdf.save(output_path)
            combined_pdf.close()
            print(f"Saved combined PDF: {output_path}")

main_folder_path = "B4"
combine_pdfs_in_subfolders(main_folder_path)
