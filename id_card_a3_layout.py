import os, fitz
from io import BytesIO
from PIL import Image, ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A3
from tkinter import filedialog, messagebox
import tkinter as tk

# ---------------- CONFIG ----------------
PAGE_W_MM, PAGE_H_MM = 297.0, 420.0
CARD_W_MM, CARD_H_MM = 57.0, 90.0
ROW_TOPS_MM = [9.697, 110.297, 219.697, 320.297]  # mm from top
COLS, ROWS = 5, 4
KIDS_PER_PAGE = 10
DPI = 300
FRONT_SUFFIX, BACK_SUFFIX = "_FRONT.pdf", "_BACK.pdf"
FIRST_COL_X_MM = 6.0

# Fixed template path (must exist in same folder as this script)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "ID Layout.pdf")

# ---------------- HELPERS ----------------
def rasterize(pdf_path):
    """Convert a one-page PDF to a Pillow image (PNG)."""
    doc = fitz.open(pdf_path)
    pix = doc[0].get_pixmap(dpi=DPI)
    img_bytes = pix.tobytes("png")
    doc.close()
    return Image.open(BytesIO(img_bytes))

def gather_pairs(folder):
    """Return all FRONT/BACK pairs from one folder."""
    kids = []
    for root, _, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(FRONT_SUFFIX.lower()):
                base = f[:-len(FRONT_SUFFIX)]
                back = os.path.join(root, base + BACK_SUFFIX)
                if os.path.exists(back):
                    kids.append((base, os.path.join(root, f), back))
    kids.sort(key=lambda t: t[0].lower())
    return kids

def mm_to_bottom_left_y(top_mm, box_h_mm):
    """Convert top-Y mm to ReportLab bottom-left coordinate."""
    return PAGE_H_MM - top_mm - box_h_mm

def draw_template_background(c):
    """Draw fixed A3 template on current page."""
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")
    doc = fitz.open(TEMPLATE_PATH)
    pix = doc[0].get_pixmap(dpi=300)
    img_bytes = pix.tobytes("png")
    doc.close()
    img = Image.open(BytesIO(img_bytes))
    c.drawInlineImage(img, 0, 0, width=PAGE_W_MM*mm, height=PAGE_H_MM*mm)

def make_sheets(folders, out_pdf, log_fn=print):
    """Process multiple folders sequentially, filling pages continuously."""
    # Collect all kids, preserving folder order
    schools = [(folder, gather_pairs(folder)) for folder in folders]
    schools = [(f, kids) for f, kids in schools if kids]
    if not schools:
        raise RuntimeError("No FRONT/BACK pairs found in selected folders.")

    c = canvas.Canvas(out_pdf, pagesize=A3)
    col_xs = [FIRST_COL_X_MM + i * CARD_W_MM for i in range(COLS)]
    row_bottoms = [mm_to_bottom_left_y(t, CARD_H_MM) for t in ROW_TOPS_MM]
    slot_index = 0

    for folder, pairs in schools:
        for base, front, back in pairs:
            slot = slot_index % KIDS_PER_PAGE
            if slot == 0:
                if slot_index > 0:
                    c.showPage()
                draw_template_background(c)

            # Row mapping: row 1→back row 4, row 2→back row 3
            if slot < 5:
                row_f, row_b, col = 0, 3, slot
            else:
                row_f, row_b, col = 1, 2, slot - 5

            try:
                # FRONT
                img_f = rasterize(front)
                fx, fy = col_xs[col]*mm, row_bottoms[row_f]*mm
                c.drawInlineImage(img_f, fx, fy, width=CARD_W_MM*mm, height=CARD_H_MM*mm)

                # BACK (rotated)
                img_b = rasterize(back).rotate(180, expand=True)
                bx, by = col_xs[col]*mm, row_bottoms[row_b]*mm
                c.drawInlineImage(img_b, bx, by, width=CARD_W_MM*mm, height=CARD_H_MM*mm)

                log_fn(f"✅ {base}")
            except Exception as e:
                log_fn(f"❌ {base}: {e}")
            slot_index += 1

    c.save()
    log_fn(f"✅ Done. Saved to {out_pdf}")
    messagebox.showinfo("Success", f"Combined PDF created:\n{out_pdf}")

# ---------------- GUI ----------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("A3 ID Card Combiner — Final Continuous Version")
        root.geometry("800x550")

        tk.Label(root, text="Select Multiple School Folders:").pack(anchor="w", padx=10, pady=(10,0))
        f1 = tk.Frame(root); f1.pack(fill="x", padx=10)
        tk.Button(f1, text="Add Folder", command=self.add_folder).pack(side="left")
        tk.Button(f1, text="Clear All", command=self.clear_folders).pack(side="left", padx=5)
        self.listbox = tk.Listbox(root, height=6)
        self.listbox.pack(fill="both", expand=False, padx=10, pady=(0,10))

        tk.Label(root, text="Output PDF File:").pack(anchor="w", padx=10)
        f2 = tk.Frame(root); f2.pack(fill="x", padx=10)
        tk.Button(f2, text="Save As", command=self.pick_output).pack(side="left")
        self.out_var = tk.StringVar()
        tk.Entry(f2, textvariable=self.out_var).pack(side="left", fill="x", expand=True, padx=5)

        tk.Button(root, text="Generate Combined PDF", bg="#2563eb", fg="white",
                  command=self.run).pack(pady=10)

        tk.Label(root, text="Log:").pack(anchor="w", padx=10)
        self.log = tk.Text(root, height=20)
        self.log.pack(fill="both", expand=True, padx=10, pady=(0,10))

    def add_folder(self):
        d = filedialog.askdirectory(title="Select School Folder")
        if d and d not in self.listbox.get(0, tk.END):
            self.listbox.insert(tk.END, d)
    def clear_folders(self):
        self.listbox.delete(0, tk.END)
    def pick_output(self):
        f = filedialog.asksaveasfilename(title="Save Combined PDF",
                                         defaultextension=".pdf",
                                         initialfile="Combined_A3_PrintReady.pdf",
                                         filetypes=[("PDF files","*.pdf")])
        if f: self.out_var.set(f)
    def logit(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update_idletasks()
    def run(self):
        out = self.out_var.get().strip()
        folders = list(self.listbox.get(0, tk.END))
        if not folders:
            messagebox.showerror("Error", "Please add at least one folder.")
            return
        if not out:
            messagebox.showerror("Error", "Please choose an output PDF filename.")
            return
        if not os.path.exists(TEMPLATE_PATH):
            messagebox.showerror("Error", f"Template file not found:\n{TEMPLATE_PATH}")
            return

        self.log.delete(1.0, tk.END)
        self.logit(f"🔍 Processing {len(folders)} folder(s)...")

        try:
            make_sheets(folders, out, log_fn=self.logit)
        except Exception as e:
            self.logit(f"❌ Error: {e}")
            messagebox.showerror("Error", str(e))

# ---------------- MAIN ----------------
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
