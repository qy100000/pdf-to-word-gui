"""Standalone PDF → Word converter GUI.

This desktop app lets a user:
- Pick a PDF file
- Pick an output folder
- Click Convert to generate a .docx in the chosen output folder
- Click Close to exit

The conversion implementation is built into this module (no dependency on `pdf_to_word.py`).
"""

import os
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

from pdf2docx import Converter


def convert_pdf_to_docx(pdf_path: str, docx_path: str) -> None:
    """Convert a PDF file into a .docx file.

    This is the same logic used by the standalone `pdf_to_word.py` script, but
    embedded here so the GUI is fully self-contained (no cross-module imports).
    """

    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"Input PDF not found: {pdf_path}")

    dest_dir = os.path.dirname(docx_path) or os.getcwd()
    os.makedirs(dest_dir, exist_ok=True)

    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

    print(f"Converted PDF -> DOCX:\n  {pdf_path}\n  {docx_path}")


def _choose_pdf(entry: tk.Entry) -> None:
    path = filedialog.askopenfilename(
        title="Select PDF file",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*")],
    )
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


def _choose_output_dir(entry: tk.Entry) -> None:
    path = filedialog.askdirectory(title="Select output folder")
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


def _convert(pdf_entry: tk.Entry, outdir_entry: tk.Entry, root: tk.Tk) -> None:
    pdf_path = pdf_entry.get().strip()
    out_dir = outdir_entry.get().strip()

    if not pdf_path:
        messagebox.showwarning("Missing PDF", "Please select a PDF file to convert.")
        return

    if not out_dir:
        messagebox.showwarning("Missing output folder", "Please select an output folder.")
        return

    if not os.path.isfile(pdf_path):
        messagebox.showerror("File not found", f"PDF not found: {pdf_path}")
        return

    try:
        os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = os.path.join(out_dir, base + ".docx")

        convert_pdf_to_docx(pdf_path, output_path)
        messagebox.showinfo("Done", f"Converted to:\n{output_path}")
    except Exception as exc:
        messagebox.showerror("Conversion failed", f"An error occurred:\n{exc}")
        traceback.print_exc()


def main() -> None:
    root = tk.Tk()
    root.title("PDF → Word converter")
    root.geometry("560x200")
    root.resizable(False, False)

    frm = tk.Frame(root, padx=12, pady=12)
    frm.pack(fill=tk.BOTH, expand=True)

    # PDF selection
    tk.Label(frm, text="PDF file:").grid(row=0, column=0, sticky=tk.W)
    pdf_entry = tk.Entry(frm, width=60)
    pdf_entry.grid(row=1, column=0, columnspan=2, sticky=tk.W)
    tk.Button(frm, text="Browse…", width=10, command=lambda: _choose_pdf(pdf_entry)).grid(
        row=1, column=2, padx=(8, 0), sticky=tk.E
    )

    # Output folder selection
    tk.Label(frm, text="Output folder:").grid(row=2, column=0, pady=(12, 0), sticky=tk.W)
    outdir_entry = tk.Entry(frm, width=60)
    outdir_entry.grid(row=3, column=0, columnspan=2, sticky=tk.W)
    tk.Button(frm, text="Browse…", width=10, command=lambda: _choose_output_dir(outdir_entry)).grid(
        row=3, column=2, padx=(8, 0), sticky=tk.E
    )

    # Buttons
    btn_frame = tk.Frame(frm)
    btn_frame.grid(row=4, column=0, columnspan=3, pady=(18, 0), sticky=tk.E)

    tk.Button(btn_frame, text="Convert", width=10, command=lambda: _convert(pdf_entry, outdir_entry, root)).pack(
        side=tk.LEFT
    )
    # Close the window (same as "Cancel" from previous versions)
    tk.Button(btn_frame, text="Close", width=10, command=root.destroy).pack(side=tk.LEFT, padx=(8, 0))

    root.mainloop()


if __name__ == "__main__":
    main()
