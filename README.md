# PDF to Word GUI

A small standalone desktop app to convert PDF files into Microsoft Word `.docx`.

## Features

- Pick a PDF file via a file picker
- Choose an output folder
- Convert the PDF to `.docx` using `pdf2docx`
- Close the app via a **Close** button

## Running from source

1. Activate the virtual environment (if not already active):

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

2. Run the GUI:

   ```powershell
   python pdf_to_word_gui.py
   ```

## Building a standalone executable (Windows)

This repository includes a built executable under `dist/`.

To rebuild it:

```powershell
python -m PyInstaller --onefile --windowed --name pdf_to_word_gui pdf_to_word_gui.py
```

Then the executable will be available at `dist\pdf_to_word_gui.exe`.
