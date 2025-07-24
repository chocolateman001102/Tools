# Batch Print Tool

A graphical batch printing tool for macOS that supports PDF, Word, Excel, PowerPoint, and image files. Print to a physical printer or save as PDF, with options for duplex printing and page range selection.

## Features
- Batch print PDF, Word, Excel, PowerPoint, and image files
- Print to physical printer or save as PDF (checkbox option)
- Printer selection
- Duplex (both sides) printing
- Page range selection
- Progress/status display

## Setup
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the app:
   ```bash
   python batch_print_gui.py
   ```

## Notes
- Designed for macOS
- Requires a CUPS-compatible printer for physical printing
- "Save as PDF" uses macOS's built-in PDF printer or file conversion 