import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QCheckBox, QComboBox, QLineEdit, QHBoxLayout, QListWidget, QMessageBox
)
from PyQt6.QtCore import Qt, QMimeData
import cups
import os
import tempfile
import json
from PyPDF2 import PdfReader, PdfWriter
from docx import Document
from pptx import Presentation
from PIL import Image
import openpyxl
import subprocess

SUPPORTED_EXTS = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.png', '.jpg', '.jpeg', '.bmp']
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'batch_print_config.json')

class DraggableFileList(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                ext = os.path.splitext(file_path)[1].lower()
                if ext in SUPPORTED_EXTS:
                    if file_path not in [self.item(i).text() for i in range(self.count())]:
                        self.addItem(file_path)
            event.acceptProposedAction()
        else:
            event.ignore()

class BatchPrintApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Batch Print Tool')
        self.resize(600, 400)
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)
        self.last_printer = self.load_last_printer()

        # File selection
        self.file_list = DraggableFileList()
        self.layout.addWidget(QLabel('Selected Files:'))
        self.layout.addWidget(self.file_list)
        self.add_files_btn = QPushButton('Add Files')
        self.add_files_btn.clicked.connect(self.add_files)
        self.layout.addWidget(self.add_files_btn)

        # Print to PDF option
        self.print_to_pdf_checkbox = QCheckBox('Print to PDF (Save as PDF)')
        self.layout.addWidget(self.print_to_pdf_checkbox)

        # Printer selection
        self.printer_combo = QComboBox()
        self.layout.addWidget(QLabel('Select Printer:'))
        self.layout.addWidget(self.printer_combo)
        self.populate_printers()

        # Duplex option
        self.duplex_checkbox = QCheckBox('Print on both sides (Duplex)')
        self.layout.addWidget(self.duplex_checkbox)

        # Color printing option
        self.color_checkbox = QCheckBox('Print in color (if supported by printer)')
        self.color_checkbox.setChecked(True)
        self.layout.addWidget(self.color_checkbox)

        # Image options
        self.image_rotate_checkbox = QCheckBox('Auto-rotate images to best fit page')
        self.image_rotate_checkbox.setChecked(True)
        self.layout.addWidget(self.image_rotate_checkbox)
        self.image_scale_checkbox = QCheckBox('Auto-scale images to fit page')
        self.image_scale_checkbox.setChecked(True)
        self.layout.addWidget(self.image_scale_checkbox)

        # Page range
        page_range_layout = QHBoxLayout()
        page_range_layout.addWidget(QLabel('Page Range:'))
        self.page_range_input = QLineEdit()
        self.page_range_input.setPlaceholderText('e.g. 1-5,8,10')
        page_range_layout.addWidget(self.page_range_input)
        self.layout.addLayout(page_range_layout)

        # Print button
        self.print_btn = QPushButton('Start Batch Print')
        self.print_btn.clicked.connect(self.start_print)
        self.layout.addWidget(self.print_btn)

        # Status
        self.status_label = QLabel('Ready.')
        self.layout.addWidget(self.status_label)

    def load_last_printer(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                    return data.get('last_printer', None)
            except Exception:
                return None
        return None

    def save_last_printer(self, printer_name):
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump({'last_printer': printer_name}, f)
        except Exception:
            pass

    def populate_printers(self):
        try:
            conn = cups.Connection()
            printers = conn.getPrinters()
            self.printer_combo.clear()
            selected_index = 0
            if printers:
                for idx, printer in enumerate(printers):
                    self.printer_combo.addItem(printer)
                    if self.last_printer and printer == self.last_printer:
                        selected_index = idx
                self.printer_combo.setCurrentIndex(selected_index)
            else:
                self.printer_combo.addItem('No printers found')
        except Exception as e:
            self.printer_combo.clear()
            self.printer_combo.addItem('Error finding printers')
            self.status_label.setText(f'Error: {e}')

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, 'Select Files', '',
            'Documents (*.pdf *.doc *.docx *.xls *.xlsx *.ppt *.pptx *.png *.jpg *.jpeg *.bmp);;All Files (*)')
        for file in files:
            if file not in [self.file_list.item(i).text() for i in range(self.file_list.count())]:
                self.file_list.addItem(file)

    def parse_page_range(self, page_range_str, num_pages):
        # Returns a list of 0-based page indices
        pages = set()
        if not page_range_str.strip():
            return list(range(num_pages))
        for part in page_range_str.split(','):
            part = part.strip()
            if '-' in part:
                start, end = part.split('-')
                start = int(start) - 1
                end = int(end)
                pages.update(range(start, end))
            else:
                idx = int(part) - 1
                if 0 <= idx < num_pages:
                    pages.add(idx)
        return sorted([p for p in pages if 0 <= p < num_pages])

    def convert_word_to_pdf(self, doc_path):
        # Try to use python-docx to read, but for PDF conversion, fallback to libreoffice if available
        try:
            # Use libreoffice for conversion
            out_pdf = tempfile.mktemp(suffix='.pdf')
            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(out_pdf), doc_path], check=True)
            pdf_path = os.path.join(os.path.dirname(doc_path), os.path.splitext(os.path.basename(doc_path))[0] + '.pdf')
            if os.path.exists(pdf_path):
                os.rename(pdf_path, out_pdf)
                return out_pdf
            else:
                return None
        except Exception as e:
            return None

    def convert_ppt_to_pdf(self, ppt_path):
        try:
            out_pdf = tempfile.mktemp(suffix='.pdf')
            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(out_pdf), ppt_path], check=True)
            pdf_path = os.path.join(os.path.dirname(ppt_path), os.path.splitext(os.path.basename(ppt_path))[0] + '.pdf')
            if os.path.exists(pdf_path):
                os.rename(pdf_path, out_pdf)
                return out_pdf
            else:
                return None
        except Exception as e:
            return None

    def convert_excel_to_pdf(self, xls_path):
        try:
            out_pdf = tempfile.mktemp(suffix='.pdf')
            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(out_pdf), xls_path], check=True)
            pdf_path = os.path.join(os.path.dirname(xls_path), os.path.splitext(os.path.basename(xls_path))[0] + '.pdf')
            if os.path.exists(pdf_path):
                os.rename(pdf_path, out_pdf)
                return out_pdf
            else:
                return None
        except Exception as e:
            return None

    def convert_image_to_pdf(self, img_path, auto_rotate=True, auto_scale=True):
        try:
            from PIL import Image, ImageOps
            out_pdf = tempfile.mktemp(suffix='.pdf')
            image = Image.open(img_path)
            # Define portrait A4 size in points (1 pt = 1/72 inch)
            a4_width, a4_height = 595, 842  # portrait A4 in points
            image = image.convert('RGB')
            img_w, img_h = image.size
            img_ratio = img_w / img_h
            a4_ratio = a4_width / a4_height
            # Always use portrait A4 page
            page_w, page_h = a4_width, a4_height
            # Rotate image if needed
            if auto_rotate and img_ratio > a4_ratio:
                image = image.rotate(90, expand=True)
                img_w, img_h = image.size
                img_ratio = img_w / img_h
            # Scale image to fit page if needed
            if auto_scale:
                scale = min(page_w / img_w, page_h / img_h)
                new_w = int(img_w * scale)
                new_h = int(img_h * scale)
                img_resized = image.resize((new_w, new_h), Image.LANCZOS)
            else:
                img_resized = image
                new_w, new_h = img_w, img_h
            # Create blank white page
            from PIL import ImageDraw
            page = Image.new('RGB', (page_w, page_h), 'white')
            # Center the image if scaled, else top-left
            if auto_scale:
                offset = ((page_w - new_w) // 2, (page_h - new_h) // 2)
            else:
                offset = (0, 0)
            page.paste(img_resized, offset)
            page.save(out_pdf, 'PDF')
            return out_pdf
        except Exception as e:
            return None

    def start_print(self):
        files = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not files:
            QMessageBox.warning(self, 'No Files', 'Please add files to print.')
            return
        print_to_pdf = self.print_to_pdf_checkbox.isChecked()
        printer_name = self.printer_combo.currentText()
        self.save_last_printer(printer_name)
        duplex = self.duplex_checkbox.isChecked()
        color = self.color_checkbox.isChecked()
        page_range_str = self.page_range_input.text()
        auto_rotate_images = self.image_rotate_checkbox.isChecked()
        auto_scale_images = self.image_scale_checkbox.isChecked()
        self.status_label.setText('Starting batch print...')
        QApplication.processEvents()

        # If printing to PDF, ask for output directory
        if print_to_pdf:
            output_dir = QFileDialog.getExistingDirectory(self, 'Select Output Folder')
            if not output_dir:
                self.status_label.setText('Batch print cancelled.')
                return
        else:
            output_dir = None

        for idx, file in enumerate(files):
            self.status_label.setText(f'Processing {os.path.basename(file)} ({idx+1}/{len(files)})...')
            QApplication.processEvents()
            ext = os.path.splitext(file)[1].lower()
            tmp_pdf = None
            try:
                if ext == '.pdf':
                    reader = PdfReader(file)
                    num_pages = len(reader.pages)
                    pages = self.parse_page_range(page_range_str, num_pages)
                    if pages and (len(pages) != num_pages):
                        writer = PdfWriter()
                        for p in pages:
                            writer.add_page(reader.pages[p])
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmpf:
                            writer.write(tmpf)
                            tmp_pdf = tmpf.name
                    else:
                        tmp_pdf = file
                elif ext in ['.doc', '.docx']:
                    tmp_pdf = self.convert_word_to_pdf(file)
                    if not tmp_pdf:
                        raise Exception('Failed to convert Word to PDF')
                elif ext in ['.ppt', '.pptx']:
                    tmp_pdf = self.convert_ppt_to_pdf(file)
                    if not tmp_pdf:
                        raise Exception('Failed to convert PowerPoint to PDF')
                elif ext in ['.xls', '.xlsx']:
                    tmp_pdf = self.convert_excel_to_pdf(file)
                    if not tmp_pdf:
                        raise Exception('Failed to convert Excel to PDF')
                elif ext in ['.png', '.jpg', '.jpeg', '.bmp']:
                    tmp_pdf = self.convert_image_to_pdf(file, auto_rotate=auto_rotate_images, auto_scale=auto_scale_images)
                    if not tmp_pdf:
                        raise Exception('Failed to convert image to PDF')
                else:
                    self.status_label.setText(f'Skipped unsupported file: {os.path.basename(file)}')
                    QApplication.processEvents()
                    continue

                # If page range is specified and not a PDF, try to apply to the converted PDF
                if tmp_pdf and ext != '.pdf' and page_range_str.strip():
                    reader = PdfReader(tmp_pdf)
                    num_pages = len(reader.pages)
                    pages = self.parse_page_range(page_range_str, num_pages)
                    if pages and (len(pages) != num_pages):
                        writer = PdfWriter()
                        for p in pages:
                            writer.add_page(reader.pages[p])
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmpf2:
                            writer.write(tmpf2)
                            tmp_pdf2 = tmpf2.name
                        if tmp_pdf != file and os.path.exists(tmp_pdf):
                            os.remove(tmp_pdf)
                        tmp_pdf = tmp_pdf2

                if print_to_pdf:
                    out_path = os.path.join(output_dir, os.path.splitext(os.path.basename(file))[0] + '.pdf')
                    with open(tmp_pdf, 'rb') as fsrc, open(out_path, 'wb') as fdst:
                        fdst.write(fsrc.read())
                else:
                    conn = cups.Connection()
                    options = {}
                    if duplex:
                        options['sides'] = 'two-sided-long-edge'
                    options['ColorModel'] = 'Color' if color else 'Gray'
                    options['FFColorMode'] = 'Color' if color else 'Gray'
                    # options['FFScreen2'] = 'Fineness' 
                    # add more options here if printer have differnt vendor-specific options

                    conn.printFile(printer_name, tmp_pdf, os.path.basename(file), options)
                if tmp_pdf and tmp_pdf != file and os.path.exists(tmp_pdf):
                    os.remove(tmp_pdf)
            except Exception as e:
                self.status_label.setText(f'Error printing {os.path.basename(file)}: {e}')
                QApplication.processEvents()
                continue
        self.status_label.setText('Batch print completed.')
        QMessageBox.information(self, 'Batch Print', 'Batch print completed.')
        self.file_list.clear()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = BatchPrintApp()
    window.show()
    sys.exit(app.exec()) 