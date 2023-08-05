import os
import sys

import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog, QMessageBox, \
    QInputDialog
from docx2pdf import convert as d2pc
from pdf2docx import Converter as p2dc
from docx import Document


class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Word and PDF Converter')
        self.setGeometry(100, 100, 400, 200)

        # Layout
        layout = QVBoxLayout()

        # Word to PDF Button
        word_to_pdf_button = QPushButton('Word to PDF')
        word_to_pdf_button.clicked.connect(self.word_to_pdf)
        layout.addWidget(word_to_pdf_button)

        # PDF to Word Button
        pdf_to_word_button = QPushButton('PDF to Word')
        pdf_to_word_button.clicked.connect(self.pdf_to_word)
        layout.addWidget(pdf_to_word_button)

        # Merge Word Button
        merge_word_button = QPushButton('Merge Word')
        merge_word_button.clicked.connect(self.merge_word)
        layout.addWidget(merge_word_button)

        # Merge PDF Button
        merge_pdf_button = QPushButton('Merge PDF')
        merge_pdf_button.clicked.connect(self.merge_pdf)
        layout.addWidget(merge_pdf_button)

        # Split PDF Button
        split_pdf_button = QPushButton('Split PDF')
        split_pdf_button.clicked.connect(self.split_pdf)
        layout.addWidget(split_pdf_button)

        # Central Widget
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def word_to_pdf(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '选择一个或多个 Docx 文件', '',
                                                     'Word Files (*.docx);;All Files (*)')
        if file_paths:
            output_directory = QFileDialog.getExistingDirectory(self, "选择一个输出目录")
            for file_path in file_paths:
                pdf_path = os.path.join(output_directory, os.path.basename(file_path)[:-5] + '.pdf')
                d2pc(file_path, pdf_path)
                print(output_directory)
            QMessageBox.information(self, '任务完成', 'Word 转 PDF 完成!', QMessageBox.Ok)

    def pdf_to_word(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '选择一个或多个 PDF 文件', '',
                                                     'PDF Files (*.pdf);;All Files (*)')
        if file_paths:
            output_directory = QFileDialog.getExistingDirectory(self, "选择一个输出目录")
            for file_path in file_paths:
                word_path = os.path.join(output_directory, os.path.basename(file_path)[:-4] + '.docx')
                cv = p2dc(file_path)
                cv.convert(word_path)
                cv.close()
            QMessageBox.information(self, '任务完成', 'PDF 转 Word 完成!', QMessageBox.Ok)

    def merge_word(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '选择多个Docx文件合并成一个Docx', '',
                                                     'Word Files (*.docx);;All Files (*)')
        if file_paths:
            merged_path = 'Merged.docx'
            merged_doc = Document()
            for file_path in file_paths:
                doc = Document(file_path)
                for element in doc.element.body:
                    merged_doc.element.body.append(element)
            merged_doc.save(merged_path)
            QMessageBox.information(self, '任务完成', 'Docx合并完成!', QMessageBox.Ok)

    def merge_pdf(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '选择多个PDF文件合并成一个PDF', '',
                                                     'PDF Files (*.pdf);;All Files (*)')
        if file_paths:
            merged_path = 'Merged.pdf'
            pdf_merger = PyPDF2.PdfMerger()
            for file_path in file_paths:
                pdf_merger.append(file_path)

            with open(merged_path, 'wb') as output_file:
                pdf_merger.write(merged_path)
            QMessageBox.information(self, '任务完成', 'PDF合并完成!', QMessageBox.Ok)

    def split_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择一个PDF文件', '', 'PDF Files (*.pdf);;All Files (*)')
        if file_path:
            pdf_reader = PdfReader(file_path)
            num_pages = len(pdf_reader.pages)
            page_range, _ = QInputDialog.getText(self, '页码范围', f'输入页码范围 (1-{num_pages}):')
            if page_range:
                try:
                    start_page, end_page = map(int, page_range.split('-'))
                    if 1 <= start_page <= end_page <= num_pages:
                        option = QMessageBox.question(self, '拆分选项',
                                                      '每一页都拆分成一个PDF ?',
                                                      QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                        if option == QMessageBox.Yes:
                            output_directory = QFileDialog.getExistingDirectory(self, "选择一个输出目录")
                            if output_directory:
                                for page_number in range(start_page - 1, end_page):
                                    pdf_writer = PdfWriter()
                                    pdf_writer.add_page(pdf_reader.pages[page_number])
                                    split_pdf_path = os.path.join(output_directory, f'Page_{page_number + 1}.pdf')
                                    with open(split_pdf_path, 'wb') as output_pdf:
                                        pdf_writer.write(output_pdf)
                                QMessageBox.information(self, '任务完成',
                                                        f'PDF 页码范围 {start_page}-{end_page} 中每一页都拆分成一个PDF!',
                                                        QMessageBox.Ok)
                        else:
                            output_directory, _ = QFileDialog.getSaveFileName(self, "Save Merged PDF", "",
                                                                              "PDF Files (*.pdf);;All Files (*)")
                            if output_directory:
                                pdf_writer = PdfWriter()
                                for page_number in range(start_page - 1, end_page):
                                    pdf_writer.add_page(pdf_reader.pages[page_number])
                                with open(output_directory, 'wb') as output_pdf:
                                    pdf_writer.write(output_pdf)
                                QMessageBox.information(self, '任务完成',
                                                        f'PDF 页码范围 {start_page}-{end_page} 合并成一个PDF!',
                                                        QMessageBox.Ok)
                    else:
                        QMessageBox.warning(self, '错误！', '错误的页码范围，请输入一个正确的页码范围！', QMessageBox.Ok)
                except ValueError:
                    QMessageBox.warning(self, '错误！', '错误的页码范围，请输入一个正确的页码范围！', QMessageBox.Ok)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
