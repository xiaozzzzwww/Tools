from os import path
import re

import PyPDF2
import docx2pdf
from tkinter import Tk, filedialog, simpledialog, Text, LEFT, RIGHT, BOTH, Scrollbar, Y, ttk, END

from PyPDF2 import PdfReader
from docx import Document
from datetime import datetime
from pdf2docx import Converter


def word2pdf():
    input_files = select_input_file("word", 1)
    if not input_files:
        update_result("Word转PDF", "未选择源文件，操作取消\n\n。")
        return
    output_directory = select_output_directory()
    if not output_directory:
        update_result("Word转PDF", "未选择目标目录，操作取消。\n\n")
        return
    for file_path in input_files:
        pdf_path = path.join(output_directory, path.basename(file_path)[:-5] + '.pdf')
        docx2pdf.convert(file_path, pdf_path)
        update_result("Word转PDF", f"{file_path} 转换完成，\n\t\t保存为 {pdf_path}\n\n")


def pdf2word():
    input_files = select_input_file("pdf", 1)
    if not input_files:
        update_result("PDF转Word", "未选择源文件，操作取消\n\n。")
        return
    output_directory = select_output_directory()
    if not output_directory:
        update_result("PDF转Word", "未选择源文件，操作取消\n\n。")
        return
    for file_path in input_files:
        word_path = path.join(output_directory, path.basename(file_path)[:-4] + '.docx')
        cv = Converter(file_path)
        cv.convert(word_path)
        cv.close()
        update_result("PDF转Word", f"{file_path} 转换完成，\n\t\t保存为 {word_path}\n\n")


def merge_word():
    input_files = select_input_file("word", 1)
    if not input_files:
        update_result("合并Words", "未选择源文件，操作取消\n\n。")
        return
    target_directory = select_output_directory()
    if not target_directory:
        update_result("合并Words", "未选择目标目录，操作取消。\n\n")
        return

    if target_directory:
        merged_path = 'Merged.docx'
        merged_doc = Document()
        for file_path in input_files:
            doc = Document(file_path)
            for element in doc.element.body:
                merged_doc.element.body.append(element)
            merged_doc.save(merged_path)
        update_result("合并Words", f"Word文件已成功合并到：\n\t\t{path.join(target_directory, merged_path)}\n\n")
    else:
        update_result("合并Words", "未选择目标目录，操作取消。\n\n")


def merge_pdf():
    # 合并PDF和转换后的Word文档
    input_files = select_input_file("pdf", 1)
    if not input_files:
        update_result("合并PDFs", "未选择源文件，操作取消\n\n。")
        return
    target_directory = select_output_directory()
    if not target_directory:
        update_result("合并PDFs", "未选择目标目录，操作取消。\n\n")
        return

    def merge_pdfs(input_paths, output_path):
        # 使用PyPDF2库将多个PDF文件合并为一个
        pdf_merger = PyPDF2.PdfMerger()
        for input_path in input_paths:
            pdf_merger.append(input_path)

        with open(output_path, 'wb') as output_file:
            pdf_merger.write(output_file)

    if target_directory:
        merged_output = f"{target_directory}/merged_output.pdf"
        merge_pdfs(input_files, merged_output)
        update_result("合并PDFs", f"{', '.join(input_files)} 合并完成，\n\t\t保存为 {merged_output}\n\n")
    else:
        update_result("合并PDFs", "未选择目标目录，操作取消。\n\n")


def split_pdf():
    input_path = select_input_file("pdf", 0)
    # 使用PyPDF2库拆分PDF文件
    if not input_path:
        update_result("PDF拆分", "未选择源文件，操作取消\n\n。")
        return

    pdf_reader = PdfReader(input_path)
    total_pages = len(pdf_reader.pages)
    if total_pages is None:
        update_result("PDF拆分", "无法获取PDF总页数，请选择有效的PDF文件\n\n")
        return
    se = simpledialog.askstring("PDF拆分", f"请输入起始页码（1 到 {total_pages}）:")
    if len(se) == 0:
        update_result("PDF拆分", "页码错误\n\n")
        return

    def splitpdf(input_path, start_page, end_page, output_directory):
        # 使用PyPDF2库拆分PDF文件
        with open(input_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfFileReader(file)
            pdf_writer = PyPDF2.PdfFileWriter()
            for page_num in range(start_page - 1, end_page):
                pdf_writer.addPage(pdf_reader.getPage(page_num))
            output_filename = path.join(output_directory,
                                        path.basename(input_path)[:-4] + "_page_{}-{}.pdf".format(start_page,
                                                                                                  page_num + 1))
            # output_filename = f"{output_directory}/page_{start_page}-{page_num + 1}.pdf"
            with open(output_filename, 'wb') as output_file:
                pdf_writer.write(output_file)
            update_result("PDF拆分",
                          f"已拆分{input_path} 的 {start_page}-{end_page}页 到文件 \n\t\t{output_filename}\n\n")

    def check_num(var):
        if not var.isdigit():
            update_result("PDF拆分", "页码输入错误，含有数字外的其他字符。\n\n")
            return 0
        return 1

    # if not (se.find(";") or se.find("；")):
    sp = '-|:|：'
    t = re.split(sp, se)
    arr = []
    if len(t) == 1:
        if se[0] in sp:
            arr.append(1)
            if check_num(t[0]):
                arr.append(int(t[0]))
        elif se[-1] in sp:
            if check_num(t[0]):
                arr.append(int(t[0]))
            arr.append(total_pages)
    elif len(t) == 2:
        check_num(t[0])
        arr.append(int(t[0]))
        check_num(t[1])
        arr.append(int(t[1]))
    else:
        update_result("PDF拆分", "页码输入错误！\n\n")
    if len(arr) == 2:
        if arr[0] < 1 or arr[0] > total_pages:
            update_result("PDF拆分", "起始页码无效，请输入有效的起始页码")
            return

        if arr[1] < arr[0] or arr[1] > total_pages:
            update_result("PDF拆分", "结束页码无效，请输入有效的结束页码")
            return
        target_directory = select_output_directory()
        splitpdf(input_path, arr[0], arr[1], target_directory)

    # else :
    #     ss = re.split(';|；',se)
    #
    # for var in t:
    #     if check_num(var):
    #         arr.append(int(var))
    # if se.endswith(':'):
    #     arr.append(total_pages)
    # for i in range(len(arr) - 1):
    #     if arr[i] > arr[i + 1]:
    #         del arr[i + 1]
    #         i = i - 1
    #
    # if not target_directory:
    #     update_result("PDF拆分", "未选择目标目录，操作取消。\n\n")
    #     return
    #
    # for i in range(len(arr) - 1):
    #     start = arr[i]
    #     end = arr[i + 1]
    #     splitpdf(input_path, start, end, target_directory)


def select_input_file(style="word", flag=0):
    # 弹出选择文件对话框，只能选择.docx文件和.pdf文件
    file_path = None
    if style == "word" and flag == 0:
        file_path = filedialog.askopenfilename(title="选择文件", filetypes=[("Word Files", "*.docx")])
    elif style == "word" and flag == 1:
        file_path = filedialog.askopenfilenames(title="选择文件", filetypes=[("Word Files", "*.docx")])
    elif style == "pdf" and flag == 0:
        file_path = filedialog.askopenfilename(title="选择文件", filetypes=[("PDF Files", "*.pdf")])
    elif style == "pdf" and flag == 1:
        file_path = filedialog.askopenfilenames(title="选择文件", filetypes=[("PDF Files", "*.pdf")])
    return file_path


def select_output_directory():
    # 弹出选择目录对话框
    output_directory = filedialog.askdirectory(title="选择目标目录")
    return output_directory


def update_result(operation, result):
    current_time = datetime.now().strftime("%m-%d %H:%M:%S")
    result_text = f"{current_time} - {operation}: {result}\n"
    text_box.insert(END, result_text)


# 创建主窗口
root = Tk()
root.title("PDF、Word文件转换与拆分工具")

# 设置按钮样式
style = ttk.Style()
style.configure("TButton", padding=6, relief="flat")

# 创建一个Frame，用于放置按钮
frame_buttons = ttk.Frame(root)
frame_buttons.pack(side=LEFT, padx=20, pady=20)

# 创建Word转PDF按钮
word_to_pdf_button = ttk.Button(frame_buttons, text="选择Word文件并转为PDF", command=word2pdf, width=30)
word_to_pdf_button.pack(pady=5)

# 创建PDF转Word按钮
pdf_to_word_button = ttk.Button(frame_buttons, text="选择PDF文件并转为Word", command=pdf2word, width=30)
pdf_to_word_button.pack(pady=5)

# 创建PDF合并按钮
merge_pdf_button = ttk.Button(frame_buttons, text="选择PDF文件并合并", command=merge_pdf, width=30)
merge_pdf_button.pack(pady=5)

# 创建Word合并按钮
merge_word_button = ttk.Button(frame_buttons, text="选择Word文件并合并", command=merge_word, width=30)
merge_word_button.pack(pady=5)

# 创建拆分PDF按钮
split_pdf_button = ttk.Button(frame_buttons, text="选择PDF文件并拆分", command=split_pdf, width=30)
split_pdf_button.pack(pady=5)

# 添加文本框
text_box = Text(root)
text_box.pack(side=RIGHT, fill=BOTH, expand=True, padx=20, pady=20)

# 添加滚动条
scrollbar = Scrollbar(root, command=text_box.yview)
scrollbar.pack(side=RIGHT, fill=Y)
text_box.config(yscrollcommand=scrollbar.set)

# 启动主事件循环
root.mainloop()
