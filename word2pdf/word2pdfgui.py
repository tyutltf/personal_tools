import os

import PySimpleGUI as sg
from docx2pdf import convert
from win32com.client import constants, gencache


def gui():
    # 设置整体样式
    sg.theme("SystemDefaultForReal")
    # 布局代码
    layout = [
        [
            sg.Text("选定Word文件夹:", font=("宋体", 10)),
            sg.Text("", key="word_text", size=(50, 1), font=("宋体", 10)),
        ],
        [
            sg.Text("Pdf储存文件夹:", font=("宋体", 10)),
            sg.Text("", key="pdf_text", size=(50, 1), font=("宋体", 10)),
        ],
        [sg.Text("程序运行记录", justification="center")],
        [sg.Output(size=(70, 20), font=("宋体", 10))],
        [
            sg.FolderBrowse("Word文件夹", key="word_folder", target="word_text"),
            sg.FolderBrowse("Pdf文件夹", key="pdf_folder", target="pdf_text"),
            sg.Button("运行"),
            sg.Button("关闭程序"),
        ],
    ]

    window = sg.Window(
        "Word批量转换Pdf工具", layout, font=("宋体", 15), default_element_size=(50, 1)
    )

    while True:
        event, values = window.read()
        if event in (None, "关闭程序"):  # 如果用户关闭窗口或点击`关闭`
            break
        if event == "运行":
            word_folder = values.get("word_folder")
            pdf_folder = values.get("pdf_folder")
            if word_folder and pdf_folder:
                print("{0}正在将word转换为pdf{0}".format("*" * 10))
                word_to_pdf(word_folder, pdf_folder)
                print("{0}转换完毕{0}".format("*" * 10))
            else:
                print("请先选择文件夹")

    window.close()


# def word_to_pdf(word_folder, pdf_folder):
#     """批量将word转为pdf"""
#     # 将Word文档批量转换为PDF并指定输出目录
#     convert(word_folder, pdf_folder)


def word_to_pdf(word_folder, pdf_folder):
    """
    word转PDF
    :wordPath word文件路径
    :pdfPth:生成PDF文件路径
    """
    word = gencache.EnsureDispatch("Word.Application")
    for word_file in os.listdir(word_folder):
        if not word_file.startswith("~"):
            word_path = os.path.join(word_folder, word_file)
            pdf_file = os.path.splitext(word_file)[0] + ".pdf"
            pdf_path = os.path.join(pdf_folder, pdf_file)
            doc = word.Documents.Open(word_path, ReadOnly=1)
            doc.ExportAsFixedFormat(
                pdf_path,
                constants.wdExportFormatPDF,
                Item=constants.wdExportDocumentWithMarkup,
                CreateBookmarks=constants.wdExportCreateHeadingBookmarks,
            )
            word.Quit(constants.wdDoNotSaveChanges)


def main():
    """主程序"""
    gui()


if __name__ == "__main__":
    # main()
    word_folder = r"C:\Users\lenovo\Desktop\word测试文件夹"
    pdf_folder = r"C:\Users\lenovo\Desktop\pdf文件夹"
    word_to_pdf(word_folder, pdf_folder)
