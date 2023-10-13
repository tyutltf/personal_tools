import os

import PySimpleGUI as sg
from pdf2docx import Converter


def gui():
    """gui设置"""
    # 设置整体样式
    sg.theme("SystemDefaultForReal")
    # 布局代码
    layout = [
        [
            sg.Text("选定Pdf文件夹:", font=("宋体", 10)),
            sg.Text("", key="pdf_text", size=(50, 1), font=("宋体", 10)),
        ],
        [
            sg.Text("Word储存文件夹:", font=("宋体", 10)),
            sg.Text("", key="word_text", size=(50, 1), font=("宋体", 10)),
        ],
        [sg.Text("程序运行记录", justification="center")],
        [sg.Output(size=(70, 20), font=("宋体", 10))],
        [
            sg.FolderBrowse("Pdf文件夹", key="pdf_folder", target="pdf_text"),
            sg.FolderBrowse("Word文件夹", key="word_folder", target="word_text"),
            sg.Button("运行"),
            sg.Button("关闭程序"),
        ],
    ]

    window = sg.Window(
        "Pdf批量转换Word工具", layout, font=("宋体", 15), default_element_size=(50, 1)
    )

    while True:
        event, values = window.read()
        if event in (None, "关闭程序"):  # 如果用户关闭窗口或点击`关闭`
            break
        if event == "运行":
            pdf_folder = values.get("pdf_folder")
            word_folder = values.get("word_folder")
            if word_folder and pdf_folder:
                print("{0}正在将pdf转换为word{0}".format("*" * 10))
                pdf_to_word(word_folder, pdf_folder)
                print("{0}转换完毕{0}".format("*" * 10))
            else:
                print("请先选择文件夹")

    window.close()


def pdf_to_word(word_folder, pdf_folder):
    """pdf转word"""
    for pdf_file in os.listdir(pdf_folder):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            docx_file = os.path.splitext(pdf_file)[0] + ".docx"
            docx_path = os.path.join(word_folder, docx_file)
            # 将PDF转换为Word
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            print(f"{pdf_file}文件转换完毕...")
    return True


def main():
    """主程序"""
    gui()


if __name__ == "__main__":
    main()
