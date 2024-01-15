import os
from multiprocessing import Lock, Pool

import PySimpleGUI as sg
import win32com.client

# 定义全局锁
lock = Lock()


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
                start_task(word_folder, pdf_folder)
                print("{0}转换完毕{0}".format("*" * 10))
            else:
                print("请先选择文件夹")

    window.close()


def initialize_word():
    try:
        return win32com.client.DispatchEx("Word.Application")
    except Exception as e:
        print(f"初始化Word应用程序失败：{str(e)}")
        return None


def word_to_pdf(args):
    word_path, pdf_path = args
    with lock:
        word = initialize_word()
        try:
            # 打开 Word 文档
            doc = word.Documents.Open(word_path)
            # 保存为 PDF
            doc.SaveAs(pdf_path)
            # 关闭 Word 文档
            doc.Close()
        except Exception as e:
            print(f"转换文件时发生错误：{str(e)}")
        finally:
            if word is not None:
                word.Quit()


def start_task(word_folder, pdf_folder):
    with Pool() as pool:
        args_list = []
        for word_file in os.listdir(word_folder):
            if word_file.endswith(".docx") or word_file.endswith(".doc"):
                word_path = os.path.join(word_folder, word_file)
                pdf_file = os.path.splitext(word_file)[0] + ".pdf"
                pdf_path = os.path.join(pdf_folder, pdf_file)
                args_list.append((word_path, pdf_path))
        pool.map(word_to_pdf, args_list)

    print("转换完成！PDF文件保存在:", pdf_folder)


# def main():
#     """主程序"""
#     gui()


if __name__ == "__main__":
    # main()
    word_folder = r"C:\Users\lenovo\Desktop\word文件夹"
    pdf_folder = r"C:\Users\lenovo\Desktop\pdf文件夹"
    start_task(word_folder, pdf_folder)
