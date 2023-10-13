import glob
import os
import time

import PySimpleGUI as sg
from win32com import client


def gui():
    # 设置整体样式
    sg.theme("SystemDefaultForReal")
    # 布局代码
    layout = [
        [
            sg.Text("选定Word文件夹:", font=("宋体", 10)),
            sg.Text("", key="all_word_text", size=(50, 1), font=("宋体", 10)),
        ],
        [
            sg.Text("合并后的Word文件夹:", font=("宋体", 10)),
            sg.Text("", key="word_text", size=(50, 1), font=("宋体", 10)),
        ],
        [sg.Text("程序运行记录", justification="center")],
        [sg.Output(size=(70, 20), font=("宋体", 10))],
        [
            sg.FolderBrowse("选定Word文件夹", key="words_folder", target="all_word_text"),
            sg.FolderBrowse("合并后的Word文件夹", key="word_folder", target="word_text"),
            sg.Button("运行"),
            sg.Button("关闭程序"),
        ],
    ]

    window = sg.Window(
        "Word合并小工具", layout, font=("宋体", 15), default_element_size=(50, 1)
    )

    while True:
        event, values = window.read()
        if event in (None, "关闭程序"):  # 如果用户关闭窗口或点击`关闭`
            break
        if event == "运行":
            words_folder = values.get("words_folder")
            word_folder = values.get("word_folder")
            if word_folder and words_folder:
                main_logic(words_folder, word_folder)
                print("合并完毕，请在合并后的Word文件夹中查看")
            else:
                print("请先选择文件夹")

    window.close()


def main_logic(words_folder, word_folder):
    """主要逻辑方法"""
    # documents 列表现在包含了文件夹中的所有文档文件
    docx_files = glob.glob(os.path.join(words_folder, "*.docx"))
    # 排序
    sorted_docx_files = sorted(docx_files, key=lambda x: os.path.basename(x))
    merged_word_path = os.path.join(word_folder, "merge_word.docx")
    # 合并word
    words_merge(merged_word_path, sorted_docx_files)
    return True


def words_merge(merged_doc_path, words_paths):
    """word合并主方法"""
    # 设置Word应用程序
    word = client.Dispatch("Word.Application")
    # 设置Word应用程序不可见 然鹅未生效
    word.Visible = False
    # 如果文件存在则先删除文件
    if os.path.exists(merged_doc_path):
        os.remove(merged_doc_path)
    # 创建一个新文档来保存合并后的内容
    merged_doc = word.Documents.Add()
    # 保存合并后的文档
    merged_doc.SaveAs(merged_doc_path)
    # 开始合并
    merged_doc = word.Documents.Open(merged_doc_path)
    # 循环打开并合并文档
    for doc_path in words_paths:
        # 打开当前文档
        current_doc = word.Documents.Open(doc_path)
        # 选择当前文档的全部内容
        current_doc.Content.WholeStory()
        # 复制当前文档的内容
        current_doc.Range().Copy()
        # 复制到第一份word文档后面（新建一行开始复制）
        # 获取文档的最后一个段落的范围
        last_paragraph_range = merged_doc.Range().Paragraphs.Last.Range
        # 将范围的位置折叠到范围的末尾 参数0表示向范围的末尾折叠
        last_paragraph_range.Collapse(0)
        # 获取新段落的范围
        new_paragraph_range = merged_doc.Range().Paragraphs.Last.Range
        # 将内容粘贴到新段落的范围并保留源格式
        new_paragraph_range.PasteAndFormat(16)
        # 获取最后一行的范围,检查最后一行的内容是否为空
        last_line_range = merged_doc.Paragraphs.Last.Range
        # 在最后一段的末尾添加一个分页符 7 = wdPageBreak
        last_line_range.InsertBreak(7)
        # 关闭当前文档
        current_doc.Close()
    # 如果为空则删除，以此来消除粘贴后默认新建一行的行为
    last_line_range = merged_doc.Paragraphs.Last.Range.Characters.Last
    if ord(last_line_range.Text) == 13:
        # 如果最后一行为空行，则删除该行
        last_line_range.Delete()
    # 保存合并后的文档
    merged_doc.Save()
    # 关闭合并文档和Word应用程序
    merged_doc.Close()
    # 退出word
    word.Quit()
    # 时间暂停2秒 不然会太快导致word程序丢失 部分电脑会存在呼叫方拒绝接受呼叫
    time.sleep(2)
    # print(f"合并完成，结果保存在 {os.path.abspath(merged_doc_path)}")
    return True


def main():
    """主程序"""
    gui()


if __name__ == "__main__":
    main()
    # words_folder = r"E:\工作文件\报告生成\2023年报告生成\测试word"
    # word_folder = r"E:\工作文件\报告生成\2023年报告生成\合并测试"
    # main_logic(words_folder, word_folder)
