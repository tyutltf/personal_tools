import glob
import os

import moviepy.editor as mp
import PySimpleGUI as sg
# from pygifsicle import optimize


def gui():
    # 设置整体样式
    sg.theme("SystemDefaultForReal")
    # 布局代码
    layout = [
        [
            sg.Text("选定Mp4文件夹:", font=("宋体", 10)),
            sg.Text("", key="mp4_text", size=(50, 1), font=("宋体", 10)),
        ],
        [
            sg.Text("生成Gif文件夹:", font=("宋体", 10)),
            sg.Text("", key="gif_text", size=(50, 1), font=("宋体", 10)),
        ],
        [sg.Text("程序运行记录", justification="center")],
        [sg.Output(size=(70, 20), font=("宋体", 10))],
        [
            sg.FolderBrowse("选定Mp4文件夹", key="mp4_folder", target="mp4_text"),
            sg.FolderBrowse("生成Gif文件夹", key="gif_folder", target="gif_text"),
            sg.Button("运行"),
            sg.Button("关闭程序"),
        ],
    ]

    window = sg.Window(
        "Mp4转Gif小工具", layout, font=("宋体", 15), default_element_size=(50, 1)
    )

    while True:
        event, values = window.read()
        if event in (None, "关闭程序"):  # 如果用户关闭窗口或点击`关闭`
            break
        if event == "运行":
            mp4_folder = values.get("mp4_folder")
            gif_folder = values.get("gif_folder")
            if gif_folder and mp4_folder:
                mp4_to_gif(mp4_folder, gif_folder)
                print("转换完毕，请在转换后的Gif文件夹中查看")
            else:
                print("请选择正确的文件夹")

    window.close()


def mp4_to_gif(mp4_folder, gif_folder):
    """mp4转gif"""
    mp4_files = glob.glob(os.path.join(mp4_folder, "*.mp4"))

    for mp4_file in mp4_files:
        print(mp4_file)
        # 生成输出GIF文件名
        gif_file = os.path.join(
            gif_folder, os.path.splitext(os.path.basename(mp4_file))[0] + ".gif"
        )

        # 打开视频文件
        clip = mp.VideoFileClip(mp4_file)

        # 生成GIF
        # fps：帧速率（Frames Per Second），指定生成GIF的每秒帧数。较高的帧速率会使动画更流畅，但可能会导致更大的文件大小
        # fuzz：模糊度，控制GIF的质量。较低的值会导致较小的文件大小，但可能会降低图像质量。较高的值会提高图像质量，但可能会导致较大的文件大小。
        clip.write_gif(gif_file, fps=5, program="ffmpeg", opt="opt", fuzz=0.5)
        # 使用pygifsicle进行优化 需要将gifsicle添加到PATH中 如果不添加则需要将gifsicle.exe放入根目录下
        # optimize(gif_file)

    return True


def main():
    """主程序"""
    gui()


if __name__ == "__main__":
    main()
    # mp4_folder = r"E:\个人文件\测试文档\MP4"
    # gif_folder = r"E:\个人文件\测试文档\GIF"
    # mp4_to_gif(mp4_folder, gif_folder)
