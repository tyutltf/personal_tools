'''
添加项目到当前Python环境里

当你决定使用这种方式来开始你的项目开发,下面是你需要了解到的
1. 推荐项目目录结构(大概样子):
    - 项目名(根目录)
        - 项目名
            - data/
            - src/
            - test/
            - xxx.py
        - addProject.py
        - main.py
        - readme.py
        - .gitignore
2. 除了一些入口程序/初始化脚本/环境配置脚本,其他源码尽量都放到项目名目录下,以避免环境变量名污染
3. 当你在导入包或者模块的时候发生了意料之外的情况,请检查在项目名(根目录下是否发生了包/模块名重名),你需要记住的是同一个虚拟环境下,优先使用addProject.py的项目具有更高的优先级

使用:

把该文件放到项目名(根目录下),然后执行下面代码:

python addProjectEnv.py

'''
from distutils.sysconfig import get_python_lib
from pathlib import Path
import os

# 项目根目录
project_path = Path(os.path.dirname((os.path.abspath(__file__))))

# pth文件目录
site_packages_path = Path(get_python_lib())
pth_path = site_packages_path / "project.pth"

if pth_path.is_file():
    data = pth_path.read_text(encoding="utf-8")
    paths = [Path(path) for path in data.split()]
    if project_path not in paths:
        paths.append(project_path)
    data = '\n'.join([str(path) for path in paths])
    pth_path.write_text(data, encoding="utf-8")
else:
    data = str(project_path)
    pth_path.write_text(data, encoding="utf-8")
