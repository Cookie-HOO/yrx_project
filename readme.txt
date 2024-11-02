# 李文萱的工作空间项目

## 设计思路
1. GUI
    - 打开后，需要登录，登录成功后进入主页面
    - 主页面左右分布，左边是场景，右边是场景详情

2. 项目结构
    conf
        场景1.json
    data
        场景1
            important
            tmp
            result
    lwx_project
        client：客户端相关代码
        scene：场景相关代码
        utils：工具函数
    static
        静态资源目录

## 一些问题
1. pyinstaller的打包问题
    - 先用pyinstaller打包一下
        pyinstaller -Fw .\lwx_project\main.py
    - 找到根目录的main.spec文件
        pathex=["."],  # 将当前项目根目录添加进去
    - 再次打包
        pyinstaller .\main_win.spec --distpath=.

2. GUI的选型问题
    tkinter：无法设置透明文字
        tk.Label的背景无法透明，在有背景图片时存在问题
    pyqt5:
        高版本的python3.12，安装pytqt5-tools时，报错poetry的兼容问题
        降级到python3.7后，安装pyqt5时，报错找不到C++ Microsoft 组件，需要下载
            配置pyqt5+designer with pycharm：
                https://blog.csdn.net/m0_57021623/article/details/123459038
            报错找不到C++ Microsoft 组件的方法
                https://blog.csdn.net/qq_37553692/article/details/128996821
    最终从后续的扩展性考虑，选择pyqt

3. pyinstaller
    路径问题
        1. 把代码目录打包到exe中
            打包：spec文件中定义打包内容
                ['lwx_project\\main_prod.py'],
                pathex=["."],
            使用
                第一行定义了入口代码，第二行定义了项目的根目录（python的搜索路径）
        2. 把静态资源打包到exe中
            打包：spec文件中定义打包内容
                datas=[('.\\lwx_project\\client\\ui', '.\\ui')],
            使用：
                将ui文件夹放到了根的ui文件夹
                根目录为：sys._MEIPASS，win下是一个C盘的tmp目录
        3. exe的外部内容
            exe文件可以看作是一个python文件
                a.txt
                main.exe
            只要a.txt和main.exe同级，那么main.exe中可以直接 with open("a.txt") 获取内容

4. python触发excel的宏
    win下用win32的包
    mac下用xlwings包
