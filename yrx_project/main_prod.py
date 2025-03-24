
if __name__ == '__main__':
    from multiprocessing import freeze_support
    freeze_support()  # windows下调用多进程且作为一个独立app使用时需要调用

    import os
    os.environ["env"] = "prod"

    import openpyxl  # 显式引入，pandas读取excel需要
    import sys
    from PyQt5.QtWidgets import QApplication
    from yrx_project.client.main import MyClient

    app = QApplication(sys.argv)
    demo = MyClient()
    demo.resize(1400, 1000)  # 根据条件调整
    demo.show()
    sys.exit(app.exec_())
