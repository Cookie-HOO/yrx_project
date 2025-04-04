def init_env(env):
    from multiprocessing import freeze_support
    freeze_support()  # windows下调用多进程且作为一个独立app使用时需要调用

    import os
    os.environ["env"] = env

    # 初始化路径
    from yrx_project.const import TEMP_PATH
    if not os.path.exists(TEMP_PATH):
        os.makedirs(TEMP_PATH)


def run_app(width, height):
    import sys
    from PyQt5.QtWidgets import QApplication
    from yrx_project.client.main import MyClient

    app = QApplication(sys.argv)
    demo = MyClient()
    demo.resize(width, height)  # 根据条件调整
    demo.show()
    sys.exit(app.exec_())
