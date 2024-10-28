import os
os.environ["env"] = "dev"

import openpyxl  # 显式引入，pandas读取excel需要
import sys
from PyQt5.QtWidgets import QApplication
from lwx_project.client.main import MyClient

if __name__ == '__main__':
    app = QApplication(sys.argv)
    demo = MyClient()
    demo.show()
    sys.exit(app.exec_())
