import win32com.client
import pythoncom
import win32clipboard
import struct
import time
from PyQt5.QtWidgets import QApplication, QLabel, QScrollArea, QWidget, QVBoxLayout
from PyQt5.QtGui import QImage, QPixmap

"""docx的截图未实现"""

def capture_word_docx_to_pixmap(docx_path):
    pythoncom.CoInitialize()
    try:
        # 启动 Word 并隐藏窗口
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 隐藏窗口

        # 显式设置窗口尺寸（关键修复！）
        word.ActiveWindow.WindowState = 0  # 0 = wdWindowStateNormal (非最大化)
        word.ActiveWindow.Width = 1600  # 足够宽的宽度
        word.ActiveWindow.Height = 1200  # 足够高的高度

        # 打开文档
        doc = word.Documents.Open(docx_path)

        # 设置视图为连续页面模式（避免分页截断）
        active_window = word.ActiveWindow
        active_window.View.Type = 6  # wdWebView (连续页面)
        active_window.View.Zoom.Percentage = 100

        # 强制重绘文档（确保内容加载）
        word.ScreenRefresh()
        time.sleep(0.5)  # 等待渲染完成

        # 选择整个文档范围并复制为图片
        doc.Range().Select()  # 显式选中整个文档
        word.Selection.CopyAsPicture()  # 复制选中内容

        # 关闭文档
        doc.Close(SaveChanges=False)
        word.Quit()

        # 从剪贴板获取图像
        return get_pixmap_from_clipboard()
    except Exception as e:
        print(f"Error: {e}")
        return None
    finally:
        pythoncom.CoUninitialize()


def get_pixmap_from_clipboard():
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
            # 获取 DIB 数据
            dib_data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)

            # 解析 BITMAPINFOHEADER (40字节)
            header = dib_data[:40]
            width = struct.unpack('<i', header[4:8])[0]
            height = struct.unpack('<i', header[8:12])[0]
            bpp = struct.unpack('<H', header[14:16])[0]

            # 构造完整 BMP 文件数据
            bmp_header = struct.pack(
                '<2sI4xI',
                b'BM',
                14 + len(dib_data),  # 文件总大小
                54  # 偏移量
            )
            full_bmp = bmp_header + dib_data

            # 转换为 QImage
            image = QImage()
            image.loadFromData(full_bmp, 'BMP')

            # 处理方向（高度为负表示自上而下）
            if height > 0:
                image = image.mirrored(False, True)

            return QPixmap.fromImage(image)
        else:
            print("剪贴板中没有图像数据")
            return None
    except Exception as e:
        print(f"剪贴板错误: {e}")
        return None
    finally:
        win32clipboard.CloseClipboard()


def display_pixmap(pixmap):
    app = QApplication([])
    window = QWidget()
    layout = QVBoxLayout()
    scroll = QScrollArea()
    label = QLabel()
    scroll.setWidget(label)
    layout.addWidget(scroll)
    window.setLayout(layout)
    window.resize(800, 600)
    label.setPixmap(pixmap)
    window.show()
    app.exec_()


if __name__ == "__main__":
    docx_path = r"D:\projects\yrx_project\test1.docx"  # 替换为实际路径
    pixmap = capture_word_docx_to_pixmap(docx_path)
    if pixmap and not pixmap.isNull():
        display_pixmap(pixmap)
    else:
        print("无法生成截图")