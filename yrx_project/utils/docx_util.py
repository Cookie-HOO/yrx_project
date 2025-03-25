import os
import time
import win32com.client
import pythoncom
from PyQt5.QtCore import QObject, pyqtSignal, QThread
from PyQt5.QtGui import QPixmap, QImage
from PIL import ImageGrab
import ctypes


class WordStreamCapture(QObject):
    page_ready = pyqtSignal(int, QPixmap)  # (页码, 图像)
    completed = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, doc_path):
        super().__init__()
        self.doc_path = os.path.abspath(doc_path)
        self.abort_flag = False
        self.word = None
        self.dpi_scaling = 1.0

    def start_capture(self):
        """启动异步捕获"""
        self.thread = QThread()
        self.moveToThread(self.thread)
        self.thread.started.connect(self._capture_worker)
        self.thread.start()

    def abort_capture(self):
        """中止捕获过程"""
        self.abort_flag = True

    def _capture_worker(self):
        """工作线程主逻辑"""
        try:
            pythoncom.CoInitialize()
            self._init_word()
            doc = self._open_document()
            total_pages = self._get_exact_page_count(doc)

            for page in range(1, total_pages + 1):
                if self.abort_flag:
                    break

                start_time = time.time()
                self._navigate_to_page(doc, page)
                pixmap = self._capture_current_page()
                self.page_ready.emit(page, pixmap)

                # 动态调整等待时间
                elapsed = time.time() - start_time
                self._adaptive_wait(elapsed)

            self.completed.emit()
        except Exception as e:
            self.error_occurred.emit(str(e))
        finally:
            self._cleanup(doc)
            self.thread.quit()

    def _init_word(self):
        """初始化 Word 实例"""
        self.word = win32com.client.Dispatch("Word.Application")
        self.word.Visible = True
        self.word.ScreenUpdating = False
        self.dpi_scaling = self._get_system_scaling()

    def _open_document(self):
        """打开文档并预处理"""
        doc = self.word.Documents.Open(
            FileName=self.doc_path,
            ReadOnly=True
        )
        doc.ActiveWindow.View.Type = 3  # 页面视图
        doc.ActiveWindow.Zoom = 100
        return doc

    def _get_exact_page_count(self, doc):
        """精确获取页数"""
        try:
            return doc.ComputeStatistics(2)
        except:
            return self._get_pdf_page_count(doc)

    def _navigate_to_page(self, doc, page_num):
        """精准页面定位"""
        # 使用定位标记增强准确性
        bookmark_name = f"__page_{page_num}_"
        try:
            doc.Bookmarks(bookmark_name).Delete()
        except:
            pass

        # 插入临时书签
        self.word.Selection.GoTo(What=1, Which=2, Name=page_num)
        self.word.Selection.Bookmarks.Add(bookmark_name)

        # 二次定位验证
        self.word.Selection.GoTo(What=-1, Which=2, Name=bookmark_name)
        self._wait_ready()

    def _capture_current_page(self):
        """捕获当前页面"""
        try:
            win = self.word.ActiveWindow
            rect = (
                win.Left * self.dpi_scaling,
                (win.Top + win.CaptionHeight) * self.dpi_scaling,
                (win.Left + win.Width) * self.dpi_scaling,
                (win.Top + win.Height - win.StatusBarHeight) * self.dpi_scaling
            )
            img = ImageGrab.grab(bbox=tuple(map(int, rect)), all_screens=True)
            return QPixmap.fromImage(
                QImage(img.tobytes(), img.width, img.height, QImage.Format_RGB888)
            )
        except Exception as e:
            self.error_occurred.emit(f"Page capture failed: {str(e)}")
            return QPixmap()

    def _adaptive_wait(self, elapsed):
        """智能等待策略"""
        if elapsed < 0.2:
            time.sleep(0.05)  # 快速模式
        else:
            time.sleep(max(0, 0.3 - elapsed))  # 动态补偿

    def _wait_ready(self, timeout=3):
        """增强型等待就绪"""
        for _ in range(int(timeout * 100)):
            if not self.word.Busy and self.word.Ready:
                return
            time.sleep(0.01)
        raise TimeoutError("Word 未及时响应")

    @staticmethod
    def _get_system_scaling():
        """获取系统 DPI 缩放比例"""
        user32 = ctypes.windll.user32
        hdc = user32.GetDC(0)
        scaling = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88) / 96.0  # LOGPIXELSX
        user32.ReleaseDC(0, hdc)
        return scaling

    def _cleanup(self, doc):
        """资源清理"""
        try:
            if doc:
                doc.Close(SaveChanges=False)
            if self.word:
                self.word.Quit()
        finally:
            pythoncom.CoUninitialize()


# 使用示例
class CaptureManager:
    def __init__(self, doc_path):
        self.capturer = WordStreamCapture(doc_path)
        self.capturer.page_ready.connect(self.handle_page)
        self.capturer.completed.connect(self.handle_complete)
        self.capturer.error_occurred.connect(self.handle_error)
        self.capturer.start_capture()

    def handle_page(self, page_num, pixmap):
        """实时处理页面"""
        print(f"收到第 {page_num} 页，尺寸：{pixmap.size()}")
        pixmap.save(f"page_{page_num}.png")

    def handle_complete(self):
        print("所有页面捕获完成")

    def handle_error(self, msg):
        print(f"发生错误: {msg}")


if __name__ == "__main__":
    from PyQt5.QtWidgets import QApplication
    import sys

    app = QApplication(sys.argv)
    manager = CaptureManager("demo.docx")
    sys.exit(app.exec_())