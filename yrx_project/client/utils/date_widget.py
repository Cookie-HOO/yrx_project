from PyQt5.QtCore import QDate

from yrx_project.utils.time_obj import TimeObj


class DateEditWidgetWrapper:
    """日期编辑的组件"""
    def __init__(self, date_widget, init_date: TimeObj = None):
        self.date_widget = date_widget
        self.date_widget.setCalendarPopup(True)  # 修改日期的时候弹出日历框
        self.__init_date = init_date or TimeObj()
        self.set(self.__init_date)

    def set(self, time_obj: TimeObj):
        date = QDate(time_obj.year, time_obj.month, time_obj.day)
        self.date_widget.setDate(date)

    def get(self) -> TimeObj:
        return TimeObj(self.date_widget.date().toString("yyyy-MM-dd"))
