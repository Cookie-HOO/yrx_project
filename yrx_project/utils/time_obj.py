import datetime
import math
import re
import typing

from yrx_project.const import FULL_TIME_FORMATTER, DATE_FORMATTER, TIME_FORMATTER, POSITIVE_NUM_CHAR_MAPPING, \
    DATE_NUM_FORMATTER


class TimeObj:
    def __init__(self, raw_time=None, **kwargs):
        """
        :param raw_time_str:
        :param equal_buffer:

        :param kwargs
            :param base_time: 只在很少的时候会用到，当需要有另一个日期作为标杆时
            :param equal_buffer: 只在很少的时候会用到，和另一个日期相比判断是否一致时
        """
        self._raw_time = raw_time
        self.today = datetime.datetime.now()

        self.equal_buffer = kwargs.get("equal_buffer", 0)
        self.base_time = kwargs.get("base_time")
        self.base_time_obj = None
        if self.base_time:
            self.base_time_obj = TimeObj(raw_time=self.base_time)

    @property
    def date_str(self):
        if not self._raw_time:
            return self.today.strftime(DATE_FORMATTER)
        if isinstance(self._raw_time, str):
            # 2024-01-01
            if re.match(r"\d{4}-\d{2}-\d{2}", self._raw_time):
                return self._raw_time
            # 20240101
            elif self._raw_time.isdigit() and len(self._raw_time) == 8:
                return self._raw_time[:4] + "-" + self._raw_time[4:6] + "-" + self._raw_time[6:]
            # 2024年01月01日
            elif "年" in self._raw_time and "月" in self._raw_time and "日" in self._raw_time:
                return self._raw_time.replace("年", "-").replace("月", "-").replace("日", "")
        elif isinstance(self._raw_time, (datetime.datetime, datetime.date)):
            return self._raw_time.strftime(DATE_FORMATTER)
        elif isinstance(self._raw_time, TimeObj):
            return self._raw_time.date_str

    @property
    def time_str(self):
        if not self._raw_time:
            return self.today.strftime(TIME_FORMATTER)
        raise ValueError(f"{self._raw_time} is not None, cant use time_str")

    @property
    def time_obj(self) -> datetime.datetime:
        if not self._raw_time:
            return datetime.datetime.strptime(self.time_str, TIME_FORMATTER)
        return datetime.datetime.strptime(self.date_str, DATE_FORMATTER)

    @property
    def full_time_str(self):
        return self.time_obj.strftime(FULL_TIME_FORMATTER)

    def __eq__(self, other) -> bool:
        return abs((self.time_obj - other.time_obj).days) <= self.equal_buffer

    def __gt__(self, other) -> bool:
        return self.time_obj > other.time_obj

    def __lt__(self, other) -> bool:
        return self.time_obj < other.time_obj

    def __sub__(self, other) -> typing.Union['TimeObj', int]:
        if isinstance(other, int):
            return TimeObj(raw_time=self.time_obj - datetime.timedelta(days=other))
        elif isinstance(other, (datetime.datetime, datetime.date)):
            return (self.time_obj - other).days

    def __add__(self, other) -> typing.Union['TimeObj', int]:
        if isinstance(other, int):
            return TimeObj(raw_time=self.time_obj + datetime.timedelta(days=other))

    @property
    def month_day(self):
        return "-".join(self.date_str.split("-")[-2:])

    @property
    def month_day_in_char(self):
        return f"{self.month}月{self.day}日"

    @property
    def year(self) -> int:
        return self.time_obj.year

    @property
    def month(self) -> int:
        return self.time_obj.month

    @property
    def last_month(self) -> int:
        return self.time_obj.month - 1 if self.time_obj.month > 1 else 12

    @property
    def day(self) -> int:
        return self.time_obj.day

    @property
    def season(self) -> int:
        return math.ceil(self.month / 3)

    @property
    def season_in_char(self) -> str:
        return POSITIVE_NUM_CHAR_MAPPING.get(self.season)

    @property
    def last_season_char_with_year_num(self) -> str:
        """2020年一季度"""
        last_season_in_char = POSITIVE_NUM_CHAR_MAPPING.get(self.season - 1, "四")
        last_season_year = self.year
        if self.season == 1:
            last_season_year = self.year - 1
        return f"{last_season_year}年{last_season_in_char}季度"

    @property
    def last_season_last_day_num(self) -> str:
        """2020年一季度"""
        # 获取上个季度的最后一天
        if self.month < 4:
            return datetime.datetime(self.year - 1, 12, 31).strftime(DATE_NUM_FORMATTER)
        elif self.month < 7:
            return datetime.datetime(self.year,3, 31).strftime(DATE_NUM_FORMATTER)
        elif self.month < 10:
            return datetime.datetime(self.year,6, 30).strftime(DATE_NUM_FORMATTER)
        else:
            return datetime.datetime(self.year,9, 30).strftime(DATE_NUM_FORMATTER)

    @property
    def until_last_season(self) -> str:
        # 前三季度、前两季度、一季度、1-12月
        season_num = self.season
        if season_num == 1:
            return f"{self.year-1}年1-12月"
        elif season_num == 2:
            return f"{self.year}年一季度"
        elif season_num == 3:
            return f"{self.year}年前两季度"
        elif season_num == 4:
            return f"{self.year}年前三季度"

    @property
    def last_day_of_month(self) -> 'TimeObj':
        # 格式：5月31日
        return TimeObj(datetime.date(year=self.year, month=self.month+1, day=1) - datetime.timedelta(days=1))

    @property
    def is_first_day_of_base_year(self) -> bool:
        return self.year == self.base_time_obj.year and self.month == 1 and self.day == 1

    @property
    def is_first_day_of_base_season(self) -> bool:
        season_month = self.base_time.month - (self.base_time.month - 1) % 3
        return self.year == self.base_time_obj.year and self.month == season_month and self.day == 1

    @property
    def is_first_day_of_base_month(self) -> bool:
        return self.year == self.base_time_obj.year and self.month == self.base_time_obj.month and self.day == 1
