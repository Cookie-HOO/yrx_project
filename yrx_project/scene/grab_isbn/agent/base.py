import abc
import re
import time

from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By

"""
基于xpath的示例：
1. 所有具有"abc"属性的div元素
xpath = //div[@abc]

2. class有"abc"的元素
xpath = //*[contains(@class, "abc")]

3. 找到属性abc是def的元素
xpath = //*[@abc="def"]
"""


class BaseAgent(abc.ABC):
    INDEX = ""

    QUERY_INPUT_XPATH = ""
    QUERY_BUTTON_XPATH = ""
    QUERY_RESULT_XPATH = ""

    def __init__(self):
        self.driver = None

    def set_driver(self, driver):
        self.driver = driver

    def get_text_by_xpath(self, xpath, regex=None):
        ele = self.find_by_xpath(xpath)
        if ele is None:  # 找不到元素
            return None
        if not ele.text:  # 元素没有字符串
            return ""
        if not regex:  # 不用正则匹配，直接返回
            print(ele.text)
            return ele.text
        # 最长的路径：需要正则匹配返回
        result = re.findall(regex, ele.text)
        if result:
            print(result[0])
            return result[0]
        return ""

    def get_attr_by_xpath(self, xpath, attr=None):
        ele = self.find_by_xpath(xpath)
        if ele is None:  # 找不到元素
            return None
        return ele.get_attribute(attr)

    def find_by_xpath(self, xpath, root=None):
        root = root or self.driver
        try:
            ele = root.find_element(By.XPATH, xpath)
            return ele
        except NoSuchElementException:
            return None

    def find_all_by_xpath(self, xpath, root=None):
        root = root or self.driver
        try:
            ele = root.find_elements(By.XPATH, xpath)
            return ele
        except NoSuchElementException:
            return []

    def find_by_id(self, id_, root=None):
        root = root or self.driver
        try:
            ele = root.find_element(By.ID, id_)
            return ele
        except NoSuchElementException:
            return None

    def find_by_css(self, css, root=None):
        root = root or self.driver
        try:
            ele = root.find_element(By.CSS_SELECTOR, css)
            return ele
        except NoSuchElementException:
            return None

    def fill_captcha(self):
        pass

    @abc.abstractmethod
    def get_results(self, query, top_k=-1):
        pass

    @abc.abstractmethod
    def into_detail(self, url):
        pass
