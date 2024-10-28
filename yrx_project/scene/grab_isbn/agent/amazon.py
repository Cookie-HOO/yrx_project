import re
import time

from selenium.common import ElementNotInteractableException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from amazoncaptcha import AmazonCaptcha
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from yrx_project.scene.grab_isbn.agent.base import BaseAgent

"""
基于xpath的示例：
1. 所有具有"abc"属性的div元素
xpath = //div[@abc]

2. class有"abc"的元素
xpath = //*[contains(@class, "abc")]

3. 找到属性abc是def的元素
xpath = //*[@abc="def"]
"""


class AmazonAgent(BaseAgent):
    INDEX = "https://www.amazon.com/"

    def fill_captcha(self):
        captcha_url_prefix = "https://images-na.ssl-images-amazon.com/captcha"
        captcha_url = self.get_attr_by_xpath(f'//img[starts-with(@src, "{captcha_url_prefix}")]', attr="src")
        captcha_input = self.find_by_id(f'captchacharacters')
        while captcha_input and captcha_url:
            captcha = AmazonCaptcha.fromlink(captcha_url)
            captcha_input.send_keys(captcha.solve())
            captcha_input.send_keys(Keys.RETURN)
            time.sleep(2)
            captcha_url = self.get_attr_by_xpath(f'//img[starts-with(@src, "{captcha_url_prefix}")]', attr="src")
            captcha_input = self.find_by_id(f'captchacharacters')

    def get_results(self, query, top_k=-1):
        """
        :param query:
        :return: [{"title": "", "detail_url": "", "publish_year": "", ""}, {}]
        """
        query_obj = self.find_by_id('twotabsearchtextbox')
        results = []
        if not query_obj:
            return []
        # 1. 搜索
        query_obj.clear()
        query_obj.send_keys(query)

        WebDriverWait(self.driver, 3).until(
            EC.presence_of_element_located((By.XPATH, """//*[@id="nav-search-submit-button"]"""))
        )
        button_obj = self.find_by_id('nav-search-submit-button')
        if button_obj:
            button_obj.click()
        else:
            return []
        # query_obj.send_keys(query)
        # query_obj.send_keys(Keys.RETURN)

        WebDriverWait(self.driver, 3).until(
            EC.presence_of_element_located((By.XPATH, """//div[contains(@class, "s-result-list")]"""))
        )
        # 2. 遍历结果
        result_elements = self.find_all_by_xpath("""//div[contains(@class, "s-result-item")]""")
        # result_elements = self.driver.find_elements(By.XPATH, """//div[@data-cy="title-recipe"]""")
        # result_elements = self.driver.find_elements(By.CSS_SELECTOR, """.s-result-list .s-result-item""")
        if top_k > 0:
            result_elements = result_elements[:top_k]
        for element in result_elements:
            this = {}
            # 1. 名称、url
            title = self.find_by_xpath(""".//a[contains(@class, "a-text-normal")]""", root=element)
            if title is None or not title.text:
                continue
            this["title"] = title.text
            this["detail_url"] = title.get_attribute("href")

            # 2. 描述信息
            # 中文版本 | 美 Dan Toomey 丹 图米 刘丽君 李成华 卢青峰 | 2016-11出版
            summary = self.find_by_xpath(""".//div[contains(@class, "a-color-secondary")]""", root=element)
            if summary and summary.text:
                parts = summary.text.split("|")
                for part in parts:
                    publish_year = re.findall(r"\d{4}", part)
                    if publish_year:
                        publish_year = publish_year[0]
                        this["publish_year"] = publish_year
                this["summary"] = summary.text
            results.append(this)
        return results

    def into_detail(self, url):
        # 1. 开启新页面
        self.driver.get(url)
        self.fill_captcha()

        # 2. 定位
        title_xpath = """//*[@id="productTitle"]"""
        detail_xpath = """//*[@id="bookDescription_feature_div"]"""
        author_xpath = """//*[@id="bylineInfo"]//a"""
        score_xpath = """//*[@id="acrPopover"]//a"""
        language_xpath = """//*[@id="rpi-attribute-language"]//div[contains(@class, "rpi-attribute-value")]"""
        publisher_xpath = """//*[@id="rpi-attribute-book_details-publisher"]//div[contains(@class, "rpi-attribute-value")]"""
        publish_date_xpath = """//*[@id="rpi-attribute-book_details-publication_date"]//div[contains(@class, "rpi-attribute-value")]"""
        isbn_10_xpath = """//*[@id="rpi-attribute-book_details-isbn10"]//div[contains(@class, "rpi-attribute-value")]"""
        isbn_13_xpath = """//*[@id="rpi-attribute-book_details-isbn13"]//div[contains(@class, "rpi-attribute-value")]"""
        edition_xpath = """//*[@id="rpi-attribute-book_details-edition"]//div[contains(@class, "rpi-attribute-value")]"""
        page_count_xpath1 = """//*[@id="rpi-attribute-book_details-ebook_pages"]//div[contains(@class, "rpi-attribute-value")]"""
        page_count_xpath2 = """//*[@id="rpi-attribute-book_details-fiona_pages"]//div[contains(@class, "rpi-attribute-value")]"""
        image_xpath = """//*[@id="imgTagWrapperId"]//img"""

        title_text = self.get_text_by_xpath(title_xpath) or ""
        detail_text = self.get_text_by_xpath(detail_xpath)
        author_text = self.get_text_by_xpath(author_xpath)
        score_text = self.get_text_by_xpath(score_xpath)
        image_url_text = self.get_attr_by_xpath(image_xpath, "src")

        # detail_group_xpath = """//ol[contains(@class, "a-carousel")]"""
        # WebDriverWait(self.driver, 3).until(
        #     EC.presence_of_element_located((By.XPATH, isbn_13_xpath))
        # )
        language_text = self.get_text_by_xpath(language_xpath)
        publisher_text = self.get_text_by_xpath(publisher_xpath)
        publish_date_text = self.get_text_by_xpath(publish_date_xpath)
        isbn10_text = self.get_text_by_xpath(isbn_10_xpath)
        isbn13_text = self.get_text_by_xpath(isbn_13_xpath)
        # 版本信息也许直接写明，也许在title中
        edition_text = self.get_text_by_xpath(edition_xpath, regex=r"\d+") or self.get_text_by_xpath(title_xpath, r'(\d+)(?:st|nd|rd|th) Edition')
        page_count_text = self.get_text_by_xpath(page_count_xpath1, regex=r"\d+") or self.get_text_by_xpath(page_count_xpath2, regex=r"\d+")

        next_page_xpath = """//a[contains(@class, "a-carousel-goto-nextpage")]"""
        next_button = self.find_by_xpath(next_page_xpath)
        counter = 0
        while next_button is not None and counter < 4:
            counter += 1
            if next_button:
                try:
                    next_button.click()
                except ElementNotInteractableException:  # 说明到达终点了
                    break
                time.sleep(0.1)
                language_text = language_text or self.get_text_by_xpath(language_xpath)
                publisher_text = publisher_text or self.get_text_by_xpath(publisher_xpath)
                publish_date_text = publish_date_text or self.get_text_by_xpath(publish_date_xpath)
                isbn10_text = isbn10_text or self.get_text_by_xpath(isbn_10_xpath)
                isbn13_text = isbn13_text or self.get_text_by_xpath(isbn_13_xpath)
                edition_text = edition_text or self.get_text_by_xpath(edition_xpath, regex=r"\d+")
                page_count_text = self.get_text_by_xpath(page_count_xpath1, regex=r"\d+") or self.get_text_by_xpath(
                    page_count_xpath2, regex=r"\d+")
                next_button = self.find_by_xpath(next_page_xpath)

        res = {
            "author": author_text,
            "score": score_text,

            "language": language_text,
            "publisher": publisher_text,
            "publish_date": publish_date_text,
            "edition": edition_text,
            "isbn10": isbn10_text,
            "isbn13": isbn13_text,
            "page_count": page_count_text,

            "detail": detail_text,
            "image_url": image_url_text,
        }

        return res


amazon_agent = AmazonAgent()


if __name__ == '__main__':
    s = ad.into_detail("https://www.amazon.com/-/zh/dp/8120348176/ref=sr_1_fkmr0_1?__mk_zh_CN=%E4%BA%9A%E9%A9%AC%E9%80%8A%E7%BD%91%E7%AB%99&crid=W02QND096R0Z&dib=eyJ2IjoiMSJ9.dM_4T4h2mud6c86u66InUoVvnym0XHZ9QZrlSpgUv9DNCrqpOLJY0GalMFG15p0vsAfmYlEhT70RAIsQlY0eG_EvHfoVgUhcQ8wvV9hXjiEqqUitKwKedyi4cNIpQIg-.MA1sKsIu9mOZuSCsnIhFIDqBltel55tX058ikpawZjA&dib_tag=se&keywords=Quantum+Chemistry+%287th+Edition%29&qid=1729515904&sprefix=quantum+chemistry+7th+edition+%2Caps%2C2039&sr=8-1-fkmr0")
    print(s)
    # s = ad.into_index().get_results("Lehninger Principles of Biochemistry")
    # print(s)
    # time.sleep(100)

# driver.close()