import time

import ddddocr
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from yrx_project.scene.grab_isbn.agent.base import BaseAgent
from yrx_project.utils.request_util import request_img_as_bytes

slide = ddddocr.DdddOcr(det=False, ocr=False)


class DangDangAgent(BaseAgent):
    INDEX = "https://www.dangdang.com/"

    def get_results(self, query, top_k=-1):
        """
                :param query:
                :return: [{"title": "", "detail_url": "", "publish_year": "", ""}, {}]
                """
        query_obj = self.find_by_id('key_S')
        results = []
        if not query_obj:
            return []
        # 1. 搜索
        query_obj.clear()
        query_obj.send_keys(query)

        WebDriverWait(self.driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="button" and @class="button" and @dd_name="搜索按钮"]'))
        )
        button_obj = self.find_by_xpath('//input[@type="button" and @class="button" and @dd_name="搜索按钮"]')
        if button_obj:
            button_obj.click()
        else:
            return []
        # query_obj.send_keys(query)
        # query_obj.send_keys(Keys.RETURN)

        WebDriverWait(self.driver, 3).until(
            EC.presence_of_element_located((By.XPATH, """//ul[starts-with(@id, "component")]"""))
        )
        # 2. 遍历结果
        result_elements = self.find_all_by_xpath("""//ul[starts-with(@id, "component")]//li""")
        # result_elements = self.driver.find_elements(By.XPATH, """//div[@data-cy="title-recipe"]""")
        # result_elements = self.driver.find_elements(By.CSS_SELECTOR, """.s-result-list .s-result-item""")
        if top_k > 0:
            result_elements = result_elements[:top_k]
        for element in result_elements:
            this = {}
            # 1. 名称、url
            title = self.find_by_xpath(""".//a[@dd_name="单品标题"]""", root=element)
            if title is None or not title.text:
                continue
            this["title"] = title.text
            this["detail_url"] = title.get_attribute("href")

            # # 2. 描述信息
            # # 中文版本 | 美 Dan Toomey 丹 图米 刘丽君 李成华 卢青峰 | 2016-11出版
            # summary = self.find_by_xpath(""".//div[contains(@class, "a-color-secondary")]""", root=element)
            # if summary and summary.text:
            #     parts = summary.text.split("|")
            #     for part in parts:
            #         publish_year = re.findall(r"\d{4}", part)
            #         if publish_year:
            #             publish_year = publish_year[0]
            #             this["publish_year"] = publish_year
            #     this["summary"] = summary.text
            results.append(this)
        return results

    def into_detail(self, url):
        # 1. 开启新页面
        self.driver.get(url)
        self.fill_captcha()
        self.driver.get(url)

        # 2. 定位
        title_xpath = """//*[@id="product_info"]//div[@class="name_info"]//h1[@title]"""
        detail0_xpath = """//*[@id="product_info"]//div[@class="name_info"]//h2"""
        detail1_xpath = """//div[@id="abstract"]//p"""
        detail2_xpath = """//*[@id="content"]//p"""
        detail3_xpath = """//*[@id="authorIntroduction"]//p"""
        author_xpath = """//*[@id="product_info"]//div[@class="messbox_info"]//span[@dd_name="作者"]"""
        score_xpath = """//span[@id="messbox_info_comm_num"]//span[@class="star"]"""  # 的style 属性
        # language_xpath = """//*[@id="rpi-attribute-language"]//div[contains(@class, "rpi-attribute-value")]"""
        publisher_xpath = """//*[@id="product_info"]//div[@class="messbox_info"]//span[@dd_name="出版社"]//a"""
        publish_date_xpath = """//*[@id="product_info"]//div[@class="messbox_info"]"""
        # isbn_10_xpath = """//*[@id="rpi-attribute-book_details-isbn10"]//div[contains(@class, "rpi-attribute-value")]"""
        isbn_13_xpath = """//*[@id="detail_describe"]//ul[@class="key clearfix"]"""
        # edition_xpath = """//*[@id="rpi-attribute-book_details-edition"]//div[contains(@class, "rpi-attribute-value")]"""
        # page_count_xpath1 = """//*[@id="rpi-attribute-book_details-ebook_pages"]//div[contains(@class, "rpi-attribute-value")]"""
        # page_count_xpath2 = """//*[@id="rpi-attribute-book_details-fiona_pages"]//div[contains(@class, "rpi-attribute-value")]"""
        image_xpath = """//*[@id="largePic"]"""


        title_text = self.get_text_by_xpath(title_xpath) or ""
        detail_text = self.get_text_by_xpath(detail0_xpath) or ""

        author_text = self.get_text_by_xpath(author_xpath)
        if author_text:
            author_text = author_text.replace("作者:", "").replace("其他", "")
        score_text = self.get_attr_by_xpath(score_xpath, "style")
        if score_text:
            score_text = score_text.replace("width: ", "")
        image_url_text = self.get_attr_by_xpath(image_xpath, "src")

        # language_text = self.get_text_by_xpath(language_xpath)
        publisher_text = self.get_text_by_xpath(publisher_xpath)
        if publisher_text:
            publisher_text = publisher_text.replace("其他", "")
        publish_date_text = self.get_text_by_xpath(publish_date_xpath, "\d{4}")
        if publish_date_text:
            publish_date_text = publish_date_text.replace("出版时间:", "")
        # isbn10_text = self.get_text_by_xpath(isbn_10_xpath)
        isbn13_text = self.get_text_by_xpath(isbn_13_xpath, "ISBN：(.*?)\n")
        # 版本信息也许直接写明，也许在title中
        edition_text = self.get_text_by_xpath(title_xpath, regex=r"第(\d+)版")
        # page_count_text = self.get_text_by_xpath(page_count_xpath1, regex=r"\d+") or self.get_text_by_xpath(page_count_xpath2, regex=r"\d+")

        # 详情页在比较下面
        # WebDriverWait(self.driver, 3).until(
        #     EC.presence_of_element_located((By.XPATH, detail1_xpath))
        # )
        # detail_text = "\n".join([
        #     str(self.get_text_by_xpath(detail1_xpath)), str(self.get_text_by_xpath(detail2_xpath)), str(self.get_text_by_xpath(detail3_xpath)),
        # ])
        res = {
            "title": title_text,
            "author": author_text,
            "score": score_text,

            # "language": language_text,
            "publisher": publisher_text,
            "publish_date": publish_date_text,
            "edition": edition_text,
            # "isbn10": isbn10_text,
            "isbn13": isbn13_text,
            # "page_count": page_count_text,

            "detail": detail_text,
            "image_url": image_url_text,
        }
        return res

    def fill_captcha(self):
        user_xpath = """/html/body/div/div[2]/div/div/div[1]/div/div/div[3]/div/div[1]/div[1]/input"""
        password_xpath = """/html/body/div/div[2]/div/div/div[1]/div/div/div[3]/div/div[2]/div[1]/input"""

        agree_xpath = """//input[@type="radio" and @class="iconfont radio"]"""
        disagree_xpath = """//input[@type="radio" and @class="iconfont radio checked"]"""
        login_button_xpath = """//a[@class="btn"]"""

        captcha_bg_xpath = """//*[@id="bgImg"]"""
        captcha_target_xpath = """//*[@id="simg"]"""
        captcha_slider_xpath = """//*[@id="sliderBtn"]"""

        user_input = self.find_by_xpath(user_xpath)
        password_input = self.find_by_xpath(password_xpath)
        login_button = self.find_by_xpath(login_button_xpath)
        agree_checkbox = self.find_by_xpath(agree_xpath)
        disagree_checkbox = self.find_by_xpath(disagree_xpath)

        if disagree_checkbox:
            disagree_checkbox.click()

        if user_input and password_input and login_button:
            user_input.clear()
            user_input.send_keys("13552881572")
            password_input.clear()
            password_input.send_keys("DangDang240520")
            login_button.click()

            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located(
                    (By.XPATH, captcha_bg_xpath))
            )
            captcha_bg = self.find_by_xpath(captcha_bg_xpath)
            captcha_target = self.find_by_xpath(captcha_target_xpath)
            captcha_slider = self.find_by_xpath(captcha_slider_xpath)

            counter = 0

            while captcha_bg and captcha_target and captcha_slider and counter < 100:
                captcha_bg_bytes = request_img_as_bytes(captcha_bg.get_attribute("src"))
                captcha_target_bytes = request_img_as_bytes(captcha_target.get_attribute("src"))

                # center_x, center_y = slider_captcha(captcha_bg_bytes, captcha_target_bytes)
                res = slide.slide_match(captcha_target_bytes, captcha_bg_bytes, simple_target=True)
                lt_x, lt_y, rb_x, rb_y = res.get("target")
                print(lt_x, lt_y, rb_x, rb_y)
                center_x, center_y = (lt_x + rb_x) // 2, (lt_y + rb_y) // 2
                # print(center_x-63)

                ActionChains(self.driver).drag_and_drop_by_offset(captcha_slider, lt_x-30, 0).perform()
                # res = slide.slide_match(captcha_target_bytes, captcha_bg_bytes, simple_target=True)

                time.sleep(1)
                captcha_bg = self.find_by_xpath(captcha_bg_xpath)
                captcha_target = self.find_by_xpath(captcha_target_xpath)
                captcha_slider = self.find_by_xpath(captcha_slider_xpath)

                counter += 1

dangdang_agent = DangDangAgent()

if __name__ == '__main__':
    dd = DangDangAgent()
    print(dd.get_results("R语言实战"))
    # s = dd.into_detail("https://product.dangdang.com/28979509.html")
    # print(s)
    # s = ad.into_index().get_results("Lehninger Principles of Biochemistry")
    # print(s)
    # time.sleep(100)

# driver.close()