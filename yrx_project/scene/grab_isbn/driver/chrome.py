from selenium import webdriver
from selenium.common import UnableToSetCookieException

from yrx_project.scene.grab_isbn.agent.amazon import amazon_agent
from yrx_project.scene.grab_isbn.agent.base import BaseAgent
from yrx_project.scene.grab_isbn.agent.dangdang import dangdang_agent


class ChromeDriver:
    def __init__(self, agent: BaseAgent = None):
        options = webdriver.ChromeOptions()
        # options.add_argument('--headless')
        # options.add_argument('--no-sandbox')
        # options.add_argument('--disable-gpu')
        # options.add_argument('--window-size=1920x1080')
        # options.add_argument('--disable-dev-shm-usage')
        self.driver = webdriver.Chrome(options=options)
        if agent:
            agent.set_driver(self.driver)
            self.agent = agent

    def add_cookies(self, cookies: dict):
        if isinstance(cookies, dict):
            for key, value in cookies.items():
                try:
                    self.driver.add_cookie({"name": key, "value": value})
                except UnableToSetCookieException as e:
                    pass
        elif isinstance(cookies, str):
            cookies_list = cookies.split(";")
            for cookie in cookies_list:
                kv = cookie.split("=")
                if len(kv) != 2:
                    print(kv)
                    continue
                key, value = kv
                try:
                    self.driver.add_cookie({"name": key, "value": value})
                except UnableToSetCookieException as e:
                    print(key, value)
                    pass

    def open_index(self):
        self.driver.get(self.agent.INDEX)
        self.agent.fill_captcha()
        return self

    def switch_agent(self, agent: BaseAgent):
        agent.set_driver(self.driver)
        self.agent = agent
        return self

    def get_results(self, query, top_k=-1):
        return self.agent.get_results(query=query, top_k=top_k)

    def into_detail(self, url):
        return self.agent.into_detail(url)


if __name__ == '__main__':
    cd = ChromeDriver(agent=amazon_agent)
    cd.switch_agent(agent=dangdang_agent)
    print(cd.open_index().get_results(query="R语言实战"))