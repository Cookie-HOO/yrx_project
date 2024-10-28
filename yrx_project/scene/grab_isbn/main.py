import concurrent.futures
import re
import threading
import time

import pandas as pd

from yrx_project.scene.grab_isbn.agent.amazon import amazon_agent
from yrx_project.scene.grab_isbn.agent.dangdang import dangdang_agent
from yrx_project.scene.grab_isbn.driver.chrome import ChromeDriver
from yrx_project.scene.grab_isbn.utils.book_util import has_author_and_in_text, use_chine_edition
from yrx_project.utils.df_util import is_not_empty, is_any_empty, all_empty
from yrx_project.utils.time_util import format_duration


def hit(detail_one, row):
    publisher = row["国内出版单位\n（中译本和影印本）"] if use_chine_edition(row) else row["原版出版单位"]
    first_isbn13 = detail_one.get("isbn13") or "--"
    detail_publish_year = re.findall(r"\d{4}", detail_one.get("publish_date") or "")
    if detail_publish_year:
        detail_publish_year = detail_publish_year[0]

    # 命中：出版年份
    first_hit_publish_year = int(
        is_not_empty(row["publish_year"]) and (
                (detail_publish_year or "") == row["publish_year"]
        )
    )
    if first_hit_publish_year == 0:
        if is_any_empty(row["publish_year"], detail_publish_year):
            first_hit_publish_year = "--"

    # 命中：出版社
    first_hit_publisher = int(
        is_not_empty(publisher) and (
                publisher == (detail_one.get("publisher") or "") or
                publisher in (detail_one.get("publisher") or "") or
                ((detail_one.get("publisher") or "") and (detail_one.get("publisher") or "") in publisher) or
                publisher in (detail_one.get("detail") or "")
        )
    )
    if first_hit_publisher == 0:
        if is_any_empty(publisher, detail_one.get("publisher")):
            first_hit_publisher = "--"

    # 命中：作者
    first_hit_author = int(
        all_empty(row["主编姓名"], row["译者姓名（中译本）"]) and (
            # 和作者类的匹配
                has_author_and_in_text(row["主编姓名"], detail_one.get("author") or "") or
                has_author_and_in_text(detail_one.get("author") or "", row["主编姓名"]) or

                # 在详情页中
                has_author_and_in_text(row["主编姓名"], detail_one.get("detail") or "") or
                has_author_and_in_text(row["译者姓名（中译本）"], detail_one.get("detail") or "") or

                # 在书名中
                has_author_and_in_text(row["主编姓名"], detail_one.get("title") or "") or
                has_author_and_in_text(row["译者姓名（中译本）"], detail_one.get("title") or "")
        )
    )
    if first_hit_author == 0:
        if is_any_empty(row["主编姓名"] or row["译者姓名（中译本）"], detail_one.get("author")):
            first_hit_author = "--"

    # 命中：版次
    first_hit_edition = int(
        is_not_empty(row["版次"]) and (
                (detail_one.get("edition") or "") == str(row["版次"])
        )
    )
    if first_hit_edition == 0:
        if is_any_empty(row["版次"], detail_one.get("edition")):
            first_hit_edition = "--"

    return (
        first_hit_publish_year,
        # f'{row["publish_year"]}\n{detail_one.get("publish_date") or "--"}',
        first_hit_publisher,
        # f'{publisher}\n{detail_one.get("publisher") or "--"}',
        first_hit_author,
        # f'{row["主编姓名"] + " " + row["译者姓名（中译本）"]}\n{detail_one.get("author") or "--"}',
        first_hit_edition,
        # f'{row["版次"]}\n{detail_one.get("edition") or "--"}',
        first_isbn13
    )


# 创建线程局部变量
global_start = time.time()
new_cols = [
    "first_hit_publish_year",
    "first_hit_publisher",
    "first_hit_author",
    "first_hit_edition",
    "first_isbn13",

    "second_hit_publish_year",
    "second_hit_publisher",
    "second_hit_author",
    "second_hit_edition",
    "second_isbn13",

    "third_hit_publish_year",
    "third_hit_publisher",
    "third_hit_author",
    "third_hit_edition",
    "third_isbn13",

    "title",
    "publish_year", "summary", "detail_url",
    "author", "score",
    "language", "publisher", "publish_date", "edition", "isbn10", "isbn13",
    "page_count", "image_url",
    "agent",
]
thread_local = threading.local()


def show_progress(start_time, counter, total):
    cost = time.time() - start_time
    avg_cost = cost / counter
    todo_cost = (total - counter) * avg_cost
    cost_str = format_duration(cost)
    avg_cost_str = format_duration(avg_cost, only_sec=True)  # 这里只需要显示秒即可
    todo_cost_str = format_duration(todo_cost)
    print(
        f"{counter}/{total}\t"
        f"{round(counter / total * 100, 2)}%\t"
        f"COST: {cost_str}\t"
        f"AVG: {avg_cost_str}/个\t"
        f"TODO: {todo_cost_str}"
    )


def get_result_by_with_threading_local(row, counter, total):
    """
    df[["first_and_hit", "search_length", "title", "publish_year", "summary", "detail_url"]] = df.apply(get_result_by_amazon, axis=1)
    :param row:
    :return:
    """
    # 结果列表页
    first_hit_publish_year = None
    first_hit_publisher = None
    first_hit_author = None
    first_hit_edition = None
    first_isbn13 = None

    search_length = None
    title = None
    publish_year = None
    detail_url = None
    summary = None

    # 结果详情页
    author = None
    score = None

    language = None
    publisher = None
    publish_date = None
    edition = None
    isbn10 = None
    isbn13 = None
    page_count = None

    detail = None
    image_url = None
    agent = None

    hit_list = [None] * 3 * 9  # 每条记录有5个hit值
    try:
        if not hasattr(thread_local, "driver"):
            thread_local.driver = ChromeDriver()

        if row["教材名称"]:
            if use_chine_edition(row):
                thread_local.driver.switch_agent(agent=dangdang_agent).open_index()
                # thread_local.driver.add_cookies(
                #     dangdang_cookie
                # )
                if is_not_empty(row["国内出版单位\n（中译本和影印本）"]):
                    row["教材名称"] += " " + row["国内出版单位\n（中译本和影印本）"]
                if is_not_empty(row["版次"]):
                    res = re.findall(r"\d+", str(row["版次"]))
                    if res and int(res[0]) > 1:
                        row["教材名称"] += " 第" + res[0] + "版"
                agent = "当当"
            else:
                thread_local.driver.switch_agent(agent=amazon_agent).open_index()
                agent = "亚马逊"

            # 获取结果列表页
            search_result = thread_local.driver.get_results(row["教材名称"], top_k=3)
            # if row["publish_year"]:
            #     search_result = [i for ind, i in enumerate(search_result) if i.get("publish_year") == row["publish_year"] or ind == 0]
            # else:
            search_result = [i for ind, i in enumerate(search_result) if ind < 3]  # 取前3条
            search_length = len(search_result)
            title = "\n".join([i.get("title", "--") or "--" for i in search_result])
            publish_year = "\n".join([i.get("publish_year", "--") or "--" for i in search_result])
            detail_urls = [i.get("detail_url", "--") or "--" for i in search_result]
            detail_url = "\n".join(detail_urls)
            summary = "\n".join([i.get("summary", "--") or "--" for i in search_result])
            # first_hit_publish_year = int(search_length == 1 and publish_year == row["publish_year"])

            # 获取结果详情页
            detail_results = [thread_local.driver.into_detail(url=detail_url) for detail_url in detail_urls]
            """
            {
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
            """
            author = "\n".join([i.get("author", "--") or "--" for i in detail_results])
            score = "\n".join([i.get("score", "--") or "--" for i in detail_results])

            language = "\n".join([i.get("language", "--") or "--" for i in detail_results])
            publisher = "\n".join([i.get("publisher", "--") or "--" for i in detail_results])
            publish_date = "\n".join([i.get("publish_date", "--") or "--" for i in detail_results])
            edition = "\n".join([i.get("edition", "--") or "--" for i in detail_results])
            isbn10 = "\n".join([i.get("isbn10", "--") or "--" for i in detail_results])
            isbn13 = "\n".join([i.get("isbn13", "--") or "--" for i in detail_results])
            page_count = "\n".join([i.get("page_count", "--") or "--" for i in detail_results])

            detail = "\n".join([i.get("detail", "--") or "--" for i in detail_results])
            image_url = "\n".join([i.get("image_url", "--") or "--" for i in detail_results])

            # 拼接是否hit
            while len(detail_results) < 3:
                detail_results.append({})
            hit_list = []
            for detail_one in detail_results:
                hit_list.extend(hit(detail_one, row))
                # res = hit(detail_one, row)
    except Exception as e:
        raise e
        pass
    finally:
        counter.add()
        show_progress(start_time=global_start, counter=counter.value, total=total)
        hit_result = hit_list
        other_result = [
            # search_length,
            title,
            publish_year,
            summary,
            detail_url,

            author,
            score,

            language,
            publisher,
            publish_date,
            edition,
            isbn10,
            isbn13,
            page_count,

            # detail,
            image_url,
            agent,
        ]
        return hit_result[:(len(new_cols)-len(other_result))] + other_result


def find_publish_year(publish_date):
    if pd.isna(publish_date):
        return publish_date
    publish_year = re.findall(r"\d{4}", str(publish_date))
    if publish_year:
        return publish_year[0]
    return ""


df = pd.read_excel("./books.xlsx")
df["publish_year"] = df["出版\n时间"].apply(find_publish_year)

# 获取df后n行
length = 582
df = df.head(length)


class AtomicCounter:
    def __init__(self):
        self.value = 0
        self._lock = threading.Lock()

    def add(self):
        with self._lock:
            self.value += 1
            return self.value


counter = AtomicCounter()

# 创建一个线程池
with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:  # 用多进程尝试下速度
    # 使用map函数，将my_function并行地应用到df的每一行
    result = list(executor.map(lambda x: get_result_by_with_threading_local(x, counter, total=len(df)), [row for _, row in df.iterrows()]))

# 将结果放到df中
df[new_cols] = pd.DataFrame(result, index=df.index)


# df[["first_and_hit", "length", "title", "publish_year", "detail_url",
#     "summary"]] = df.apply(get_result_by_amazon, axis=1)
df.to_excel(f"./全部结果-{length}.xlsx", index=False)
