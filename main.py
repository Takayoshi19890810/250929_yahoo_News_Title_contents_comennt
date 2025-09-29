import os
import re
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import openpyxl

KEYWORD = "日産"
OUTPUT_DIR = "output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "yahoo_news.xlsx")

def format_datetime(dt_obj):
    return dt_obj.strftime("%Y/%m/%d %H:%M")

def get_yahoo_news_with_selenium(keyword: str) -> list[dict]:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    search_url = f"https://news.yahoo.co.jp/search?p={keyword}&ei=utf-8&categories=domestic,world,business,it,science,life,local"
    driver.get(search_url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()
    articles = soup.find_all("li", class_=re.compile("sc-1u4589e-0"))
    articles_data = []

    for article in articles:
        try:
            title_tag = article.find("div", class_=re.compile("sc-3ls169-0"))
            title = title_tag.text.strip() if title_tag else ""
            link_tag = article.find("a", href=True)
            url = link_tag["href"] if link_tag else ""
            time_tag = article.find("time")
            date_str = time_tag.text.strip() if time_tag else ""
            formatted_date = ""

            if date_str:
                date_str = re.sub(r'\([月火水木金土日]\)', '', date_str).strip()
                try:
                    dt_obj = datetime.strptime(date_str, "%Y/%m/%d %H:%M")
                    formatted_date = format_datetime(dt_obj)
                except:
                    formatted_date = date_str

            source_text = ""
            source_tag = article.find("div", class_="sc-n3vj8g-0 yoLqH")
            if source_tag:
                inner = source_tag.find("div", class_="sc-110wjhy-8 bsEjY")
                if inner and inner.span:
                    candidate = inner.span.text.strip()
                    if not candidate.isdigit():
                        source_text = candidate

            if title and url:
                articles_data.append({
                    "タイトル": title,
                    "URL": url,
                    "投稿日": formatted_date if formatted_date else "取得不可",
                    "引用元": source_text
                })
        except:
            continue

    print(f"✅ Yahoo!ニュース件数: {len(articles_data)} 件")
    return articles_data

def save_to_excel(articles: list[dict], filepath: str):
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Yahoo"

    headers = ["タイトル", "URL", "投稿日", "引用元"]
    ws.append(headers)

    for a in articles:
        ws.append([a["タイトル"], a["URL"], a["投稿日"], a["引用元"]])

    wb.save(filepath)
    print(f"✅ Excelに保存しました: {filepath}")

if __name__ == "__main__":
    yahoo_news_articles = get_yahoo_news_with_selenium(KEYWORD)
    if yahoo_news_articles:
        save_to_excel(yahoo_news_articles, OUTPUT_FILE)
