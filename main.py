import os
import re
import time
from datetime import datetime
import pandas as pd
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

# --- 定数 ---
KEYWORD = "日産"
EXCEL_FILENAME = "yahoo_news_articles.xlsx"


def get_yahoo_news_with_selenium(keyword: str) -> list[dict]:
    """
    Seleniumを使ってYahoo!ニュースから指定されたキーワードの記事をスクレイピングする
    """
    print("--- Yahoo! News スクレイピング開始 ---")
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    search_url = f"https://news.yahoo.co.jp/search?p={keyword}&ei=utf-8"
    
    articles_data = []
    try:
        driver.get(search_url)
        
        # ▼▼▼ 変更点1: 待機する要素を新しい構造に合わせる ▼▼▼
        # "news-list" というIDを持つ要素を探すように変更
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.ID, "news-list")))
        
        with open('debug_page.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        print("💡 デバッグ用のHTMLファイル 'debug_page.html' を保存しました。")

        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # ▼▼▼ 変更点2: 記事リストの親要素を新しいIDセレクタに変更 ▼▼▼
        list_container = soup.find(id="news-list")
        
        if not list_container:
            print("❌ 記事リストが見つかりませんでした。")
            return []

        # ▼▼▼ 変更点3: 記事ひとつひとつのセレクタも全面的に見直し ▼▼▼
        # `div`タグの`news-list__item`というクラスを持つ要素を探す
        articles = list_container.find_all('div', class_=re.compile(r"news-list__item"))

        for article in articles:
            try:
                title_tag = article.find('h2', class_=re.compile(r"news-list__title"))
                title = title_tag.text.strip() if title_tag else ""

                link_tag = article.find("a", href=True)
                url = link_tag["href"] if link_tag else ""

                time_tag = article.find("time")
                date_str = time_tag.text.strip() if time_tag else "取得不可"
                
                source_tag = article.find('span', class_=re.compile(r"news-list__provider"))
                source_text = source_tag.text.strip() if source_tag else "取得不可"

                if title and url:
                    articles_data.append({
                        "タイトル": title,
                        "URL": url,
                        "投稿日": date_str,
                        "引用元": source_text
                    })
            except Exception:
                continue

    except Exception as e:
        print(f"❌ スクレイピング中にエラーが発生しました: {e}")
    finally:
        driver.quit()

    print(f"✅ Yahoo!ニュース取得件数: {len(articles_data)} 件")
    return articles_data


def write_to_excel(articles: list[dict], filename: str):
    new_df = pd.DataFrame(articles)
    file_path = Path(filename)

    if file_path.exists():
        print(f"📖 既存ファイル '{filename}' を読み込みます...")
        existing_df = pd.read_excel(file_path)
        
        if new_df.empty:
            print("✅ 新しい記事はありませんでした。ファイルは更新されません。")
            return
            
        existing_urls = set(existing_df['URL'])
        new_articles_df = new_df[~new_df['URL'].isin(existing_urls)]
        
        if new_articles_df.empty:
            print("✅ 新しい記事はありませんでした。ファイルは更新されません。")
            return
            
        print(f"➕ {len(new_articles_df)}件の新しい記事を追記します。")
        combined_df = pd.concat([existing_df, new_articles_df], ignore_index=True)
    else:
        if new_df.empty:
            print(f"📄 記事が見つかりませんでしたが、ヘッダーのみの新規ファイル '{filename}' を作成します。")
            combined_df = pd.DataFrame(columns=['タイトル', 'URL', '投稿日', '引用元'])
        else:
            print(f"📄 新規ファイル '{filename}' を作成します。")
            combined_df = new_df

    combined_df = combined_df.drop_duplicates(subset=['URL'], keep='last')
    combined_df = combined_df.sort_values(by='投稿日', ascending=False)

    combined_df.to_excel(filename, index=False, engine='openpyxl')
    print(f"💾 Excelファイル '{filename}' への書き込みが完了しました。総件数: {len(combined_df)} 件")


if __name__ == "__main__":
    yahoo_news_articles = get_yahoo_news_with_selenium(KEYWORD)
    write_to_excel(yahoo_news_articles, EXCEL_FILENAME)
