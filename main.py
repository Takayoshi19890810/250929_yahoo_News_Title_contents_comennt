import os
import re
import time
from datetime import datetime
import pandas as pd
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
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
        time.sleep(5)  # ページ読み込み待機

        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # 記事がリストされている親要素を取得
        list_container = soup.find('ul', class_=re.compile(r"SearchList_"))
        
        if not list_container:
            print("❌ 記事リストが見つかりませんでした。サイト構造が変更された可能性があります。")
            return []

        articles = list_container.find_all('li', class_=re.compile(r"SearchList_item"))

        for article in articles:
            try:
                title_tag = article.find("p", class_=re.compile(r"SearchList_title"))
                title = title_tag.text.strip() if title_tag else ""

                link_tag = article.find("a", href=True)
                url = link_tag["href"] if link_tag else ""

                time_tag = article.find("time")
                date_str = time_tag.text.strip() if time_tag else "取得不可"
                
                source_tag = article.find("p", class_=re.compile(r"SearchList_provider"))
                source_text = source_tag.text.strip() if source_tag else "取得不可"

                if title and url:
                    articles_data.append({
                        "タイトル": title,
                        "URL": url,
                        "投稿日": date_str,
                        "引用元": source_text
                    })
            except Exception:
                continue # 個別の記事でエラーが起きても処理を続ける

    except Exception as e:
        print(f"❌ スクレイピング中にエラーが発生しました: {e}")
    finally:
        driver.quit()

    print(f"✅ Yahoo!ニュース取得件数: {len(articles_data)} 件")
    return articles_data


def write_to_excel(articles: list[dict], filename: str):
    """
    記事リストをExcelファイルに書き込む。既存ファイルがあれば新しい記事のみを追記する。
    """
    if not articles:
        print("⚠️ 新しい記事はありません。")
        return

    new_df = pd.DataFrame(articles)
    file_path = Path(filename)

    if file_path.exists():
        print(f"📖 既存ファイル '{filename}' を読み込みます...")
        existing_df = pd.read_excel(file_path)
        existing_urls = set(existing_df['URL'])
        
        # 既存リストにないURLの記事のみをフィルタリング
        new_articles_df = new_df[~new_df['URL'].isin(existing_urls)]
        
        if new_articles_df.empty:
            print("✅ 新しい記事はありませんでした。ファイルは更新されません。")
            return
            
        print(f"➕ {len(new_articles_df)}件の新しい記事を追記します。")
        combined_df = pd.concat([existing_df, new_articles_df], ignore_index=True)
    else:
        print(f"📄 新規ファイル '{filename}' を作成します。")
        combined_df = new_df

    # 重複を最終確認して削除し、投稿日でソート（日付らしき文字列でソートを試みる）
    combined_df = combined_df.drop_duplicates(subset=['URL'], keep='last')
    combined_df = combined_df.sort_values(by='投稿日', ascending=False)

    combined_df.to_excel(filename, index=False, engine='openpyxl')
    print(f"💾 Excelファイル '{filename}' への書き込みが完了しました。総件数: {len(combined_df)} 件")


if __name__ == "__main__":
    yahoo_news_articles = get_yahoo_news_with_selenium(KEYWORD)
    # 記事が1件以上あれば書き込み処理を実行
    if yahoo_news_articles:
        write_to_excel(yahoo_news_articles, EXCEL_FILENAME)
