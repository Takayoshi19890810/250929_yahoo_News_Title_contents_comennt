import os
import re
import time
from datetime import datetime, timedelta
import pandas as pd
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

# --- 定数 ---
KEYWORD = "日産"
EXCEL_FILENAME = "nissan_yahoo_news.xlsx"

def format_datetime(dt_obj):
    """datetimeオブジェクトを指定の書式（YYYY/MM/DD HH:MM）の文字列に変換する"""
    return dt_obj.strftime("%Y/%m/%d %H:%M")

def get_yahoo_news_with_selenium(keyword: str) -> list[dict]:
    """
    Seleniumを使ってYahoo!ニュースから指定されたキーワードの記事をスクレイピングする
    """
    print("--- Yahoo! News スクレイピング開始 ---")
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    search_url = f"https://news.yahoo.co.jp/search?p={keyword}&ei=utf-8"
    driver.get(search_url)
    time.sleep(5) # ページ読み込み待機

    # 複数回スクロールして記事を最大限読み込む
    for _ in range(5):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    articles = soup.find_all("li", class_=re.compile(r"SearchList_item"))
    articles_data = []

    for article in articles:
        try:
            title_tag = article.find("p", class_=re.compile(r"SearchList_title"))
            title = title_tag.text.strip() if title_tag else ""

            link_tag = article.find("a", href=True)
            url = link_tag["href"] if link_tag else ""

            time_tag = article.find("time")
            date_str = time_tag.text.strip() if time_tag else ""
            formatted_date = "取得不可"
            if date_str:
                # "（月）"などの曜日表記を削除
                date_str_no_day = re.sub(r'\s*\（[月火水木金土日]\）', '', date_str)
                try:
                    # 'M/D HH:mm'形式をパース
                    dt_obj = datetime.strptime(f"{datetime.now().year}/{date_str_no_day}", "%Y/%m/%d %H:%M")
                    formatted_date = format_datetime(dt_obj)
                except ValueError:
                    # パース失敗時は元の文字列をそのまま利用
                    formatted_date = date_str

            source_tag = article.find("p", class_=re.compile(r"SearchList_provider"))
            source_text = source_tag.text.strip() if source_tag else "取得不可"

            if title and url:
                articles_data.append({
                    "タイトル": title,
                    "URL": url,
                    "投稿日": formatted_date,
                    "引用元": source_text
                })
        except Exception as e:
            print(f"⚠️ 記事処理エラー: {e}")
            continue

    print(f"✅ Yahoo!ニュース取得件数: {len(articles_data)} 件")
    return articles_data

def write_to_excel(articles: list[dict], filename: str):
    """
    記事リストをExcelファイルに書き込む。既存ファイルがあれば新しい記事のみを追記する。
    """
    if not articles:
        print("⚠️ 追記する新しい記事はありません。")
        return

    # 新しく取得したデータをDataFrameに変換
    new_df = pd.DataFrame(articles)
    
    # 日付文字列をdatetimeオブジェクトに変換（変換できないものはNaT）
    new_df['投稿日時'] = pd.to_datetime(new_df['投稿日'], format='%Y/%m/%d %H:%M', errors='coerce')

    excel_file = Path(filename)
    if excel_file.exists():
        print(f"📖 既存ファイル '{filename}' を読み込んでいます...")
        try:
            existing_df = pd.read_excel(excel_file)
            existing_urls = set(existing_df['URL'])

            # 既存リストにないURLの記事のみをフィルタリング
            new_articles_df = new_df[~new_df['URL'].isin(existing_urls)].copy()
            
            if new_articles_df.empty:
                print("✅ 新しい記事はありませんでした。ファイルは更新されません。")
                return

            print(f"➕ {len(new_articles_df)}件の新しい記事を追記します。")
            
            # 既存のDataFrameと新しい記事のDataFrameを結合
            combined_df = pd.concat([existing_df, new_articles_df], ignore_index=True)
            # 日付列も同様に変換
            combined_df['投稿日時'] = pd.to_datetime(combined_df['投稿日'], format='%Y/%m/%d %H:%M', errors='coerce')

        except Exception as e:
            print(f"⚠️ Excelファイルの読み込みに失敗しました: {e}。ファイルを上書きします。")
            combined_df = new_df

    else:
        print(f"📄 新規ファイル '{filename}' を作成します。")
        combined_df = new_df

    # 投稿日時で降順にソート（NaTは末尾へ）
    combined_df.sort_values(by='投稿日時', ascending=False, inplace=True, na_position='last')
    
    # 一時的な日時列を削除
    final_df = combined_df.drop(columns=['投稿日時'])

    # Excelファイルに書き出し
    try:
        final_df.to_excel(filename, index=False, engine='openpyxl')
        print(f"💾 Excelファイル '{filename}' への書き込みが完了しました。総件数: {len(final_df)} 件")
    except Exception as e:
        print(f"❌ Excelファイルへの書き込み中にエラーが発生しました: {e}")


if __name__ == "__main__":
    yahoo_news_articles = get_yahoo_news_with_selenium(KEYWORD)
    if yahoo_news_articles:
        write_to_excel(yahoo_news_articles, EXCEL_FILENAME)
