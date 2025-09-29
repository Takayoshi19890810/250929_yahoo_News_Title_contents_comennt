import os
import json
import time
import re
import pandas as pd
from datetime import datetime
import google.generativeai as genai

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import requests

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# --- 設定項目 ---
KEYWORD = "日産"
EXCEL_FILE = "yahoo_news_analysis.xlsx"
AI_MODEL_NAME = "gemini-1.0-pro" # 安定版のモデル名

MAX_BODY_PAGES = 10
MAX_COMMENT_PAGES = 10
MAX_TOTAL_COMMENTS = 500


# --- AI分析の設定 ---
try:
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
    if not GEMINI_API_KEY:
        print("警告: 環境変数 'GEMINI_API_KEY' が設定されていません。AI分析はスキップされます。")
        AI_ENABLED = False
    else:
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel(AI_MODEL_NAME)
        AI_ENABLED = True
except Exception as e:
    print(f"AIモデルの初期化に失敗しました: {e}")
    AI_ENABLED = False

# --- 共通関数 ---

def format_datetime_str(dt_obj):
    return dt_obj.strftime("%Y/%m/%d %H:%M")

def get_yahoo_news_with_selenium(keyword: str) -> list[dict]:
    print("--- Yahoo!ニュースのスクレイピングを開始 ---")
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    articles_data = []
    try:
        search_url = f"https://news.yahoo.co.jp/search?p={keyword}&ei=utf-8"
        driver.get(search_url)
        time.sleep(3)

        soup = BeautifulSoup(driver.page_source, "html.parser")
        articles = soup.find_all("li", class_=re.compile("sc-1u4589e-0"))

        for article in articles:
            title_tag = article.find("div", class_=re.compile("sc-3ls169-0"))
            link_tag = article.find("a", href=True)
            time_tag = article.find("time")
            source_tag_outer = article.find("div", class_=re.compile("sc-n3vj8g-0"))
            
            if not all([title_tag, link_tag, time_tag, source_tag_outer]):
                continue

            title = title_tag.text.strip()
            url = link_tag["href"]
            date_str = time_tag.text.strip()
            source = source_tag_outer.text.strip()
            
            try:
                cleaned_date_str = re.sub(r'\([月火水木金土日]\)', '', date_str)
                dt_obj = datetime.strptime(cleaned_date_str, "%Y/%m/%d %H:%M")
                formatted_date = format_datetime_str(dt_obj)
            except ValueError:
                formatted_date = date_str

            if title and url:
                articles_data.append({
                    "タイトル": title, "URL": url, "投稿日": formatted_date, "引用元": source
                })
    finally:
        driver.quit()
        
    print(f"✅ Yahoo!ニュースから {len(articles_data)} 件の記事情報を取得しました。")
    return articles_data

def fetch_article_pages(base_url: str) -> list[str]:
    body_pages = []
    for page in range(1, MAX_BODY_PAGES + 1):
        url = base_url if page == 1 else f"{base_url}?page={page}"
        try:
            res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
            res.raise_for_status()
            soup = BeautifulSoup(res.text, "html.parser")
            
            article_body = soup.find("div", class_=re.compile("article_body"))
            if not article_body:
                if page > 1: break
                continue

            page_text = "\n".join(p.get_text(strip=True) for p in article_body.find_all("p"))
            
            if not page_text or (body_pages and page_text == body_pages[-1]):
                break
            
            body_pages.append(page_text)
            time.sleep(1)
        except requests.RequestException:
            break
    return body_pages

def fetch_comments_with_selenium(base_url: str) -> tuple[int, list[str]]:
    total_comments = []
    per_page_comments_text = []

    if "/articles/" not in base_url:
        return 0, []

    article_id = base_url.split("/articles/")[-1]
    comment_base_url = f"https://news.yahoo.co.jp/articles/{article_id}/comments"
    
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        for page in range(1, MAX_COMMENT_PAGES + 1):
            comment_page_url = f"{comment_base_url}?page={page}"
            driver.get(comment_page_url)

            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "ul[class*='comments-list']"))
                )
            except TimeoutException:
                if page == 1:
                    print("  - コメントが見つかりませんでした（タイムアウト）。")
                break

            soup = BeautifulSoup(driver.page_source, "html.parser")
            
            selectors = [
                "p.sc-169yn8p-10",
                "p[data-ylk*='cm_body']",
                "p[class*='comment']",
                "div[data-testid='comment-body-text']",
                "div.commentBody, p.commentBody",
                "div[data-ylk*='cm_body']"
            ]
            
            p_candidates = []
            for sel in selectors:
                p_candidates.extend(soup.select(sel))

            current_page_comments = [p.get_text(strip=True) for p in p_candidates if p.get_text(strip=True)]
            current_page_comments = list(dict.fromkeys(current_page_comments))

            if not current_page_comments:
                break

            total_comments.extend(current_page_comments)
            per_page_comments_text.append("\n---\n".join(current_page_comments))

            if len(total_comments) >= MAX_TOTAL_COMMENTS:
                break
    except Exception as e:
        print(f"  - コメント取得中に予期せぬエラー: {e}")
    finally:
        driver.quit()
        
    return len(total_comments), per_page_comments_text

def analyze_article_with_ai(title: str, body_text: str) -> dict:
    default_response = {
        "title_sentiment": "N/A", "title_category": "N/A",
        "body_sentiment": "N/A", "body_category": "N/A"
    }

    if not AI_ENABLED or not title:
        return default_response

    analyze_body = bool(body_text and isinstance(body_text, str) and len(body_text.strip()) >= 10)

    optional_keys_string = ""
    if analyze_body:
        optional_keys_string = '\n- "body_sentiment"\n- "body_category"'

    prompt = f"""
    以下のニュース記事の「タイトル」と「本文」を分析し、結果を単一のJSON形式で返してください。

    分析項目:
    1. sentiment: 全体の論調が「ポジティブ」「ネガティブ」「ニュートラル」のいずれか。
    2. category: 内容に最も合致するカテゴリを以下から1つ選択してください。
       ["新技術・研究開発", "経営・財務", "販売・マーケティング", "生産・品質", "人事・労務", "不祥事・訴訟", "その他"]

    分析対象:
    ---
    タイトル: {title}
    ---
    本文: {body_text[:2000] if analyze_body else "（本文なし）"}
    ---

    JSONのキーは以下のように指定してください:
    - "title_sentiment"
    - "title_category"{optional_keys_string}

    JSON出力のみ記述してください:
    """
    try:
        response = genai.GenerativeModel(AI_MODEL_NAME).generate_content(prompt)
        json_str = response.text.strip().replace("```json", "").replace("```", "")
        analysis_result = json.loads(json_str)

        return {
            "title_sentiment": analysis_result.get("title_sentiment", "N/A"),
            "title_category": analysis_result.get("title_category", "N/A"),
            "body_sentiment": analysis_result.get("body_sentiment", "N/A"),
            "body_category": analysis_result.get("body_category", "N/A"),
        }
    except Exception as e:
        print(f"  - AI分析エラー: {e}")
        return default_response

def main():
    print("--- プログラム開始 ---")
    
    try:
        df_existing = pd.read_excel(EXCEL_FILE)
        existing_urls = set(df_existing['URL'])
        print(f"🔍 既存のExcelファイルを読み込みました。({len(existing_urls)}件の記事)")
    except FileNotFoundError:
        df_existing = pd.DataFrame()
        existing_urls = set()
        print("🔍 既存のExcelファイルが見つからないため、新規に作成します。")

    articles = get_yahoo_news_with_selenium(KEYWORD)
    if not articles:
        print("⚠️ 処理対象の記事が見つかりませんでした。プログラムを終了します。")
        return

    new_articles_data = []
    for i, article in enumerate(articles):
        if article['URL'] in existing_urls:
            continue
        
        print(f"\n--- 新規記事の処理 ({i+1}/{len(articles)}): {article['タイトル']} ---")
        
        print("  - 本文をページごとに取得中...")
        body_pages = fetch_article_pages(article['URL'])
        print(f"  - 本文を {len(body_pages)} ページ分取得しました。")
        
        print("  - コメントをページごとに取得中...")
        comments_count, comment_pages = fetch_comments_with_selenium(article['URL'])
        print(f"  - 合計 {comments_count} 件のコメントを {len(comment_pages)} ページ分取得しました。")

        print("  - AIによる分析を実行中 (API消費量最適化版)...")
        first_page_body = body_pages[0] if body_pages else ""
        analysis = analyze_article_with_ai(article['タイトル'], first_page_body)
        
        new_article_dict = {
            'タイトル': article['タイトル'],
            'URL': article['URL'],
            '投稿日': article['投稿日'],
            '引用元': article['引用元'],
            'コメント数': comments_count,
            'タイトルからポジネガ判定': analysis.get('title_sentiment'),
            'タイトルからカテゴリ分け': analysis.get('title_category'),
            'ニュース記事本文からポジネガ判定': analysis.get('body_sentiment'),
            'ニュース記事本文からカテゴリ分け': analysis.get('body_category'),
        }

        for idx, page_content in enumerate(body_pages):
            new_article_dict[f'本文({idx+1}ページ目)'] = page_content
        
        for idx, page_content in enumerate(comment_pages):
            new_article_dict[f'閲覧者コメント({idx+1}ページ目)'] = page_content
            
        new_articles_data.append(new_article_dict)
        time.sleep(1)

    if new_articles_data:
        df_new = pd.DataFrame(new_articles_data)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        
        static_columns = [
            'タイトル', 'URL', '投稿日', '引用元', 'コメント数',
            'タイトルからポジネガ判定', 'タイトルからカテゴリ分け',
            'ニュース記事本文からポジネガ判定', 'ニュース記事本文からカテゴリ分け'
        ]
        
        body_columns = sorted([col for col in df_combined.columns if col.startswith('本文(')])
        comment_columns = sorted([col for col in df_combined.columns if col.startswith('閲覧者コメント(')])
        
        final_column_order = static_columns + body_columns + comment_columns
        
        final_df = df_combined[[col for col in final_column_order if col in df_combined.columns]]
        
        final_df.to_excel(EXCEL_FILE, index=False)
        print(f"\n✅ {len(new_articles_data)}件の新規記事を '{EXCEL_FILE}' に追記しました。")
    else:
        print("\n⚠️ 追記すべき新しい記事はありませんでした。")

    print("--- プログラム終了 ---")

if __name__ == "__main__":
    main()
