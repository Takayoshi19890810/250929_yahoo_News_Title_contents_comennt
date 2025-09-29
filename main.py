import os
import json
import time
import re
import pandas as pd
from datetime import datetime, timedelta
import google.generativeai as genai

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import requests

# --- 設定項目 ---
KEYWORD = "日産"  # 検索したいキーワード
EXCEL_FILE = "yahoo_news_analysis.xlsx"  # 保存するExcelファイル名
MAX_BODY_PAGES = 5  # 記事本文の最大取得ページ数
MAX_TOTAL_COMMENTS = 100 # 取得するコメントの最大数（負荷軽減のため）


# --- AI分析の設定 ---
try:
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
    if not GEMINI_API_KEY:
        print("警告: 環境変数 'GEMINI_API_KEY' が設定されていません。AI分析はスキップされます。")
        AI_ENABLED = False
    else:
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel('gemini-pro')
        AI_ENABLED = True
except Exception as e:
    print(f"AIモデルの初期化に失敗しました: {e}")
    AI_ENABLED = False

# --- 共通関数 (main1.py, main2.pyより引用・改変) ---

def format_datetime_str(dt_obj):
    """datetimeオブジェクトを 'yyyy/mm/dd hh:mm' 形式の文字列に変換"""
    return dt_obj.strftime("%Y/%m/%d %H:%M")

def get_yahoo_news_with_selenium(keyword: str) -> list[dict]:
    """Yahooニュースから記事の基本情報リストを取得する"""
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
        time.sleep(3) # ページ読み込み待機

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
            
            # 日時フォーマットを統一
            try:
                # "2024/9/29(月) 16:35" のような形式をパース
                cleaned_date_str = re.sub(r'\([月火水木金土日]\)', '', date_str)
                dt_obj = datetime.strptime(cleaned_date_str, "%Y/%m/%d %H:%M")
                formatted_date = format_datetime_str(dt_obj)
            except ValueError:
                formatted_date = date_str # パース失敗時は元の文字列を保持

            if title and url:
                articles_data.append({
                    "タイトル": title, "URL": url, "投稿日": formatted_date, "引用元": source
                })
    finally:
        driver.quit()
        
    print(f"✅ Yahoo!ニュースから {len(articles_data)} 件の記事情報を取得しました。")
    return articles_data

def fetch_article_pages(base_url: str) -> str:
    """記事のURLから本文を取得する"""
    full_text = []
    for page in range(1, MAX_BODY_PAGES + 1):
        url = base_url if page == 1 else f"{base_url}?page={page}"
        try:
            res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
            res.raise_for_status()
            soup = BeautifulSoup(res.text, "html.parser")
            
            # 本文が含まれていそうな領域を探す
            article_body = soup.find("div", class_=re.compile("article_body"))
            if not article_body:
                continue

            # pタグのテキストを抽出・結合
            page_text = "\n".join(p.get_text(strip=True) for p in article_body.find_all("p"))
            
            if not page_text or (full_text and page_text == full_text[-1]):
                break # ページが空か、前のページと同じ内容なら終了
            
            full_text.append(page_text)
            time.sleep(1) # サーバー負荷軽減
        except requests.RequestException:
            break
    return "\n\n".join(full_text)

def fetch_comments_with_selenium(base_url: str) -> list[str]:
    """記事のコメントを取得する"""
    comments = []
    if "/articles/" not in base_url:
        return comments

    comment_url = base_url.split("/articles/")[0] + "/comments/" + base_url.split("/articles/")[1]
    
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        driver.get(comment_url)
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # コメントのテキスト部分を抽出
        comment_tags = soup.select("p[class*='comment']")
        comments = [tag.get_text(strip=True) for tag in comment_tags if tag.get_text(strip=True)]
        
        # 上限数で切り捨て
        comments = comments[:MAX_TOTAL_COMMENTS]
    except Exception as e:
        print(f"  - コメント取得中にエラー: {e}")
    finally:
        driver.quit()
        
    return comments

def analyze_text_with_ai(text: str) -> dict:
    """AIを使ってテキストの感情とカテゴリを分析する"""
    if not AI_ENABLED or not text or not isinstance(text, str) or len(text.strip()) < 10:
        return {"sentiment": "N/A", "category": "N/A"}

    prompt = f"""
    以下のニュース記事のテキストを分析し、結果をJSON形式で返してください。

    分析項目:
    1. sentiment: 全体の論調が「ポジティブ」「ネガティブ」「ニュートラル」のいずれか。
    2. category: 内容に最も合致するカテゴリを以下から1つ選択してください。
       ["新技術・研究開発", "経営・財務", "販売・マーケティング", "生産・品質", "人事・労務", "不祥事・訴訟", "その他"]

    テキスト:
    ---
    {text[:2000]} # テキストが長すぎる場合を考慮して先頭2000文字に制限
    ---
    
    JSON出力のみ記述してください:
    """
    try:
        response = model.generate_content(prompt)
        # レスポンスからJSON部分を抽出してパース
        json_str = response.text.strip().replace("```json", "").replace("```", "")
        return json.loads(json_str)
    except Exception as e:
        print(f"  - AI分析エラー: {e}")
        return {"sentiment": "Error", "category": "Error"}

def main():
    """メイン処理"""
    print("--- プログラム開始 ---")
    
    # 1. 既存のExcelファイルを読み込み、処理済みのURLリストを作成
    try:
        df_existing = pd.read_excel(EXCEL_FILE)
        existing_urls = set(df_existing['URL'])
        print(f"🔍 既存のExcelファイルを読み込みました。({len(existing_urls)}件の記事)")
    except FileNotFoundError:
        df_existing = pd.DataFrame()
        existing_urls = set()
        print("🔍 既存のExcelファイルが見つからないため、新規に作成します。")

    # 2. Yahooニュースから基本情報を取得
    articles = get_yahoo_news_with_selenium(KEYWORD)
    if not articles:
        print("⚠️ 新規記事が見つかりませんでした。プログラムを終了します。")
        return

    # 3. 新規記事のみを抽出し、詳細情報を取得・分析
    new_articles_data = []
    for i, article in enumerate(articles):
        if article['URL'] in existing_urls:
            continue
        
        print(f"\n--- 新規記事の処理 ({i+1}/{len(articles)}): {article['タイトル']} ---")
        
        # 詳細情報（本文とコメント）を取得
        print("  - 本文を取得中...")
        full_text = fetch_article_pages(article['URL'])
        print(f"  - 本文を {len(full_text)} 文字取得しました。")
        
        print("  - コメントを取得中...")
        comments = fetch_comments_with_selenium(article['URL'])
        print(f"  - {len(comments)} 件のコメントを取得しました。")

        # AIによる分析
        print("  - AIによる分析を実行中...")
        title_analysis = analyze_text_with_ai(article['タイトル'])
        body_analysis = analyze_text_with_ai(full_text)
        
        new_articles_data.append({
            'タイトル': article['タイトル'],
            'URL': article['URL'],
            '投稿日': article['投稿日'],
            '引用元': article['引用元'],
            '本文': full_text,
            'コメント数': len(comments),
            '閲覧者コメント': "\n---\n".join(comments), # コメントは改行で区切る
            'タイトルからポジネガ判定': title_analysis.get('sentiment', 'N/A'),
            'タイトルからカテゴリ分け': title_analysis.get('category', 'N/A'),
            'ニュース記事本文からポジネガ判定': body_analysis.get('sentiment', 'N/A'),
            'ニュース記事本文からカテゴリ分け': body_analysis.get('category', 'N/A'),
        })
        time.sleep(2) # APIへの連続アクセスを避ける

    # 4. 新規データがあればExcelに追記して保存
    if new_articles_data:
        df_new = pd.DataFrame(new_articles_data)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        
        # カラムの順序を定義
        column_order = [
            '投稿日', '引用元', 'タイトル', 'URL', 'コメント数', 
            'タイトルからポジネガ判定', 'タイトルからカテゴリ分け', 
            'ニュース記事本文からポジネガ判定', 'ニュース記事本文からカテゴリ分け', 
            '本文', '閲覧者コメント'
        ]
        # 存在するカラムのみで順序を再定義
        final_columns = [col for col in column_order if col in df_combined.columns]
        df_combined = df_combined[final_columns]
        
        df_combined.to_excel(EXCEL_FILE, index=False)
        print(f"\n✅ {len(new_articles_data)}件の新規記事を '{EXCEL_FILE}' に追記しました。")
    else:
        print("\n⚠️ 追記すべき新しい記事はありませんでした。")

    print("--- プログラム終了 ---")

if __name__ == "__main__":
    main()
