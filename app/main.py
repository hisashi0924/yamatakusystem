from flask import Flask, request, render_template, send_file
import requests
from bs4 import BeautifulSoup
import pandas as pd
import io

app = Flask(__name__)

def get_reviews(url):
    reviews = []
    page = 1
    max_pages = 100  # 上限を設定（例: 最大100ページ）

    while page <= max_pages:
        # ページ番号を含むURLを生成
        page_url = f"{url.split('?')[0]}?p={page}"  # ?以降をリセットし、p=pageを追加
        print(f"Fetching: {page_url}")  # デバッグ用メッセージ
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(page_url, headers=headers)

        # サーバー応答の確認
        if response.status_code != 200:
            print(f"Error: HTTP {response.status_code}")
            break

        soup = BeautifulSoup(response.text, 'html.parser')

        # レビューを取得
        review_elements = soup.find_all("div", class_="review-body--1pESv")
        if not review_elements:
            print(f"No reviews found on page {page}. Stopping.")
            break  # レビューが見つからない場合、終了

        for review in review_elements:
            content = review.text.strip()
            reviews.append(content)

        print(f"Page {page} processed, {len(review_elements)} reviews found.")
        page += 1  # 次のページへ進む

    print(f"Total reviews fetched: {len(reviews)}")
    return reviews


@app.route('/')
def index():
    # トップページを表示
    return render_template('index.html')

@app.route('/download', methods=['POST'])
def download():
    # ユーザーから送信されたURLを取得
    url = request.form.get('url')

    # 商品タイトルを取得
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    product_title = soup.find("title").text.strip()

    # レビューを取得
    reviews = get_reviews(url)
    if not reviews:
        return "レビューが見つかりませんでした。URLやHTML構造を確認してください。", 400

    # データをExcelに変換
    data = [{"No": i, "レビュー内容": review} for i, review in enumerate(reviews, start=1)]
    df = pd.DataFrame(data)
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="レビュー")
    excel_buffer.seek(0)

    # Excelを返却
    return send_file(
        excel_buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"{product_title}_reviews.xlsx"
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

