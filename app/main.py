from flask import Flask, request, render_template, send_file, session
import requests
from bs4 import BeautifulSoup
import pandas as pd
import io
import os
import tempfile

app = Flask(__name__)
app.secret_key = os.urandom(24)  # セッション用の秘密鍵

def get_reviews(url):
    reviews = []
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }

    # 基本URL部分を切り取る（/1.1の前まで）
    base_url = url.split('/1.1')[0]  # '/1.1'でURLを分割し、基本URL部分を取得

    page_num = 1  # 最初のページ番号から開始
    while True:
        page_url = f"{base_url}?p={page_num}"  # ページ番号をURLに追加
        print(f"Fetching reviews from: {page_url}")

        try:
            # HTTPリクエストでページを取得
            response = requests.get(page_url, headers=headers)
            response.raise_for_status()  # HTTPステータスコードエラーを検出
            soup = BeautifulSoup(response.text, 'html.parser')

            # レビュー要素を抽出 (適切なクラス名を使用する)
            review_elements = soup.find_all("div", class_="review-body--1pESv")
            if not review_elements:
                print(f"No reviews found on page {page_num}.")
                break  # 次のページがない場合は終了

            # レビューをリストに格納
            for review in review_elements:
                reviews.append(review.text.strip())

            page_num += 1  # 次のページへ進む

        except requests.exceptions.RequestException as e:
            print(f"Error fetching reviews: {e}")
            break  # エラーが発生した場合も終了

    return reviews

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/fetch', methods=['POST'])
def fetch_reviews():
    url = request.form.get('url')

    # 商品タイトルを取得
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    product_title = soup.find("title").text.strip()

    # レビュー取得
    reviews = get_reviews(url)
    if not reviews:
        return "レビューが見つかりませんでした。", 400

    # データをExcelに変換
    data = [{"No": i, "レビュー内容": review} for i, review in enumerate(reviews, start=1)]
    df = pd.DataFrame(data)

    # 一時ファイルにExcelを保存
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    with pd.ExcelWriter(temp_file.name, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="レビュー")

    # セッションにファイル名を保存
    session['excel_file'] = temp_file.name

    # 結果をテンプレートに渡して表示
    return render_template('result.html', reviews=reviews, product_title=product_title)

@app.route('/download', methods=['POST'])
def download_excel():
    # セッションからファイル名を取得
    excel_file = session.get('excel_file')
    if not excel_file or not os.path.exists(excel_file):
        return "エラー: Excelファイルが見つかりません。", 400

    return send_file(
        excel_file,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="reviews.xlsx"
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

