# 対話式PDF分割スクリプト
# PyPDF2を使ってPDFファイルの特定ページを抽出します
# 前回の入力値を記憶する機能付き

import os
import json
from PyPDF2 import PdfReader, PdfWriter

# 環境変数から設定を取得するか、デフォルト値を使用
DEFAULT_INPUT_PDF = ""  # 空の文字列をデフォルトに
DEFAULT_OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Documents")  # ユーザーのドキュメントフォルダをデフォルトに

# 設定ファイルのパス - ユーザーのホームディレクトリに保存
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".pdf_split_config.json")

def load_config():
    """設定ファイルから前回の設定を読み込む"""
    default_config = {
        "input_pdf": DEFAULT_INPUT_PDF,
        "output_dir": DEFAULT_OUTPUT_DIR,
        "last_start_page": 1,
        "last_end_page": 10
    }
    
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            return config
        else:
            return default_config
    except Exception as e:
        print(f"設定ファイルの読み込みに失敗しました: {e}")
        return default_config

def save_config(config):
    """設定ファイルに現在の設定を保存する"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        print("設定を保存しました")
    except Exception as e:
        print(f"設定ファイルの保存に失敗しました: {e}")

def split_pdf(input_path, output_path, start_page, end_page):
    """
    PDFファイルを指定したページ範囲で分割します
    
    :param input_path: 入力PDFファイルのパス
    :param output_path: 出力PDFファイルのパス
    :param start_page: 開始ページ (1から始まる)
    :param end_page: 終了ページ (1から始まる)
    """
    # PDFファイルを読み込む
    pdf_reader = PdfReader(input_path)
    
    # ページ番号をチェック
    total_pages = len(pdf_reader.pages)
    print(f"元のPDFファイルの総ページ数: {total_pages}")
    
    if start_page < 1 or end_page > total_pages or start_page > end_page:
        print("ページ範囲が無効です")
        return False
    
    # 新しいPDFライターを作成
    pdf_writer = PdfWriter()
    
    # 指定されたページ範囲を追加 (PyPDF2は0から始まるインデックスを使用)
    for page_num in range(start_page - 1, end_page):
        pdf_writer.add_page(pdf_reader.pages[page_num])
    
    # 出力ディレクトリが存在するか確認し、なければ作成
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 新しいPDFファイルを書き込む
    with open(output_path, "wb") as output_file:
        pdf_writer.write(output_file)
    
    print(f"PDFファイルの分割が完了しました: {output_path}")
    print(f"ページ範囲: {start_page}〜{end_page} (合計: {end_page - start_page + 1}ページ)")
    return True

def get_valid_integer(prompt, min_value=1, max_value=None):
    """ユーザーから有効な整数を取得する"""
    while True:
        try:
            value = int(input(prompt))
            if value < min_value:
                print(f"{min_value}以上の値を入力してください。")
                continue
            if max_value is not None and value > max_value:
                print(f"{max_value}以下の値を入力してください。")
                continue
            return value
        except ValueError:
            print("有効な数値を入力してください。")

if __name__ == "__main__":
    print("===== 対話式PDF分割ツール (設定記憶機能付き) =====")
    
    # 前回の設定を読み込む
    config = load_config()
    default_input_pdf = config["input_pdf"]
    default_output_dir = config["output_dir"]
    default_start_page = config["last_start_page"]
    default_end_page = config["last_end_page"]
    
    # 入力ファイルの確認
    print(f"\n前回使用した入力PDFファイル: {default_input_pdf}")
    change_input = input("入力ファイルを変更しますか？ (y/n、デフォルトはn): ").lower()
    
    if change_input == 'y':
        input_pdf = input("新しい入力PDFファイルのパスを入力してください: ")
    else:
        input_pdf = default_input_pdf
    
    # PDFファイルの情報を取得
    try:
        pdf_reader = PdfReader(input_pdf)
        total_pages = len(pdf_reader.pages)
        print(f"\n元のPDFファイルの総ページ数: {total_pages}")
    except Exception as e:
        print(f"エラー: PDFファイルを読み込めませんでした。{e}")
        input("\n終了するには Enter キーを押してください...")
        exit(1)
    
    # ページ範囲の入力
    print("\nページ範囲を指定してください。")
    print(f"前回使用した範囲: {default_start_page}〜{default_end_page}")
    
    start_page = get_valid_integer(f"開始ページ (1〜{total_pages}, デフォルト:{default_start_page}): ", 1, total_pages)
    end_page = get_valid_integer(f"終了ページ ({start_page}〜{total_pages}, デフォルト:{max(default_end_page, start_page)}): ", start_page, total_pages)
    
    # 出力ディレクトリの確認
    print(f"\n前回使用した出力ディレクトリ: {default_output_dir}")
    change_output_dir = input("出力ディレクトリを変更しますか？ (y/n、デフォルトはn): ").lower()
    
    if change_output_dir == 'y':
        output_dir = input("新しい出力ディレクトリのパスを入力してください: ")
    else:
        output_dir = default_output_dir
    
    # 出力ファイル名の生成
    default_output_name = f"分割_{start_page}-{end_page}.pdf"
    output_pdf = os.path.join(output_dir, default_output_name)
    
    print(f"\n出力PDFファイル: {output_pdf}")
    change_output_name = input("出力ファイル名を変更しますか？ (y/n、デフォルトはn): ").lower()
    
    if change_output_name == 'y':
        output_name = input("新しい出力ファイル名を入力してください: ")
        if not output_name.lower().endswith('.pdf'):
            output_name += '.pdf'
        output_pdf = os.path.join(output_dir, output_name)
    
    # PDFを分割
    print("\nPDFファイルを分割しています...")
    success = split_pdf(input_pdf, output_pdf, start_page, end_page)
    
    if success:
        print("\n処理が成功しました！")
        # 設定を保存
        config["input_pdf"] = input_pdf
        config["output_dir"] = output_dir
        config["last_start_page"] = start_page
        config["last_end_page"] = end_page
        save_config(config)
    else:
        print("\n処理中にエラーが発生しました。")
    
    input("\n終了するには Enter キーを押してください...")
