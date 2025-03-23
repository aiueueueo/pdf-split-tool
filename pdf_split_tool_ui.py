import sys
import os
import threading
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QPushButton, QLabel, 
                             QVBoxLayout, QHBoxLayout, QRadioButton, QButtonGroup, 
                             QLineEdit, QFileDialog, QMessageBox, QGroupBox, QFrame,
                             QProgressBar, QSpacerItem, QSizePolicy)
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QObject
from PyQt5.QtGui import QFont, QIcon
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image
import re

# シグナルクラス（スレッド間通信用）
class WorkerSignals(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    error = pyqtSignal(str)

class PDFSplitToolUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.input_pdf_path = ""
        self.output_dir_path = ""
        self.signals = WorkerSignals()
        self.signals.progress.connect(self.update_progress)
        self.signals.finished.connect(self.process_finished)
        self.signals.error.connect(self.process_error)

    def initUI(self):
        self.setWindowTitle("PDF分割・変換ツール")
        self.setMinimumSize(500, 400)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QPushButton {
                background-color: #4a86e8;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3a76d8;
            }
            QLabel {
                color: #333333;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #cccccc;
                border-radius: 4px;
            }
            QGroupBox {
                border: 1px solid #cccccc;
                border-radius: 4px;
                margin-top: 12px;
                font-weight: bold;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 5px;
            }
        """)

        # メインウィジェットとレイアウト
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        self.setCentralWidget(main_widget)

        # ファイル選択セクション
        file_group = QGroupBox("ファイル選択")
        file_layout = QVBoxLayout(file_group)
        
        select_file_btn = QPushButton("PDFファイルを選択")
        select_file_btn.clicked.connect(self.select_file)
        select_file_btn.setFixedHeight(35)
        
        self.file_path_label = QLabel("選択されたファイル: なし")
        self.file_path_label.setWordWrap(True)
        
        file_layout.addWidget(select_file_btn)
        file_layout.addWidget(self.file_path_label)
        main_layout.addWidget(file_group)
        
        # ページ範囲セクション
        page_group = QGroupBox("ページ範囲")
        page_layout = QVBoxLayout(page_group)
        
        page_range_label = QLabel("分割するページ範囲を入力してください (例: 1-5):")
        self.page_range_input = QLineEdit()
        self.page_range_input.setPlaceholderText("例: 1-5")
        
        page_layout.addWidget(page_range_label)
        page_layout.addWidget(self.page_range_input)
        main_layout.addWidget(page_group)
        
        # 出力形式セクション
        output_group = QGroupBox("出力形式")
        output_layout = QVBoxLayout(output_group)
        
        self.output_format_group = QButtonGroup(self)
        
        pdf_radio = QRadioButton("PDF")
        jpeg_radio = QRadioButton("JPEG")
        png_radio = QRadioButton("PNG")
        
        self.output_format_group.addButton(pdf_radio, 1)
        self.output_format_group.addButton(jpeg_radio, 2)
        self.output_format_group.addButton(png_radio, 3)
        
        pdf_radio.setChecked(True)
        
        output_layout.addWidget(pdf_radio)
        output_layout.addWidget(jpeg_radio)
        output_layout.addWidget(png_radio)
        main_layout.addWidget(output_group)
        
        # 出力ディレクトリセクション
        dir_group = QGroupBox("出力ディレクトリ")
        dir_layout = QVBoxLayout(dir_group)
        
        select_dir_btn = QPushButton("出力先フォルダを選択")
        select_dir_btn.clicked.connect(self.select_output_dir)
        select_dir_btn.setFixedHeight(35)
        
        self.dir_path_label = QLabel("選択されたフォルダ: なし")
        self.dir_path_label.setWordWrap(True)
        
        dir_layout.addWidget(select_dir_btn)
        dir_layout.addWidget(self.dir_path_label)
        main_layout.addWidget(dir_group)
        
        # プログレスバー
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # スペーサー
        spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        main_layout.addItem(spacer)
        
        # 実行ボタン
        execute_btn = QPushButton("実行")
        execute_btn.clicked.connect(self.execute)
        execute_btn.setFixedHeight(40)
        execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #4caf50;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        main_layout.addWidget(execute_btn)
        
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "PDFファイルを選択", "", "PDF Files (*.pdf)")
        if file_path:
            self.input_pdf_path = file_path
            self.file_path_label.setText(f"選択されたファイル: {os.path.basename(file_path)}")

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "出力先フォルダを選択")
        if dir_path:
            self.output_dir_path = dir_path
            self.dir_path_label.setText(f"選択されたフォルダ: {dir_path}")

    def execute(self):
        if not self.input_pdf_path:
            QMessageBox.warning(self, "警告", "PDFファイルを選択してください。")
            return
            
        if not self.output_dir_path:
            QMessageBox.warning(self, "警告", "出力先フォルダを選択してください。")
            return
            
        page_range = self.page_range_input.text().strip()
        if not page_range:
            QMessageBox.warning(self, "警告", "ページ範囲を入力してください。")
            return

        # ページ範囲のバリデーション
        if not self.validate_page_range(page_range):
            QMessageBox.warning(self, "警告", "無効なページ範囲です。正しい形式（例: 1-5, 7, 9-11）で入力してください。")
            return
            
        output_format_id = self.output_format_group.checkedId()
        output_format = {1: "PDF", 2: "JPEG", 3: "PNG"}.get(output_format_id)
        
        # プログレスバーの初期化
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        
        # 処理の開始
        thread = threading.Thread(
            target=self.process_pdf, 
            args=(self.input_pdf_path, self.output_dir_path, page_range, output_format)
        )
        thread.daemon = True
        thread.start()

    def validate_page_range(self, page_range):
        """ページ範囲の形式を検証する"""
        # 1-5, 7, 9-11 のような形式をチェック
        pattern = r'^(\d+(-\d+)?)(,\s*\d+(-\d+)?)*$'
        if not re.match(pattern, page_range):
            return False

        # 各範囲の開始ページと終了ページを確認
        ranges = page_range.split(',')
        for r in ranges:
            r = r.strip()
            if '-' in r:
                start, end = map(int, r.split('-'))
                if start > end or start < 1:
                    return False
            else:
                if int(r) < 1:
                    return False
        return True

    def process_pdf(self, input_path, output_dir, page_range, output_format):
        """PDFの処理を行う（分割または変換）"""
        try:
            # PDFを読み込む
            pdf_reader = PdfReader(input_path)
            total_pages = len(pdf_reader.pages)
            
            # ページ範囲をパース
            pages_to_extract = self.parse_page_range(page_range, total_pages)
            if not pages_to_extract:
                self.signals.error.emit("指定されたページ範囲が無効です。")
                return
                
            if output_format == "PDF":
                self.split_pdf(pdf_reader, output_dir, pages_to_extract)
            else:  # JPEG or PNG
                self.convert_pdf_to_image(input_path, output_dir, pages_to_extract, output_format)
                
            self.signals.finished.emit()
        except Exception as e:
            self.signals.error.emit(str(e))

    def parse_page_range(self, page_range, total_pages):
        """ページ範囲の文字列をページ番号のリストに変換する"""
        pages = []
        ranges = page_range.split(',')
        
        for r in ranges:
            r = r.strip()
            if '-' in r:
                start, end = map(int, r.split('-'))
                # ページ番号の範囲チェック
                start = max(1, min(start, total_pages))
                end = max(1, min(end, total_pages))
                pages.extend(range(start, end + 1))
            else:
                page = int(r)
                if 1 <= page <= total_pages:
                    pages.append(page)
                    
        return sorted(set(pages))  # 重複を削除してソート

    def split_pdf(self, pdf_reader, output_dir, pages):
        """PDFを指定したページで分割する"""
        output_pdf = PdfWriter()
        total_pages = len(pages)
        
        for i, page_num in enumerate(pages):
            # PyPDF2は0から始まるインデックスを使用
            output_pdf.add_page(pdf_reader.pages[page_num - 1])
            progress = int((i + 1) / total_pages * 100)
            self.signals.progress.emit(progress)
            
        # 出力ファイル名を生成
        base_name = os.path.basename(self.input_pdf_path)
        name_without_ext = os.path.splitext(base_name)[0]
        page_range_text = f"{min(pages)}-{max(pages)}" if len(pages) > 1 else str(pages[0])
        output_file = os.path.join(output_dir, f"{name_without_ext}_p{page_range_text}.pdf")
        
        # 出力ファイルに書き込む
        with open(output_file, 'wb') as f:
            output_pdf.write(f)

    def convert_pdf_to_image(self, input_path, output_dir, pages, format_name):
        """PDFを画像(JPEG/PNG)に変換する (PyMuPDFを使用)"""
        total_pages = len(pages)
        
        try:
            # PyMuPDFでPDFを開く
            doc = fitz.open(input_path)
            
            for i, page_num in enumerate(pages):
                if 1 <= page_num <= len(doc):
                    # ページをレンダリング (300 dpi)
                    page = doc.load_page(page_num - 1)  # 0-indexed
                    zoom = 300 / 72  # 300 dpi
                    matrix = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=matrix)
                    
                    # 出力ファイル名を生成
                    base_name = os.path.basename(input_path)
                    name_without_ext = os.path.splitext(base_name)[0]
                    output_file = os.path.join(output_dir, f"{name_without_ext}_p{page_num}.{format_name.lower()}")
                    
                    # 画像として保存
                    if format_name.upper() == "JPEG":
                        pix.save(output_file, "jpeg")
                    else:  # PNG
                        pix.save(output_file, "png")
                    
                    progress = int((i + 1) / total_pages * 100)
                    self.signals.progress.emit(progress)
            
            doc.close()
        except Exception as e:
            self.signals.error.emit(str(e))

    def update_progress(self, value):
        """プログレスバーを更新する"""
        self.progress_bar.setValue(value)

    def process_finished(self):
        """処理完了時の処理"""
        self.progress_bar.setValue(100)
        QMessageBox.information(self, "完了", "処理が完了しました。")

    def process_error(self, error_message):
        """エラー発生時の処理"""
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "エラー", f"処理中にエラーが発生しました：\n{error_message}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PDFSplitToolUI()
    window.show()
    sys.exit(app.exec_()) 