import sys
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QFileDialog, QScrollArea,
    QWidget, QVBoxLayout, QLabel, QMessageBox
)
from PyQt6.QtGui import QPixmap, QPainter, QAction
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
from PyQt6.QtCore import Qt, QTimer, QRect

class PdfViewerMiya(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Viewer")
        self.resize(900, 800)
        
        self.doc = None
        self.zoom = 1.0 # デフォルトのズーム率
        self.page_labels = []

        self._init_ui()

    def _init_ui(self):
        # ツールバーの構築
        toolbar = QToolBar("Main Toolbar")
        self.addToolBar(toolbar)

        open_action = QAction("開く", self)
        open_action.triggered.connect(self.open_pdf)
        toolbar.addAction(open_action)

        print_action = QAction("印刷", self)
        print_action.triggered.connect(self.print_pdf)
        toolbar.addAction(print_action)

        toolbar.addSeparator()

        zoom_in_action = QAction("拡大", self)
        zoom_in_action.triggered.connect(self.zoom_in)
        toolbar.addAction(zoom_in_action)

        zoom_out_action = QAction("縮小", self)
        zoom_out_action.triggered.connect(self.zoom_out)
        toolbar.addAction(zoom_out_action)
        
        # スクロールエリア（メインビュー）
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background-color: #525659;")
        
        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        self.scroll_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.scroll_layout.setSpacing(10)
        self.scroll_layout.setContentsMargins(10, 10, 10, 10)
        
        self.scroll_area.setWidget(self.scroll_widget)
        self.setCentralWidget(self.scroll_area)

        # スクロールイベントにフック
        self.scroll_area.verticalScrollBar().valueChanged.connect(self.on_scroll)

    def open_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "PDFを開く", "", "PDF Files (*.pdf)")
        if file_path:
            self.load_pdf(file_path)

    def load_pdf(self, file_path):
        if self.doc:
            self.doc.close()
            self.doc = None
        
        try:
            self.doc = fitz.open(file_path)
            self.setWindowTitle(f"PDF Viewer - {file_path}")
            self.zoom = 1.5
            self.setup_pages()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"PDFの読み込みに失敗しました:\n{e}")

    def setup_pages(self):
        """プレースホルダーとしてQLabelを配置し、サイズだけを確保する"""
        # 既存のウィジェットを確実に削除（メモリリーク防止）
        for i in reversed(range(self.scroll_layout.count())): 
            item = self.scroll_layout.itemAt(i)
            if item.widget():
                item.widget().setParent(None)
                item.widget().deleteLater()
        self.page_labels.clear()

        if not self.doc or len(self.doc) == 0:
            return

        for page_num in range(len(self.doc)):
            label = QLabel()
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            # 視認性を上げるため枠線を追加、背景を白に
            label.setStyleSheet("background-color: white; border: 1px solid #cccccc;")
            label.setScaledContents(True) # 画像サイズをLabelに自動追従させる
            
            # 初期サイズの計算
            page = self.doc.load_page(page_num)
            rect = page.rect
            label.setFixedSize(int(rect.width * self.zoom), int(rect.height * self.zoom))
            
            self.scroll_layout.addWidget(label)
            self.page_labels.append(label)

        # QScrollAreaの内部ウィジェットのサイズを強制的に再計算させる
        self.scroll_widget.adjustSize()

        # GUIのレイアウト更新を確実に待ってから初期描画を行う
        QTimer.singleShot(50, self.on_scroll)

    def on_scroll(self):
        """画面に見えているページのみをレンダリングして高速化とメモリ節約"""
        if not self.doc:
            return
            
        # スクロールエリアのビューポートの矩形を取得
        viewport_rect = self.scroll_area.viewport().rect()
        
        for i, label in enumerate(self.page_labels):
            # ラベルの座標をビューポート基準の座標にマッピングして確実な交差判定を行う
            top_left = label.mapTo(self.scroll_area.viewport(), label.rect().topLeft())
            bottom_right = label.mapTo(self.scroll_area.viewport(), label.rect().bottomRight())
            mapped_rect = QRect(top_left, bottom_right)
            
            # ビューポートとラベルが少しでも重なっているか判定
            if viewport_rect.intersects(mapped_rect):
                if label.pixmap() is None or label.pixmap().isNull():
                    self.render_single_page(i, label)
            else:
                # 画面外のページは確実にメモリを解放して白紙化を防ぐ
                if label.pixmap() is not None and not label.pixmap().isNull():
                    label.setPixmap(QPixmap())

    def render_single_page(self, index, label):
        """1ページ分のレンダリング処理（真っ白になる不具合を根本解決）"""
        try:
            page = self.doc.load_page(index)
            mat = fitz.Matrix(self.zoom, self.zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            # QImageのメモリアライメント問題を回避するため、一度バイナリデータ(PPM)に
            # 変換してからQPixmapに読み込ませる。これにより確実に画像が表示される。
            img_data = pix.tobytes("ppm")
            pixmap = QPixmap()
            pixmap.loadFromData(img_data)
            
            label.setPixmap(pixmap)
        except Exception as e:
            print(f"Page {index} rendering error: {e}")

    def update_sizes(self):
        """ズーム変更時のサイズ再計算"""
        if not self.doc:
            return
        for i, label in enumerate(self.page_labels):
            page = self.doc.load_page(i)
            rect = page.rect
            label.setFixedSize(int(rect.width * self.zoom), int(rect.height * self.zoom))
            label.setPixmap(QPixmap()) # ズーム変更時は全ページのキャッシュを破棄
            
        self.scroll_widget.adjustSize()
        QTimer.singleShot(50, self.on_scroll)

    def zoom_in(self):
        # 際限ない拡大によるクラッシュを防止
        if self.zoom < 5.0:
            self.zoom *= 1.2
            self.update_sizes()

    def zoom_out(self):
        # 際限ない縮小によるクラッシュを防止
        if self.zoom > 0.2:
            self.zoom /= 1.2
            self.update_sizes()

    def print_pdf(self):
        """一般的なネイティブの印刷ダイアログと印刷処理"""
        if not self.doc:
            return
            
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self)
        
        if dialog.exec() == QPrintDialog.DialogCode.Accepted:
            painter = QPainter()
            painter.begin(printer)
            
            for i in range(len(self.doc)):
                if i > 0:
                    printer.newPage()
                
                page = self.doc.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False) 
                
                # 印刷時も同様のバイナリ経由で確実に出力
                img_data = pix.tobytes("ppm")
                pixmap = QPixmap()
                pixmap.loadFromData(img_data)
                
                rect = painter.viewport()
                size = pixmap.size()
                size.scale(rect.size(), Qt.AspectRatioMode.KeepAspectRatio)
                painter.setViewport(rect.x(), rect.y(), size.width(), size.height())
                painter.setWindow(pixmap.rect())
                
                painter.drawPixmap(0, 0, pixmap)
                
            painter.end()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = PdfViewerMiya()
    viewer.show()
    sys.exit(app.exec())