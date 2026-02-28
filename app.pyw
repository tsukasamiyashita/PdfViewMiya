import sys
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QFileDialog, QScrollArea,
    QWidget, QVBoxLayout, QLabel, QMessageBox, QSplitter, QTreeWidget,
    QTreeWidgetItem, QSpinBox, QHBoxLayout, QFrame
)
from PyQt6.QtGui import QPixmap, QPainter, QAction, QKeySequence, QCursor
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
from PyQt6.QtCore import Qt, QTimer, QRect, QEvent

class PdfEditMiya(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PdfEditMiya")
        # 最大化解除時のデフォルトサイズとして保持
        self.resize(1200, 800) 
        
        # 状態管理
        self.doc = None
        self.zoom = 1.5
        self.rotation = 0
        self.page_labels = []
        self.current_page = 0
        self.is_updating_ui = False
        self.is_single_page_mode = False
        
        # 自動ズーム追従モード ("width", "page", or None)
        self.auto_fit_mode = None

        # パン（ドラッグ移動）用の状態
        self.is_panning = False
        self.last_mouse_pos = None

        self._init_ui()

    def _init_ui(self):
        # メニューバーの構築
        self._init_menu()

        # ツールバーの構築
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False)
        self.addToolBar(self.toolbar)

        # ツールバー：ファイル操作
        open_action = QAction("📁 開く", self)
        open_action.setShortcut(QKeySequence.StandardKey.Open)
        open_action.triggered.connect(self.open_pdf)
        self.toolbar.addAction(open_action)

        print_action = QAction("🖨️ 印刷", self)
        print_action.setShortcut(QKeySequence.StandardKey.Print)
        print_action.triggered.connect(self.print_pdf)
        self.toolbar.addAction(print_action)

        self.toolbar.addSeparator()

        # ツールバー：表示モード（目次・単一ページ）
        self.toc_action = QAction("📑 目次", self)
        self.toc_action.setCheckable(True)
        self.toc_action.setChecked(False)
        self.toc_action.triggered.connect(self.toggle_toc)
        self.toolbar.addAction(self.toc_action)

        self.single_page_action = QAction("📄 単一ページ表示", self)
        self.single_page_action.setCheckable(True)
        self.single_page_action.setChecked(False)
        self.single_page_action.triggered.connect(self.toggle_single_page)
        self.toolbar.addAction(self.single_page_action)

        self.toolbar.addSeparator()

        # ツールバー：ページナビゲーション
        prev_action = QAction("◀ 前へ", self)
        prev_action.triggered.connect(self.prev_page)
        self.toolbar.addAction(prev_action)

        self.page_spinbox = QSpinBox()
        self.page_spinbox.setMinimum(1)
        self.page_spinbox.setMaximum(1)
        self.page_spinbox.setSuffix(" / 1")
        self.page_spinbox.setKeyboardTracking(False)
        self.page_spinbox.valueChanged.connect(self.jump_to_page)
        self.toolbar.addWidget(self.page_spinbox)

        next_action = QAction("次へ ▶", self)
        next_action.triggered.connect(self.next_page)
        self.toolbar.addAction(next_action)

        self.toolbar.addSeparator()

        # ツールバー：ズーム操作
        zoom_out_action = QAction("🔍- 縮小", self)
        zoom_out_action.triggered.connect(self.zoom_out)
        self.toolbar.addAction(zoom_out_action)

        zoom_in_action = QAction("🔍+ 拡大", self)
        zoom_in_action.triggered.connect(self.zoom_in)
        self.toolbar.addAction(zoom_in_action)

        actual_size_action = QAction("1:1 実際のサイズ", self)
        actual_size_action.triggered.connect(self.actual_size)
        self.toolbar.addAction(actual_size_action)

        fit_width_action = QAction("↔ 幅に合わせる", self)
        fit_width_action.triggered.connect(lambda: self.fit_to_width())
        self.toolbar.addAction(fit_width_action)

        fit_page_action = QAction("↕ ページに合わせる", self)
        fit_page_action.triggered.connect(lambda: self.fit_to_page())
        self.toolbar.addAction(fit_page_action)

        self.toolbar.addSeparator()

        # ツールバー：回転操作
        rotate_ccw_action = QAction("↺ 左回転", self)
        rotate_ccw_action.triggered.connect(self.rotate_ccw)
        self.toolbar.addAction(rotate_ccw_action)

        rotate_cw_action = QAction("↻ 右回転", self)
        rotate_cw_action.triggered.connect(self.rotate_cw)
        self.toolbar.addAction(rotate_cw_action)

        # メインレイアウト（Splitterで目次とビューアを分割）
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.setCentralWidget(self.splitter)

        # 左側：目次（TOC）ツリー
        self.toc_tree = QTreeWidget()
        self.toc_tree.setHeaderHidden(True)
        self.toc_tree.itemClicked.connect(self.on_toc_clicked)
        self.splitter.addWidget(self.toc_tree)
        self.toc_tree.setVisible(False)

        # 右側：スクロールエリア（メインビュー）
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background-color: #525659;")
        self.scroll_area.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.splitter.addWidget(self.scroll_area)
        
        # Splitterの初期比率
        self.splitter.setSizes([200, 800])

        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        self.scroll_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.scroll_layout.setSpacing(15)
        self.scroll_layout.setContentsMargins(20, 20, 20, 20)
        
        self.scroll_area.setWidget(self.scroll_widget)

        # スクロールイベントにフック
        self.scroll_area.verticalScrollBar().valueChanged.connect(self.on_scroll)
        self.scroll_area.horizontalScrollBar().valueChanged.connect(self.on_scroll)

        # イベントフィルターの設定
        self.scroll_area.installEventFilter(self)
        self.scroll_area.viewport().installEventFilter(self)
        self.scroll_area.viewport().setCursor(Qt.CursorShape.OpenHandCursor)

    def _init_menu(self):
        menubar = self.menuBar()
        help_menu = menubar.addMenu("ヘルプ(&H)")
        
        readme_action = QAction("Readme", self)
        readme_action.triggered.connect(self.show_readme)
        help_menu.addAction(readme_action)
        
        about_action = QAction("バージョン情報", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def show_readme(self):
        readme_text = """<h2>PdfEditMiya</h2>
        <p>起動と描画の高速化に特化し、スクロール遅延ロード（Lazy Loading）を採用したデスクトップPDFビューアです。</p>
        <h3>【主な機能】</h3>
        <ul>
        <li><b>マウスホイールによるページ切り替え</b>：単一ページ表示モード時、ページ最上部/最下部でのホイール操作で前後のページへジャンプ。</li>
        <li><b>A3・A4混在の自動サイズ調整</b>：単一ページ表示時、ページの物理サイズに合わせて自動ズーム。</li>
        <li><b>高度なスクロール操作</b>：Ctrl+ホイールでズーム、Shift+ホイールで横スクロール。上下矢印キーで高速スクロール。</li>
        <li><b>ドラッグ移動（パン機能）</b>：マウスの左クリックドラッグで直感的に画面を移動。</li>
        <li><b>その他</b>：目次表示、回転、各種ズーム、印刷など。</li>
        </ul>
        """
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Readme")
        msg_box.setTextFormat(Qt.TextFormat.RichText)
        msg_box.setText(readme_text)
        msg_box.exec()

    def show_about(self):
        QMessageBox.about(self, "バージョン情報", "<b>PdfEditMiya</b><br><br>バージョン: v1.0.0<br>Powered by PyQt6 & PyMuPDF")

    def eventFilter(self, obj, event):
        """マウスおよびキーボードによる高度なスクロール・ナビゲーション操作を処理する"""
        if obj in (self.scroll_area, self.scroll_area.viewport()):
            # 1. マウスホイール操作
            if event.type() == QEvent.Type.Wheel:
                modifiers = QApplication.keyboardModifiers()
                if modifiers == Qt.KeyboardModifier.ControlModifier:
                    # Ctrl + ホイール：ズーム
                    if event.angleDelta().y() > 0:
                        self.zoom_in()
                    else:
                        self.zoom_out()
                    return True
                elif modifiers == Qt.KeyboardModifier.ShiftModifier:
                    # Shift + ホイール：横スクロール
                    h_bar = self.scroll_area.horizontalScrollBar()
                    h_bar.setValue(h_bar.value() - event.angleDelta().y())
                    return True
                else:
                    # 修飾キーなしの通常のホイール操作
                    if self.is_single_page_mode:
                        v_bar = self.scroll_area.verticalScrollBar()
                        delta = event.angleDelta().y()
                        
                        # 単一ページモード時、ページ最上部で上スクロールすると「前へ」
                        if delta > 0: 
                            if v_bar.value() <= v_bar.minimum():
                                self.prev_page()
                                return True
                        # 単一ページモード時、ページ最下部で下スクロールすると「次へ」
                        elif delta < 0: 
                            if v_bar.value() >= v_bar.maximum():
                                self.next_page()
                                return True

            # 2. キーボードショートカット
            elif event.type() == QEvent.Type.KeyPress:
                v_bar = self.scroll_area.verticalScrollBar()
                
                if event.key() == Qt.Key.Key_Space:
                    modifiers = QApplication.keyboardModifiers()
                    if modifiers == Qt.KeyboardModifier.ShiftModifier:
                        if self.is_single_page_mode and v_bar.value() == v_bar.minimum():
                            self.prev_page()
                        else:
                            v_bar.setValue(v_bar.value() - v_bar.pageStep())
                    else:
                        if self.is_single_page_mode and v_bar.value() == v_bar.maximum():
                            self.next_page()
                        else:
                            v_bar.setValue(v_bar.value() + v_bar.pageStep())
                    return True
                
                elif event.key() == Qt.Key.Key_Up:
                    v_bar.setValue(v_bar.value() - 150)
                    return True
                    
                elif event.key() == Qt.Key.Key_Down:
                    v_bar.setValue(v_bar.value() + 150)
                    return True
                    
                elif event.key() == Qt.Key.Key_Left:
                    self.prev_page()
                    return True
                    
                elif event.key() == Qt.Key.Key_Right:
                    self.next_page()
                    return True

            # 3. マウスのドラッグによる画面移動（パン）
            elif event.type() == QEvent.Type.MouseButtonPress:
                if event.button() == Qt.MouseButton.LeftButton:
                    self.is_panning = True
                    self.last_mouse_pos = event.position()
                    self.scroll_area.viewport().setCursor(Qt.CursorShape.ClosedHandCursor)
                    return True
                    
            elif event.type() == QEvent.Type.MouseMove:
                if self.is_panning and self.last_mouse_pos is not None:
                    delta = event.position() - self.last_mouse_pos
                    h_bar = self.scroll_area.horizontalScrollBar()
                    h_bar.setValue(int(h_bar.value() - delta.x()))
                    v_bar.setValue(int(v_bar.value() - delta.y()))
                    self.last_mouse_pos = event.position()
                    return True
                    
            elif event.type() == QEvent.Type.MouseButtonRelease:
                if event.button() == Qt.MouseButton.LeftButton:
                    self.is_panning = False
                    self.last_mouse_pos = None
                    self.scroll_area.viewport().setCursor(Qt.CursorShape.OpenHandCursor)
                    return True
                    
        return super().eventFilter(obj, event)

    def open_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "PDFを開く", "", "PDF Files (*.pdf)")
        if file_path:
            self.load_pdf(file_path)

    def load_pdf(self, file_path):
        if self.doc is not None:
            self.doc.close()
            self.doc = None
        
        try:
            self.doc = fitz.open(file_path)
            self.setWindowTitle(f"PdfEditMiya - {file_path}")
            self.zoom = 1.5
            self.rotation = 0
            self.current_page = 0
            self.auto_fit_mode = None
            
            self.is_updating_ui = True
            self.page_spinbox.setMinimum(1)
            self.page_spinbox.setMaximum(self.doc.page_count)
            self.page_spinbox.setSuffix(f" / {self.doc.page_count}")
            self.page_spinbox.setValue(1)
            self.is_updating_ui = False

            self.load_toc()
            self.setup_pages()
            self.scroll_area.setFocus()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"PDFの読み込みに失敗しました:\n{e}")

    def load_toc(self):
        self.toc_tree.clear()
        if not self.doc:
            return
            
        toc = self.doc.get_toc()
        if not toc:
            item = QTreeWidgetItem(["目次がありません"])
            self.toc_tree.addTopLevelItem(item)
            return

        items = {}
        for level, title, page in toc:
            item = QTreeWidgetItem([title])
            item.setData(0, Qt.ItemDataRole.UserRole, page)

            if level == 1:
                self.toc_tree.addTopLevelItem(item)
            else:
                parent_level = level - 1
                while parent_level > 0 and parent_level not in items:
                    parent_level -= 1
                if parent_level in items:
                    items[parent_level].addChild(item)
                else:
                    self.toc_tree.addTopLevelItem(item)
            items[level] = item
        
        self.toc_tree.expandAll()

    def on_toc_clicked(self, item, column):
        page = item.data(0, Qt.ItemDataRole.UserRole)
        if page is not None:
            self.jump_to_page(page)

    def toggle_toc(self, checked):
        self.toc_tree.setVisible(checked)

    def toggle_single_page(self, checked):
        self.is_single_page_mode = checked
        self.update_page_visibility()
        
        if checked:
            if self.auto_fit_mode == "width":
                self.fit_to_width(auto=True)
            else:
                self.fit_to_page(auto=True)
                
            self.scroll_area.verticalScrollBar().setValue(0)
            self.scroll_area.horizontalScrollBar().setValue(0)
        else:
            self.apply_transformations()
            self.jump_to_page(self.current_page + 1)

    def update_page_visibility(self):
        for i, label in enumerate(self.page_labels):
            if self.is_single_page_mode:
                label.setVisible(i == self.current_page)
            else:
                label.setVisible(True)

    def setup_pages(self):
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        
        self.page_labels.clear()

        if not self.doc or self.doc.page_count == 0:
            return

        for page_num in range(self.doc.page_count):
            label = QLabel()
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            label.setStyleSheet("""
                background-color: white; 
                border: 1px solid #999;
            """)
            label.setScaledContents(True)
            
            self.scroll_layout.addWidget(label)
            self.page_labels.append(label)

        self.update_page_visibility()
        self.apply_transformations()

    def apply_transformations(self):
        if not self.doc or self.doc.is_closed:
            return

        for i, label in enumerate(self.page_labels):
            if i >= self.doc.page_count:
                break
            
            if self.is_single_page_mode and i != self.current_page:
                label.setFixedSize(0, 0)
                label.setPixmap(QPixmap())
                continue
            
            page = self.doc.load_page(i)
            mat = fitz.Matrix(self.zoom, self.zoom).prerotate(self.rotation)
            rect = page.rect.transform(mat)
            
            label.setFixedSize(int(rect.width), int(rect.height))
            label.setPixmap(QPixmap())
            
        self.scroll_widget.adjustSize()
        QTimer.singleShot(50, self.on_scroll)

    def on_scroll(self):
        if not self.doc or self.doc.is_closed:
            return
            
        viewport_rect = self.scroll_area.viewport().rect()
        
        if self.is_single_page_mode:
            if 0 <= self.current_page < len(self.page_labels):
                label = self.page_labels[self.current_page]
                if label.pixmap() is None or label.pixmap().isNull():
                    self.render_single_page(self.current_page, label)
            return

        center_y = viewport_rect.center().y()
        current_visible_page = self.current_page
        min_distance = float('inf')
        
        for i, label in enumerate(self.page_labels):
            if label is None or not label.isVisible():
                continue

            top_left = label.mapTo(self.scroll_area.viewport(), label.rect().topLeft())
            bottom_right = label.mapTo(self.scroll_area.viewport(), label.rect().bottomRight())
            mapped_rect = QRect(top_left, bottom_right)
            
            dist = abs(mapped_rect.center().y() - center_y)
            if dist < min_distance:
                min_distance = dist
                current_visible_page = i

            if viewport_rect.intersects(mapped_rect):
                if label.pixmap() is None or label.pixmap().isNull():
                    self.render_single_page(i, label)
            else:
                if label.pixmap() is not None and not label.pixmap().isNull():
                    label.setPixmap(QPixmap())

        if current_visible_page != self.current_page:
            self.current_page = current_visible_page
            self.is_updating_ui = True
            self.page_spinbox.setValue(self.current_page + 1)
            self.is_updating_ui = False

    def render_single_page(self, index, label):
        if not self.doc or self.doc.is_closed or index >= self.doc.page_count:
            return

        try:
            page = self.doc.load_page(index)
            mat = fitz.Matrix(self.zoom, self.zoom).prerotate(self.rotation)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            img_data = pix.tobytes("ppm")
            pixmap = QPixmap()
            pixmap.loadFromData(img_data)
            
            label.setPixmap(pixmap)
        except Exception as e:
            print(f"Page {index} rendering error: {e}")

    def jump_to_page(self, page_num):
        if self.is_updating_ui or not self.doc or not self.page_labels:
            return
        
        target_index = page_num - 1
        if 0 <= target_index < len(self.page_labels):
            if self.is_single_page_mode:
                self.current_page = target_index
                self.update_page_visibility()
                
                if self.auto_fit_mode == "width":
                    self.fit_to_width(auto=True)
                else:
                    self.fit_to_page(auto=True)
                    
                self.scroll_area.verticalScrollBar().setValue(0)
                self.scroll_area.horizontalScrollBar().setValue(0)
            else:
                target_label = self.page_labels[target_index]
                y_pos = target_label.y()
                self.scroll_area.verticalScrollBar().setValue(y_pos)

    def next_page(self):
        val = self.page_spinbox.value()
        if val < self.page_spinbox.maximum():
            self.page_spinbox.setValue(val + 1)

    def prev_page(self):
        val = self.page_spinbox.value()
        if val > 1:
            self.page_spinbox.setValue(val - 1)

    def zoom_in(self):
        self.auto_fit_mode = None
        if self.zoom < 5.0:
            self.zoom *= 1.2
            self.apply_transformations()

    def zoom_out(self):
        self.auto_fit_mode = None
        if self.zoom > 0.2:
            self.zoom /= 1.2
            self.apply_transformations()

    def actual_size(self):
        if not self.doc or self.doc.is_closed:
            return
        self.auto_fit_mode = None
        self.zoom = 1.0
        self.apply_transformations()

    def fit_to_width(self, auto=False):
        if not self.doc or self.doc.is_closed or not self.page_labels:
            return
            
        if not auto:
            self.auto_fit_mode = "width"
        
        page = self.doc.load_page(self.current_page)
        viewport_width = self.scroll_area.viewport().width() - 40 
        
        mat = fitz.Matrix(1.0, 1.0).prerotate(self.rotation)
        base_width = page.rect.transform(mat).width
        
        new_zoom = viewport_width / base_width
        if 0.2 <= new_zoom <= 5.0:
            self.zoom = new_zoom
            self.apply_transformations()

    def fit_to_page(self, auto=False):
        if not self.doc or self.doc.is_closed or not self.page_labels:
            return
            
        if not auto:
            self.auto_fit_mode = "page"
        
        page = self.doc.load_page(self.current_page)
        viewport_width = self.scroll_area.viewport().width() - 40 
        viewport_height = self.scroll_area.viewport().height() - 40 
        
        mat = fitz.Matrix(1.0, 1.0).prerotate(self.rotation)
        base_rect = page.rect.transform(mat)
        
        zoom_w = viewport_width / base_rect.width
        zoom_h = viewport_height / base_rect.height
        
        new_zoom = min(zoom_w, zoom_h)
        if 0.2 <= new_zoom <= 5.0:
            self.zoom = new_zoom
            self.apply_transformations()

    def rotate_cw(self):
        self.rotation = (self.rotation + 90) % 360
        self.apply_transformations()

    def rotate_ccw(self):
        self.rotation = (self.rotation - 90) % 360
        self.apply_transformations()

    def print_pdf(self):
        if not self.doc or self.doc.is_closed:
            return
            
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self)
        
        if dialog.exec() == QPrintDialog.DialogCode.Accepted:
            painter = QPainter()
            painter.begin(printer)
            
            for i in range(self.doc.page_count):
                if i > 0:
                    printer.newPage()
                
                page = self.doc.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False) 
                
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
    viewer = PdfEditMiya()
    viewer.showMaximized()
    sys.exit(app.exec())