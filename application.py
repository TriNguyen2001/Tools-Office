import sys
import fitz  # PyMuPDF
import requests  # thêm thư viện requests
import webbrowser
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QListWidget, QMessageBox, QHBoxLayout, QLabel, QScrollArea
)
from PySide6.QtGui import QPixmap, QImage

APP_VERSION = "1.0.0"  # version hiện tại
VERSION_URL = "https://github.com/TriNguyen2001/Tools-Office/blob/main/version.json"  


class PDFProcessor:
    """Xử lý PDF: merge và lưu file"""
    def __init__(self):
        self.output_pdf = None

    def merge_and_save(self, files, output_path):
        if not files:
            raise ValueError("Danh sách file rỗng!")

        pdf_out = fitz.open()
        for f in files:
            pdf_in = fitz.open(f)
            pdf_out.insert_pdf(pdf_in)
            pdf_in.close()

        pdf_out.save(output_path)
        pdf_out.close()
        self.output_pdf = output_path
        return self.output_pdf


class PDFMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PeTricute tập code - PDF Merger")
        self.resize(1000, 700)

        self.processor = PDFProcessor()

        layout = QHBoxLayout(self)

        # Bên trái: danh sách file + nút điều khiển
        left_layout = QVBoxLayout()
        self.list_widget = QListWidget()
        left_layout.addWidget(self.list_widget)

        btn_select = QPushButton("Chọn file PDF")
        btn_select.clicked.connect(self.select_files)
        left_layout.addWidget(btn_select)

        hlayout = QHBoxLayout()
        btn_up = QPushButton("⬆️ Lên")
        btn_down = QPushButton("⬇️ Xuống")
        btn_remove = QPushButton("❌ Xoá")
        btn_up.clicked.connect(self.move_up)
        btn_down.clicked.connect(self.move_down)
        btn_remove.clicked.connect(self.remove_item)
        hlayout.addWidget(btn_up)
        hlayout.addWidget(btn_down)
        hlayout.addWidget(btn_remove)
        left_layout.addLayout(hlayout)

        btn_merge_save_preview = QPushButton("📑 Ghép + Lưu")
        btn_merge_save_preview.clicked.connect(self.merge_save_preview)
        left_layout.addWidget(btn_merge_save_preview)

        # nút kiểm tra update
        btn_update = QPushButton("🔄 Kiểm tra cập nhật")
        btn_update.clicked.connect(self.check_update)
        left_layout.addWidget(btn_update)

        layout.addLayout(left_layout, 2)

        # Bên phải: ScrollArea để hiển thị nhiều trang
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.container = QWidget()
        self.vbox_preview = QVBoxLayout(self.container)
        self.scroll.setWidget(self.container)
        layout.addWidget(self.scroll, 5)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Chọn file PDF", "", "PDF Files (*.pdf)"
        )
        if files:
            for f in files:
                self.list_widget.addItem(f)

    def move_up(self):
        row = self.list_widget.currentRow()
        if row > 0:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row - 1, item)
            self.list_widget.setCurrentRow(row - 1)

    def move_down(self):
        row = self.list_widget.currentRow()
        if row < self.list_widget.count() - 1 and row >= 0:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row + 1, item)
            self.list_widget.setCurrentRow(row + 1)

    def remove_item(self):
        row = self.list_widget.currentRow()
        if row >= 0:
            self.list_widget.takeItem(row)

    def merge_save_preview(self):
        if self.list_widget.count() == 0:
            QMessageBox.warning(self, "Lỗi", "Chưa chọn file PDF nào!")
            return

        files = [self.list_widget.item(i).text() for i in range(self.list_widget.count())]

        output_file, _ = QFileDialog.getSaveFileName(
            self, "Lưu file PDF sau khi ghép", "", "PDF Files (*.pdf)"
        )
        if not output_file:
            return

        try:
            final_pdf = self.processor.merge_and_save(files, output_file)
            QMessageBox.information(self, "Thành công", f"Đã ghép và lưu PDF tại:\n{final_pdf}")
            self.show_pdf(final_pdf)
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể ghép PDF:\n{e}")

    def show_pdf(self, pdf_path):
        # Xóa preview cũ
        while self.vbox_preview.count():
            item = self.vbox_preview.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        # Render tất cả các trang
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))  # tăng scale cho rõ
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            lbl = QLabel()
            lbl.setPixmap(QPixmap.fromImage(img))
            lbl.setScaledContents(True)
            self.vbox_preview.addWidget(lbl)

        doc.close()

    def check_update(self):
        try:
            resp = requests.get(VERSION_URL, timeout=5)
            data = resp.json()
            latest_ver = data.get("version", APP_VERSION)
            if latest_ver != APP_VERSION:
                changelog = data.get("changelog", "Không có thông tin.")
                download_url = data.get("download_url", "")
                reply = QMessageBox.question(
                    self, "Có bản cập nhật mới!",
                    f"Phiên bản hiện tại: {APP_VERSION}\n"
                    f"Phiên bản mới: {latest_ver}\n\n"
                    f"Cập nhật:\n{changelog}\n\n"
                    "Bạn có muốn tải về không?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.Yes and download_url:
                    webbrowser.open(download_url)
            else:
                QMessageBox.information(self, "Thông báo", "Bạn đang dùng bản mới nhất.")
        except Exception as e:
            QMessageBox.warning(self, "Lỗi", f"Không thể kiểm tra cập nhật:\n{e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFMergerApp()
    window.show()
    sys.exit(app.exec())
