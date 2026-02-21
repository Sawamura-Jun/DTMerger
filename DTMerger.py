from __future__ import annotations

import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

from PIL import Image
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)
from xdwlib import xdwopen

# Large engineering TIFFs can exceed Pillow's decompression-bomb threshold.
Image.MAX_IMAGE_PIXELS = None

WINDOW_WIDTH = 900
WINDOW_HEIGHT = 640


TIFF_EXTENSIONS = {".tif", ".tiff"}
DOCUWORKS_EXTENSIONS = {".xdw", ".xbd", ".xct"}
SUPPORTED_EXTENSIONS = TIFF_EXTENSIONS | DOCUWORKS_EXTENSIONS


@dataclass
class PageEntry:
    source_path: Path
    page_index: int
    source_type: str  # "tiff" or "docuworks"

    @property
    def label(self) -> str:
        return f"{self.source_path.name}-p{self.page_index + 1:03d}"


def extract_supported_paths_from_mime(mime_data) -> List[Path]:
    if not mime_data.hasUrls():
        return []

    paths: List[Path] = []
    for url in mime_data.urls():
        if not url.isLocalFile():
            continue
        path = Path(url.toLocalFile())
        if not path.is_file():
            continue
        if path.suffix.lower() in SUPPORTED_EXTENSIONS:
            paths.append(path)
    return paths


class PageListWidget(QListWidget):
    files_dropped = Signal(list)

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.NoDragDrop)
        self.setAlternatingRowColors(True)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if extract_supported_paths_from_mime(event.mimeData()):
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event) -> None:
        if extract_supported_paths_from_mime(event.mimeData()):
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = extract_supported_paths_from_mime(event.mimeData())
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
            return
        event.ignore()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("DTMerger")
        self.resize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.setAcceptDrops(True)

        self.page_list = PageListWidget()
        self.page_list.files_dropped.connect(self.add_files)

        self.up_button = QPushButton("↑")
        self.down_button = QPushButton("↓")
        self.g4_button = QPushButton("G4形式で出力")
        self.lzw_button = QPushButton("LZW形式で出力")

        self.up_button.clicked.connect(self.move_selected_up)
        self.down_button.clicked.connect(self.move_selected_down)
        self.g4_button.clicked.connect(lambda: self.export_tiff("group4"))
        self.lzw_button.clicked.connect(lambda: self.export_tiff("tiff_lzw"))

        right_layout = QVBoxLayout()
        right_layout.addWidget(self.up_button)
        right_layout.addWidget(self.down_button)
        right_layout.addSpacing(24)
        right_layout.addWidget(self.g4_button)
        right_layout.addWidget(self.lzw_button)
        right_layout.addStretch(1)

        main_layout = QHBoxLayout()
        main_layout.addWidget(self.page_list, stretch=1)
        main_layout.addLayout(right_layout)

        central = QWidget()
        central.setLayout(main_layout)
        self.setCentralWidget(central)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if extract_supported_paths_from_mime(event.mimeData()):
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event) -> None:
        if extract_supported_paths_from_mime(event.mimeData()):
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = extract_supported_paths_from_mime(event.mimeData())
        if paths:
            self.add_files(paths)
            event.acceptProposedAction()
            return
        event.ignore()

    def add_files(self, paths: List[Path]) -> None:
        errors: List[str] = []
        for path in paths:
            ext = path.suffix.lower()
            try:
                if ext in TIFF_EXTENSIONS:
                    page_count = self.get_tiff_page_count(path)
                    source_type = "tiff"
                elif ext in DOCUWORKS_EXTENSIONS:
                    page_count = self.get_docuworks_page_count(path)
                    source_type = "docuworks"
                else:
                    continue
            except Exception as exc:
                errors.append(f"{path.name}: {exc}")
                continue

            for page_index in range(page_count):
                entry = PageEntry(
                    source_path=path,
                    page_index=page_index,
                    source_type=source_type,
                )
                item = QListWidgetItem(entry.label)
                item.setData(Qt.UserRole, entry)
                self.page_list.addItem(item)

        if errors:
            QMessageBox.warning(
                self,
                "読み込みエラー",
                "一部のファイルを読み込めませんでした。\n\n" + "\n".join(errors),
            )

    @staticmethod
    def get_tiff_page_count(path: Path) -> int:
        with Image.open(path) as img:
            return getattr(img, "n_frames", 1)

    @staticmethod
    def get_docuworks_page_count(path: Path) -> int:
        with xdwopen(str(path), readonly=True) as doc:
            return doc.pages

    def move_selected_up(self) -> None:
        row = self.page_list.currentRow()
        if row <= 0:
            return
        item = self.page_list.takeItem(row)
        self.page_list.insertItem(row - 1, item)
        self.page_list.setCurrentRow(row - 1)

    def move_selected_down(self) -> None:
        row = self.page_list.currentRow()
        if row < 0 or row >= self.page_list.count() - 1:
            return
        item = self.page_list.takeItem(row)
        self.page_list.insertItem(row + 1, item)
        self.page_list.setCurrentRow(row + 1)

    def export_tiff(self, compression: str) -> None:
        if self.page_list.count() == 0:
            QMessageBox.information(self, "情報", "出力対象のページがありません。")
            return

        output_path_text, _ = QFileDialog.getSaveFileName(
            self,
            "保存先を選択",
            str(Path.home() / "merged_output.tif"),
            "TIFF (*.tif *.tiff)",
        )
        if not output_path_text:
            return

        output_path = Path(output_path_text)
        if output_path.suffix.lower() not in TIFF_EXTENSIONS:
            output_path = output_path.with_suffix(".tif")

        entries = self.collect_entries()
        try:
            self.create_merged_tiff(entries, output_path, compression)
        except Exception as exc:
            QMessageBox.critical(self, "出力エラー", f"TIFF出力に失敗しました。\n\n{exc}")
            return

        QMessageBox.information(
            self,
            "完了",
            f"出力が完了しました。\n{output_path}",
        )

    def collect_entries(self) -> List[PageEntry]:
        entries: List[PageEntry] = []
        for row in range(self.page_list.count()):
            item = self.page_list.item(row)
            entry = item.data(Qt.UserRole)
            entries.append(entry)
        return entries

    def create_merged_tiff(
        self,
        entries: List[PageEntry],
        output_path: Path,
        compression: str,
    ) -> None:
        frames: List[Image.Image] = []
        xdw_docs: Dict[Path, object] = {}

        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                for order, entry in enumerate(entries):
                    if entry.source_type == "tiff":
                        frame = self.load_tiff_page(entry.source_path, entry.page_index)
                    else:
                        frame = self.convert_docuworks_page(
                            entry=entry,
                            order=order,
                            temp_dir=Path(temp_dir),
                            opened_docs=xdw_docs,
                        )

                    if compression == "group4":
                        frame = self.ensure_group4_mode(frame)
                    frames.append(frame)
            finally:
                for doc in xdw_docs.values():
                    doc.close()

        if not frames:
            raise RuntimeError("出力対象のページがありません。")

        first, *rest = frames
        first.save(
            output_path,
            format="TIFF",
            save_all=True,
            append_images=rest,
            compression=compression,
            dpi=(400, 400),
        )

        for frame in frames:
            frame.close()

    @staticmethod
    def load_tiff_page(path: Path, page_index: int) -> Image.Image:
        with Image.open(path) as img:
            max_index = getattr(img, "n_frames", 1) - 1
            if page_index > max_index:
                raise IndexError(
                    f"ページ {page_index + 1} は存在しません。"
                    f" ({path.name}: 1..{max_index + 1})"
                )
            img.seek(page_index)
            return img.copy()

    @staticmethod
    def convert_docuworks_page(
        entry: PageEntry,
        order: int,
        temp_dir: Path,
        opened_docs: Dict[Path, object],
    ) -> Image.Image:
        doc = opened_docs.get(entry.source_path)
        if doc is None:
            doc = xdwopen(str(entry.source_path), readonly=True)
            opened_docs[entry.source_path] = doc

        temp_tiff = temp_dir / f"docuworks_{order:05d}.tif"
        doc.page(entry.page_index).export_image(
            path=str(temp_tiff),
            dpi=400,
            format="TIFF",
            compress="NOCOMPRESS",
        )

        with Image.open(temp_tiff) as img:
            return img.copy()

    @staticmethod
    def ensure_group4_mode(img: Image.Image) -> Image.Image:
        if img.mode == "1":
            return img
        return img.convert("1")


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
