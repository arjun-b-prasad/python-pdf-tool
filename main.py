import os
import shutil
import sys
import tempfile
from pathlib import Path
from typing import List

from PyQt6.QtCore import Qt, QEvent
from PyQt6.QtGui import QPalette, QColor, QIcon, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QInputDialog,
)


SUPPORTED_EXTENSIONS = {".pdf", ".tif", ".tiff", ".jpg", ".jpeg"}
APP_TITLE = "Stronghold File Editor"


def resource_path(relative: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / relative


APP_ICON_PATH = resource_path("img/logo.ico")
UI_LOGO_PATH = resource_path("img/logo-white.png")


class MergeWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(APP_ICON_PATH)))
        self.resize(860, 540)
        self._build_ui()
        self.setAcceptDrops(True)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(18)

        if UI_LOGO_PATH.exists():
            logo_pixmap = QPixmap(str(UI_LOGO_PATH))
            if logo_pixmap.width() > 320:
                logo_pixmap = logo_pixmap.scaledToWidth(320, Qt.TransformationMode.SmoothTransformation)
            self.logo_label = QLabel()
            self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.logo_label.setPixmap(logo_pixmap)
            root.addWidget(self.logo_label)
        else:
            self.logo_label = None

        info = QLabel(
            "Add PDF, TIFF, or JPG files, reorder them, double-click or select Rename to adjust file names, "
            "and merge everything into a single PDF."
        )
        info.setWordWrap(True)
        info.setObjectName("infoLabel")
        root.addWidget(info)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.file_list.setAlternatingRowColors(True)
        self.file_list.setDragEnabled(True)
        self.file_list.setAcceptDrops(True)
        self.file_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.file_list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.file_list.itemSelectionChanged.connect(self._sync_button_states)
        self.file_list.model().rowsMoved.connect(self._sync_button_states)
        self.file_list.itemDoubleClicked.connect(self._rename_item_inline)
        self.file_list.installEventFilter(self)
        root.addWidget(self.file_list, stretch=1)

        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(12)

        self.add_button = QPushButton("Add Files")
        self.add_button.clicked.connect(self._add_files)
        controls_layout.addWidget(self.add_button)

        self.remove_button = QPushButton("Remove Selected")
        self.remove_button.clicked.connect(self._remove_selected)
        controls_layout.addWidget(self.remove_button)

        self.rename_button = QPushButton("Rename Selected")
        self.rename_button.clicked.connect(self._rename_selected)
        controls_layout.addWidget(self.rename_button)

        self.move_up_button = QPushButton("Move Up")
        self.move_up_button.clicked.connect(lambda: self._move_selection(-1))
        controls_layout.addWidget(self.move_up_button)

        self.move_down_button = QPushButton("Move Down")
        self.move_down_button.clicked.connect(lambda: self._move_selection(1))
        controls_layout.addWidget(self.move_down_button)

        self.merge_button = QPushButton("Merge Files")
        self.merge_button.setObjectName("primaryButton")
        self.merge_button.clicked.connect(self._merge_files)
        controls_layout.addWidget(self.merge_button)

        self.export_button = QPushButton("Export to JPG")
        self.export_button.setObjectName("exportButton")
        self.export_button.clicked.connect(self._export_to_jpg)
        controls_layout.addWidget(self.export_button)

        controls_layout.addStretch(1)
        root.addLayout(controls_layout)

        self.feedback = QLabel("")
        self.feedback.setObjectName("feedbackLabel")
        root.addWidget(self.feedback)

        self._sync_button_states()

    # --- UI callbacks -------------------------------------------------
    def _add_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PDF, TIFF, or JPG files",
            str(Path.home()),
            "PDF, TIFF, and JPG (*.pdf *.tif *.tiff *.jpg *.jpeg);;PDF (*.pdf);;TIFF (*.tif *.tiff);;JPG (*.jpg *.jpeg);;All Files (*.*)",
        )
        self._add_path_batch(paths)
        self._sync_button_states()

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls() and self._contains_supported_files(event.mimeData().urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event) -> None:
        if event.mimeData().hasUrls() and self._contains_supported_files(event.mimeData().urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:
        if not event.mimeData().hasUrls():
            event.ignore()
            return

        paths = []
        for url in event.mimeData().urls():
            local_path = url.toLocalFile()
            if local_path:
                paths.append(local_path)

        self._add_path_batch(paths)
        self._sync_button_states()
        event.acceptProposedAction()

    def _remove_selected(self) -> None:
        selected = self.file_list.selectedItems()
        if not selected:
            return
        for item in selected:
            self.file_list.takeItem(self.file_list.row(item))
        self.feedback.setText("Removed selected entries.")
        self._sync_button_states()

    def eventFilter(self, source, event):
        if source is self.file_list and event.type() == QEvent.Type.KeyPress:
            if event.key() == Qt.Key.Key_Delete:
                self._remove_selected()
                return True
        return super().eventFilter(source, event)

    def _rename_selected(self) -> None:
        selected = self.file_list.selectedItems()
        if not selected:
            self._show_info("Select at least one file to rename.")
            return
        for item in selected:
            self._rename_item(item)

    def _rename_item_inline(self, item: QListWidgetItem) -> None:
        self._rename_item(item)

    def _move_selection(self, offset: int) -> None:
        if offset == 0:
            return
        selected_rows = sorted({self.file_list.row(item) for item in self.file_list.selectedItems()})
        if not selected_rows:
            return

        if offset < 0:
            iterator = selected_rows
        else:
            iterator = reversed(selected_rows)

        for row in iterator:
            target = row + offset
            if target < 0 or target >= self.file_list.count():
                continue
            item = self.file_list.takeItem(row)
            self.file_list.insertItem(target, item)
            item.setSelected(True)
        self.feedback.setText("Reordered selection.")
        self._sync_button_states()

    def _merge_files(self) -> None:
        paths = self._gather_paths()
        if not paths:
            self._show_info("Add at least one file before merging.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save merged PDF as...",
            str(Path.home() / "merged.pdf"),
            "PDF (*.pdf)",
        )
        if not output_path:
            return

        output_path_obj = Path(output_path)
        if output_path_obj.suffix.lower() != ".pdf":
            if output_path_obj.suffix:
                output_path_obj = output_path_obj.with_suffix(".pdf")
            else:
                output_path_obj = output_path_obj.parent / f"{output_path_obj.name}.pdf"

        try:
            self._merge_into_pdf(paths, str(output_path_obj))
        except ImportError as exc:
            self._show_error(f"Missing dependency: {exc}. Install the required package and try again.")
            return
        except Exception as exc:  # pylint: disable=broad-except
            self._show_error(f"Merge failed: {exc}")
            return

        self._show_info(f"Merged {len(paths)} file(s) into {output_path_obj}.")

    def _export_to_jpg(self) -> None:
        paths = self._gather_paths()
        if not paths:
            self._show_info("Add at least one file before exporting.")
            return

        destination = QFileDialog.getExistingDirectory(
            self,
            "Select folder to store JPG files",
            str(Path.home()),
        )
        if not destination:
            return

        exported_files = 0
        issues: List[str] = []
        output_dir = Path(destination)

        for path in paths:
            try:
                exported_files += self._convert_file_to_jpg(path, output_dir)
            except ImportError as exc:
                self._show_error(f"Missing dependency: {exc}. Install the required package and try again.")
                return
            except Exception as exc:  # pylint: disable=broad-except
                issues.append(f"{Path(path).name}: {exc}")

        if exported_files:
            self._show_info(f"Exported {exported_files} JPG file(s) to {output_dir}.")
        else:
            self._show_error("No JPG files were exported.")

        if issues:
            QMessageBox.warning(
                self,
                "Export completed with issues",
                "\n".join(issues),
            )

    # --- Helper logic -------------------------------------------------
    def _path_already_listed(self, path: str) -> bool:
        for index in range(self.file_list.count()):
            if self.file_list.item(index).data(Qt.ItemDataRole.UserRole) == path:
                return True
        return False

    def _contains_supported_files(self, urls) -> bool:
        for url in urls:
            local_path = Path(url.toLocalFile())
            if not local_path.is_file():
                continue
            if local_path.suffix.lower() in SUPPORTED_EXTENSIONS:
                return True
        return False

    def _add_path_batch(self, paths: List[str]) -> None:
        if not paths:
            return

        added = 0
        for path in paths:
            candidate = Path(path)
            if not candidate.is_file():
                continue
            suffix = candidate.suffix.lower()
            if suffix not in SUPPORTED_EXTENSIONS:
                continue
            if self._path_already_listed(path):
                continue
            item = QListWidgetItem(candidate.name)
            item.setData(Qt.ItemDataRole.UserRole, path)
            self.file_list.addItem(item)
            added += 1

        if added:
            self.feedback.setText(f"Added {added} file(s).")
        else:
            self.feedback.setText("No new supported files were added.")

    def _gather_paths(self) -> List[str]:
        return [
            self.file_list.item(index).data(Qt.ItemDataRole.UserRole)
            for index in range(self.file_list.count())
        ]

    def _rename_item(self, item: QListWidgetItem) -> None:
        original_path = Path(item.data(Qt.ItemDataRole.UserRole))
        new_name, ok = QInputDialog.getText(self, "Rename File", "New file name:", text=original_path.name)
        if not ok or not new_name.strip():
            return

        new_name = new_name.strip()
        new_suffix = Path(new_name).suffix.lower()
        current_suffix = original_path.suffix.lower()
        if new_suffix and new_suffix not in SUPPORTED_EXTENSIONS:
            self._show_error("Unsupported extension. Keep the file as PDF, TIF, TIFF, or JPG.")
            return
        if not new_suffix:
            new_name = f"{new_name}{current_suffix}"

        target_path = original_path.with_name(new_name)
        if target_path.exists():
            overwrite = QMessageBox.question(
                self,
                "Overwrite file?",
                f"{target_path.name} already exists. Overwrite it?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if overwrite != QMessageBox.StandardButton.Yes:
                return

        try:
            os.rename(original_path, target_path)
        except OSError as exc:
            self._show_error(f"Failed to rename file: {exc}")
            return

        item.setData(Qt.ItemDataRole.UserRole, str(target_path))
        item.setText(target_path.name)
        self.feedback.setText(f"Renamed to {target_path.name}.")

    def _merge_into_pdf(self, paths: List[str], output_path: str) -> None:
        try:
            from PyPDF2 import PdfMerger
        except ImportError as exc:
            raise ImportError("PyPDF2 is required to write PDF output") from exc

        temp_files: List[str] = []
        merger = PdfMerger()

        try:
            for path in paths:
                suffix = Path(path).suffix.lower()
                if suffix == ".pdf":
                    merger.append(path)
                elif suffix in {".tif", ".tiff"}:
                    temp_pdf = self._convert_tiff_to_pdf(path)
                    temp_files.append(temp_pdf)
                    merger.append(temp_pdf)
                elif suffix in {".jpg", ".jpeg"}:
                    temp_pdf = self._convert_image_to_pdf(path)
                    temp_files.append(temp_pdf)
                    merger.append(temp_pdf)
                else:
                    raise ValueError(f"Unsupported file encountered: {path}")

            with open(output_path, "wb") as handle:
                merger.write(handle)
        finally:
            merger.close()
            for temp_file in temp_files:
                try:
                    Path(temp_file).unlink()
                except FileNotFoundError:
                    pass

    def _merge_into_tiff(self, paths: List[str], output_path: str) -> None:
        try:
            from PIL import Image, ImageSequence
        except ImportError as exc:
            raise ImportError("Pillow is required to write TIFF output") from exc

        frames = []
        for path in paths:
            suffix = Path(path).suffix.lower()
            if suffix not in {".tif", ".tiff"}:
                raise ValueError("Only TIFF files can be merged into a TIFF output.")
            with Image.open(path) as image:
                for frame in ImageSequence.Iterator(image):
                    frames.append(frame.convert("RGB").copy())

        if not frames:
            raise ValueError("No image frames to write.")

        first, *rest = frames
        first.save(
            output_path,
            format="TIFF",
            save_all=True,
            append_images=rest,
            compression="tiff_deflate",
        )

    def _convert_tiff_to_pdf(self, path: str) -> str:
        try:
            from PIL import Image, ImageSequence
        except ImportError as exc:
            raise ImportError("Pillow is required to convert TIFF files") from exc

        frames = []
        with Image.open(path) as image:
            for frame in ImageSequence.Iterator(image):
                frames.append(frame.convert("RGB").copy())

        if not frames:
            raise ValueError(f"No frames found in TIFF: {path}")

        descriptor = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        descriptor.close()
        frames[0].save(descriptor.name, format="PDF", save_all=True, append_images=frames[1:])
        return descriptor.name

    def _convert_image_to_pdf(self, path: str) -> str:
        try:
            from PIL import Image
        except ImportError as exc:
            raise ImportError("Pillow is required to convert image files") from exc

        with Image.open(path) as image:
            rgb_image = image.convert("RGB")
            descriptor = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            descriptor.close()
            rgb_image.save(descriptor.name, format="PDF")
            rgb_image.close()
        return descriptor.name

    def _convert_file_to_jpg(self, path: str, output_dir: Path) -> int:
        suffix = Path(path).suffix.lower()
        if suffix == ".pdf":
            return self._convert_pdf_to_jpg(path, output_dir)
        if suffix in {".tif", ".tiff"}:
            return self._convert_tiff_to_jpg(path, output_dir)
        if suffix in {".jpg", ".jpeg"}:
            return self._copy_jpg_to_output(path, output_dir)
        raise ValueError(f"Unsupported file encountered: {path}")

    def _convert_pdf_to_jpg(self, path: str, output_dir: Path) -> int:
        try:
            import fitz  # type: ignore
        except ImportError as exc:
            raise ImportError("PyMuPDF is required to export PDF pages to JPG") from exc

        exported = 0
        with fitz.open(path) as document:
            for index, page in enumerate(document, start=1):
                pixmap = page.get_pixmap(dpi=200, alpha=False)
                target = output_dir / f"{Path(path).stem}_page{index:03d}.jpg"
                target = self._resolve_conflict_path(target)
                pixmap.save(str(target))
                exported += 1
        return exported

    def _convert_tiff_to_jpg(self, path: str, output_dir: Path) -> int:
        try:
            from PIL import Image, ImageSequence
        except ImportError as exc:
            raise ImportError("Pillow is required to export TIFF frames to JPG") from exc

        exported = 0
        with Image.open(path) as image:
            for index, frame in enumerate(ImageSequence.Iterator(image), start=1):
                target = output_dir / f"{Path(path).stem}_frame{index:03d}.jpg"
                target = self._resolve_conflict_path(target)
                frame.convert("RGB").save(target, format="JPEG", quality=90)
                exported += 1
        return exported

    def _resolve_conflict_path(self, path: Path) -> Path:
        if not path.exists():
            return path

        counter = 1
        while True:
            candidate = path.with_stem(f"{path.stem}_{counter}")
            if not candidate.exists():
                return candidate
            counter += 1

    def _copy_jpg_to_output(self, path: str, output_dir: Path) -> int:
        source = Path(path)
        target = output_dir / source.name
        target = self._resolve_conflict_path(target)
        shutil.copy2(source, target)
        return 1

    def _show_error(self, message: str) -> None:
        QMessageBox.critical(self, "Error", message)
        self.feedback.setText(message)

    def _show_info(self, message: str) -> None:
        QMessageBox.information(self, "Info", message)
        self.feedback.setText(message)

    def _sync_button_states(self) -> None:
        has_items = self.file_list.count() > 0
        has_selection = len(self.file_list.selectedItems()) > 0

        self.remove_button.setEnabled(has_selection)
        self.rename_button.setEnabled(has_selection)
        self.move_up_button.setEnabled(has_selection)
        self.move_down_button.setEnabled(has_selection)
        self.merge_button.setEnabled(has_items)
        self.export_button.setEnabled(has_items)


def apply_modern_palette(app: QApplication) -> None:
    app.setStyle("Fusion")

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(40, 44, 52))
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Base, QColor(30, 34, 40))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(45, 49, 58))
    palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Button, QColor(55, 61, 72))
    palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Highlight, QColor(45, 136, 255))
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
    app.setPalette(palette)

    app.setStyleSheet(
        """
        QWidget {
            font-family: 'Segoe UI';
            font-size: 11pt;
            color: #F1F1F1;
        }
        QListWidget {
            border: 1px solid #3A3F4B;
            border-radius: 6px;
            padding: 6px;
            background-color: #1E2229;
        }
        QListWidget::item {
            padding: 6px;
        }
        QListWidget::item:selected {
            background-color: #2F6FEB;
            color: #FFFFFF;
        }
        QPushButton {
            background-color: #3A3F4B;
            border: 1px solid #4B5160;
            border-radius: 6px;
            padding: 8px 16px;
        }
        QPushButton:hover {
            background-color: #4B5160;
        }
        QPushButton:pressed {
            background-color: #2F6FEB;
            border: 1px solid #2F6FEB;
        }
        QPushButton:disabled {
            color: #777;
            background-color: #2A2D37;
            border-color: #2A2D37;
        }
        QPushButton#primaryButton {
            background-color: #2F6FEB;
            border: 1px solid #2F6FEB;
            color: #FFFFFF;
        }
        QPushButton#primaryButton:hover {
            background-color: #407CFF;
        }
        QPushButton#exportButton {
            background-color: #1DB954;
            border: 1px solid #12833B;
            color: #ffffff;
        }
        QPushButton#exportButton:hover {
            background-color: #1ED760;
        }
        QPushButton#exportButton:pressed {
            background-color: #159C46;
            border: 1px solid #159C46;
        }
        QLabel#infoLabel {
            color: #C8CCD7;
        }
        QLabel#feedbackLabel {
            color: #9DA5B4;
        }
        QMessageBox {
            background-color: #2A2D37;
        }
        """
    )


def main() -> None:
    if sys.platform.startswith("win"):
        try:
            import ctypes  # type: ignore

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Stronghold.File.Editor")
        except (ImportError, AttributeError, OSError):
            pass

    app = QApplication(sys.argv)
    apply_modern_palette(app)
    if APP_ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(APP_ICON_PATH)))
    window = MergeWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
