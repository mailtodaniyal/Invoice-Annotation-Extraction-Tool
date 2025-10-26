import sys, os, json, io, math, random
from PyQt5 import QtWidgets, QtGui, QtCore
import fitz
import spacy
from spacy.training import Example
import pandas as pd
from openpyxl import Workbook
from difflib import SequenceMatcher

class Annotator(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Invoice Annotation & Extraction')
        self.resize(1400, 900)
        self.pdf_path = None
        self.doc = None
        self.page_index = 0
        self.annotations = {}
        self.model_dir = 'models'
        self.labels = []
        self.image_scale = 1.0
        self._start = None
        self.current_rect = None
        self.setup_ui()
        
    def setup_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        main_layout = QtWidgets.QHBoxLayout(central)
        left_panel = self.create_left_panel()
        right_panel = self.create_right_panel()
        main_layout.addWidget(left_panel, 3)
        main_layout.addWidget(right_panel, 1)
        
    def create_left_panel(self):
        panel = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(panel)
        nav_layout = QtWidgets.QHBoxLayout()
        self.openBtn = QtWidgets.QPushButton('Open PDF')
        self.prevBtn = QtWidgets.QPushButton('Prev Page')
        self.nextBtn = QtWidgets.QPushButton('Next Page')
        self.pageLabel = QtWidgets.QLabel('Page: 0/0')
        nav_layout.addWidget(self.openBtn)
        nav_layout.addWidget(self.prevBtn)
        nav_layout.addWidget(self.nextBtn)
        nav_layout.addWidget(self.pageLabel)
        nav_layout.addStretch()
        self.canvas = QtWidgets.QLabel()
        self.canvas.setAlignment(QtCore.Qt.AlignTop | QtCore.Qt.AlignLeft)
        self.canvas.setStyleSheet("background-color: white; border: 1px solid #ccc;")
        self.canvas.setMouseTracking(True)
        self.canvas.installEventFilter(self)
        scroll = QtWidgets.QScrollArea()
        scroll.setWidget(self.canvas)
        scroll.setWidgetResizable(True)
        layout.addLayout(nav_layout)
        layout.addWidget(scroll)
        self.openBtn.clicked.connect(self.open_pdf)
        self.prevBtn.clicked.connect(self.prev_page)
        self.nextBtn.clicked.connect(self.next_page)
        return panel
        
    def create_right_panel(self):
        panel = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(panel)
        labels_group = QtWidgets.QGroupBox("Field Labels")
        labels_layout = QtWidgets.QVBoxLayout(labels_group)
        labels_control_layout = QtWidgets.QHBoxLayout()
        self.labelEdit = QtWidgets.QLineEdit()
        self.labelEdit.setPlaceholderText("Enter new field label")
        self.addLabelBtn = QtWidgets.QPushButton("Add")
        labels_control_layout.addWidget(self.labelEdit)
        labels_control_layout.addWidget(self.addLabelBtn)
        self.labelCombo = QtWidgets.QComboBox()
        self.labelsList = QtWidgets.QListWidget()
        self.removeLabelBtn = QtWidgets.QPushButton("Remove Selected")
        labels_layout.addLayout(labels_control_layout)
        labels_layout.addWidget(QtWidgets.QLabel("Available Labels:"))
        labels_layout.addWidget(self.labelCombo)
        labels_layout.addWidget(self.labelsList)
        labels_layout.addWidget(self.removeLabelBtn)
        annotation_group = QtWidgets.QGroupBox("Annotations")
        annotation_layout = QtWidgets.QVBoxLayout(annotation_group)
        self.annotationList = QtWidgets.QListWidget()
        self.removeAnnoBtn = QtWidgets.QPushButton("Remove Selected Annotation")
        annotation_layout.addWidget(self.annotationList)
        annotation_layout.addWidget(self.removeAnnoBtn)
        actions_group = QtWidgets.QGroupBox("Actions")
        actions_layout = QtWidgets.QVBoxLayout(actions_group)
        self.saveAnnoBtn = QtWidgets.QPushButton('Save Annotations')
        self.exportTrainBtn = QtWidgets.QPushButton('Export Training Data')
        self.trainBtn = QtWidgets.QPushButton('Train Model')
        self.runExtractBtn = QtWidgets.QPushButton('Run Extraction')
        actions_layout.addWidget(self.saveAnnoBtn)
        actions_layout.addWidget(self.exportTrainBtn)
        actions_layout.addWidget(self.trainBtn)
        actions_layout.addWidget(self.runExtractBtn)
        layout.addWidget(labels_group)
        layout.addWidget(annotation_group)
        layout.addWidget(actions_group)
        layout.addStretch()
        self.addLabelBtn.clicked.connect(self.add_label)
        self.removeLabelBtn.clicked.connect(self.remove_label)
        self.saveAnnoBtn.clicked.connect(self.save_current_annotation)
        self.exportTrainBtn.clicked.connect(self.export_training)
        self.trainBtn.clicked.connect(self.train_model)
        self.runExtractBtn.clicked.connect(self.run_extraction)
        self.removeAnnoBtn.clicked.connect(self.remove_annotation)
        return panel
        
    def add_label(self):
        label = self.labelEdit.text().strip()
        if label and label not in self.labels:
            self.labels.append(label)
            self.labelCombo.addItem(label)
            self.labelsList.addItem(label)
            self.labelEdit.clear()
            
    def remove_label(self):
        current = self.labelsList.currentRow()
        if current >= 0:
            label = self.labels.pop(current)
            self.labelCombo.removeItem(self.labelCombo.findText(label))
            self.labelsList.takeItem(current)
            
    def open_pdf(self):
        path = QtWidgets.QFileDialog.getOpenFileName(self, 'Open PDF', '', 'PDF Files (*.pdf)')[0]
        if not path: return
        self.pdf_path = path
        self.doc = fitz.open(path)
        self.page_index = 0
        self.annotations.setdefault(os.path.basename(path), {})
        self.annotations[os.path.basename(path)].setdefault(str(self.page_index), [])
        self.show_page()
        
    def render_page_image(self, page):
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = QtGui.QImage(pix.samples, pix.width, pix.height, pix.stride, QtGui.QImage.Format_RGB888)
        return img
        
    def show_page(self):
        if not self.doc: return
        page = self.doc[self.page_index]
        img = self.render_page_image(page)
        self.page_image = img
        pix = QtGui.QPixmap.fromImage(img)
        self.canvas.setPixmap(pix)
        self.canvas.adjustSize()
        self.rects = self.annotations.get(os.path.basename(self.pdf_path), {}).get(str(self.page_index), [])
        self.pageLabel.setText(f'Page: {self.page_index + 1}/{len(self.doc)}')
        self.update_annotation_list()
        self.update_overlay()
        
    def prev_page(self):
        if not self.doc: return
        if self.page_index > 0:
            self.page_index -= 1
            self.show_page()
            
    def next_page(self):
        if not self.doc: return
        if self.page_index < len(self.doc) - 1:
            self.page_index += 1
            self.show_page()
            
    def eventFilter(self, source, event):
        if source is self.canvas:
            if event.type() == QtCore.QEvent.MouseButtonPress:
                self._start = event.pos()
                return True
            if event.type() == QtCore.QEvent.MouseMove and self._start:
                self.current_rect = (self._start, event.pos())
                self.update_overlay()
                return True
            if event.type() == QtCore.QEvent.MouseButtonRelease and self._start:
                start, end = self._start, event.pos()
                self._start = None
                x1, y1, x2, y2 = min(start.x(), end.x()), min(start.y(), end.y()), max(start.x(), end.x()), max(start.y(), end.y())
                if abs(x2 - x1) > 5 and abs(y2 - y1) > 5:
                    label = self.labelCombo.currentText()
                    if label:
                        bbox = (x1, y1, x2, y2)
                        text = self.extract_text_from_bbox(bbox)
                        entry = {'label': label, 'bbox': bbox, 'text': text}
                        self.annotations.setdefault(os.path.basename(self.pdf_path), {}).setdefault(str(self.page_index), []).append(entry)
                        self.rects.append(entry)
                        self.update_annotation_list()
                        self.update_overlay()
                return True
        return super().eventFilter(source, event)
        
    def update_overlay(self):
        if not hasattr(self, 'page_image'): return
        img = self.page_image.copy()
        painter = QtGui.QPainter(img)
        for r in self.rects:
            bx1, by1, bx2, by2 = r['bbox']
            pen = QtGui.QPen(QtCore.Qt.red, 3)
            painter.setPen(pen)
            painter.drawRect(bx1, by1, bx2 - bx1, by2 - by1)
            painter.drawText(bx1, by1 - 5, r['label'])
        if self.current_rect:
            s, e = self.current_rect
            pen = QtGui.QPen(QtCore.Qt.blue, 2, QtCore.Qt.DashLine)
            painter.setPen(pen)
            painter.drawRect(s.x(), s.y(), e.x() - s.x(), e.y() - s.y())
        painter.end()
        pix = QtGui.QPixmap.fromImage(img)
        self.canvas.setPixmap(pix)
        
    def extract_text_from_bbox(self, bbox):
        if not self.doc: return ''
        page = self.doc[self.page_index]
        x1, y1, x2, y2 = bbox
        w, h = self.canvas.pixmap().width(), self.canvas.pixmap().height()
        page_rect = page.rect
        px_scale, py_scale = page_rect.width / w, page_rect.height / h
        rx1, ry1, rx2, ry2 = x1 * px_scale, y1 * py_scale, x2 * px_scale, y2 * py_scale

        # Try to get text layer only; no OCR fallback
        words = page.get_text("words")
        if not words:
            return "[no text layer detected]"
        words_in = [(wy1, wx1, text) for wx1, wy1, wx2, wy2, text, *_ in words if (wx1 >= rx1 and wy1 >= ry1 and wx2 <= rx2 and wy2 <= ry2)]
        words_in.sort()
        text = ' '.join([t[2] for t in words_in])
        return text.strip() if text else "[empty]"

    def update_annotation_list(self):
        self.annotationList.clear()
        for i, r in enumerate(self.rects):
            self.annotationList.addItem(f"{i}: {r['label']} -> {r['text']}")

    def remove_annotation(self):
        current = self.annotationList.currentRow()
        if current >= 0 and current < len(self.rects):
            self.rects.pop(current)
            self.update_annotation_list()
            self.update_overlay()
            
    def save_current_annotation(self):
        if not self.pdf_path:
            QtWidgets.QMessageBox.warning(self, 'Warning', 'No PDF loaded')
            return
        save_path = QtWidgets.QFileDialog.getSaveFileName(self, 'Save Annotations', 'annotations.json', 'JSON Files (*.json)')[0]
        if not save_path: return
        data = {'labels': self.labels, 'annotations': self.annotations}
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        QtWidgets.QMessageBox.information(self, 'Saved', 'Annotations saved successfully')

    def export_training(self):
        if not self.annotations:
            QtWidgets.QMessageBox.warning(self, 'Warning', 'No annotations to export')
            return
        save_path = QtWidgets.QFileDialog.getSaveFileName(self, 'Export Training', 'training.jsonl', 'JSONL Files (*.jsonl)')[0]
        if not save_path: return
        examples = []
        for fname, pages in self.annotations.items():
            for pidx, anns in pages.items():
                pdf_path = self.find_pdf_path(fname)
                if not pdf_path:
                    continue
                try:
                    doc = fitz.open(pdf_path)
                    page = doc[int(pidx)]
                    page_text = page.get_text('text').replace('\n', ' ').strip()
                    doc.close()
                    entities = []
                    def overlaps(e1, e2):
                        return not (e1[1] <= e2[0] or e2[1] <= e1[0])
                    for a in anns:
                        t = a['text'].strip().replace('\n', ' ').strip()
                        if not t or "[no text" in t or "[empty]" in t:
                            continue
                        start = page_text.find(t)
                        if start == -1:
                            # Try exact match with normalized text
                            t_norm = t.replace(' ', '').lower()
                            page_norm = page_text.replace(' ', '').lower()
                            start = page_norm.find(t_norm)
                            if start != -1:
                                # Map back to original positions
                                orig_pos = 0
                                norm_pos = 0
                                while norm_pos < start:
                                    if page_text[orig_pos] != ' ':
                                        norm_pos += 1
                                    orig_pos += 1
                                end = orig_pos + len(t)
                            else:
                                # Fuzzy match as fallback
                                best_ratio, best_start = 0, -1
                                for i in range(len(page_text) - len(t) + 1):
                                    snippet = page_text[i:i+len(t)]
                                    ratio = SequenceMatcher(None, snippet.lower(), t.lower()).ratio()
                                    if ratio > best_ratio:
                                        best_ratio, best_start = ratio, i
                                if best_ratio > 0.6:
                                    start = best_start
                        if start == -1:
                            continue
                        end = start + len(t)
                        ent = (start, end, a['label'])
                        if not any(overlaps(ent, existing) for existing in entities):
                            entities.append(ent)
                    if entities:
                        examples.append({'text': page_text, 'entities': entities})
                except:
                    continue
        if not examples:
            QtWidgets.QMessageBox.warning(self, 'Warning', 'No valid training examples found')
            return
        # Filter out examples with no entities
        examples = [ex for ex in examples if ex['entities']]
        with open(save_path, 'w', encoding='utf-8') as f:
            for ex in examples:
                f.write(json.dumps(ex, ensure_ascii=False) + '\n')
        QtWidgets.QMessageBox.information(self, 'Export', f'Exported {len(examples)} training examples')

    def find_pdf_path(self, filename):
        if self.pdf_path and os.path.basename(self.pdf_path) == filename:
            return self.pdf_path
        for root, _, files in os.walk('.'):
            if filename in files:
                return os.path.join(root, filename)
        return None

    def train_model(self):
        train_path = QtWidgets.QFileDialog.getOpenFileName(self, 'Open training jsonl', '', 'JSONL Files (*.jsonl)')[0]
        if not train_path:
            return
        TRAIN_DATA = []
        with open(train_path, 'r', encoding='utf-8') as f:
            for line in f:
                obj = json.loads(line.strip())
                if obj['entities']:
                    TRAIN_DATA.append((obj['text'], {'entities': obj['entities']}))
        if not TRAIN_DATA:
            QtWidgets.QMessageBox.warning(self, 'Warning', 'No training data found')
            return

        try:
            nlp = spacy.blank("en")
            if "ner" not in nlp.pipe_names:
                ner = nlp.add_pipe("ner")
            else:
                ner = nlp.get_pipe("ner")

            for _, ann in TRAIN_DATA:
                for ent in ann["entities"]:
                    ner.add_label(ent[2])

            optimizer = nlp.begin_training()
            for itn in range(30):
                random.shuffle(TRAIN_DATA)
                losses = {}
                for text, annotations in TRAIN_DATA:
                    doc = nlp.make_doc(text)
                    example = Example.from_dict(doc, annotations)
                    nlp.update([example], sgd=optimizer, drop=0.2, losses=losses)

            os.makedirs(self.model_dir, exist_ok=True)
            nlp.to_disk(self.model_dir)
            QtWidgets.QMessageBox.information(self, "Training Complete", "Model trained and saved successfully!")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Training failed: {str(e)}")

    def run_extraction(self):
        QtWidgets.QMessageBox.information(self, "Note", "Extraction step ready — waiting for new PDFs.")

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    win = Annotator()
    win.show()
    sys.exit(app.exec_())
