import sys
import os
import unicodedata
import re
import json
from PyQt5 import QtWidgets, QtGui, QtCore

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except Exception:
    PANDAS_AVAILABLE = False

from rapidfuzz import process, fuzz


class HomeoWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Homeopathy Name Search (PyQt)')
        self.resize(800, 550)
        self.df = None
        self.latin_col = 'latin_col'
        self.common_col = 'common_col'

        # settings persistence
        self._settings_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.json')
        self._settings = {}
        try:
            if os.path.exists(self._settings_path):
                with open(self._settings_path, 'r', encoding='utf-8') as sf:
                    self._settings = json.load(sf)
        except Exception:
            self._settings = {}

        layout = QtWidgets.QVBoxLayout(self)

        # --- Top bar ---
        top = QtWidgets.QHBoxLayout()
        self.query = QtWidgets.QLineEdit()
        self.query.setPlaceholderText('Enter Common Name (English or Bengali)')
        self.search_btn = QtWidgets.QPushButton('Search')
        self.materia_btn = QtWidgets.QPushButton('Materia')
        self.ai_btn = QtWidgets.QPushButton('AI Suggest')
        self.load_btn = QtWidgets.QPushButton('Load Excel')
        self.add_btn = QtWidgets.QPushButton('Add New')  # NEW BUTTON

        top.addWidget(self.query)
        top.addWidget(self.search_btn)
        top.addWidget(self.ai_btn)
        top.addWidget(self.materia_btn)
        top.addWidget(self.load_btn)
        top.addWidget(self.add_btn)
        layout.addLayout(top)

        # --- Control row ---
        control_row = QtWidgets.QHBoxLayout()
        control_row.addStretch()
        self.mode_combo = QtWidgets.QComboBox()
        self.mode_combo.addItems(['Contains (Excel)', 'Starts-with', 'Word-prefix', 'Fuzzy'])
        control_row.addWidget(QtWidgets.QLabel('Mode:'))
        control_row.addWidget(self.mode_combo)
        self.incremental_chk = QtWidgets.QCheckBox('Incremental')
        self.incremental_chk.setChecked(self._settings.get('incremental', True))

        mode_saved = self._settings.get('mode')
        if mode_saved:
            idx = self.mode_combo.findText(mode_saved)
            if idx >= 0:
                self.mode_combo.setCurrentIndex(idx)

        control_row.addWidget(self.incremental_chk)
        control_row.addStretch()
        layout.addLayout(control_row)

        # --- Status + Table ---
        self.status = QtWidgets.QLabel('Ready. Load remedies.xlsx')
        layout.addWidget(self.status)

        self.table = QtWidgets.QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(['Common', 'Latin'])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        # --- Connect buttons ---
        self.search_btn.clicked.connect(self.on_search)
        self.materia_btn.clicked.connect(self.on_materia_search)
        self.ai_btn.clicked.connect(self.on_ai_suggest)
        self.load_btn.clicked.connect(self.load_excel)
        self.add_btn.clicked.connect(self.on_add_new)
        self.query.returnPressed.connect(self.on_search)

        # --- Debounce for incremental search ---
        self._debounce_timer = QtCore.QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.setInterval(250)
        self._debounce_timer.timeout.connect(self.on_search)

        def _on_text_changed(_):
            if self.incremental_chk.isChecked():
                self._debounce_timer.start()
        self.query.textChanged.connect(_on_text_changed)

        # --- Save settings when changed ---
        def _save_settings(*_args):
            s = {'mode': self.mode_combo.currentText(), 'incremental': bool(self.incremental_chk.isChecked())}
            try:
                with open(self._settings_path, 'w', encoding='utf-8') as sf:
                    json.dump(s, sf)
                self._settings = s
            except Exception:
                pass

        self.mode_combo.currentIndexChanged.connect(_save_settings)
        self.incremental_chk.stateChanged.connect(_save_settings)

        # --- Auto load Excel ---
        if PANDAS_AVAILABLE and os.path.exists('remedies.xlsx'):
            try:
                self._load_df('remedies.xlsx')
                self.status.setText(f'Loaded remedies.xlsx with {len(self.df)} entries.')
            except Exception as e:
                self.status.setText(f'Failed to load: {e}')
        elif not PANDAS_AVAILABLE:
            self.status.setText('pandas not installed; install pandas and openpyxl to load Excel.')

    # ============================================================
    # Core Excel & Search Logic
    # ============================================================
    def _load_df(self, path):
        if not PANDAS_AVAILABLE:
            raise RuntimeError('pandas not available')
        df = pd.read_excel(path, engine='openpyxl')
        df.rename(columns={c: c.lower().strip() for c in df.columns}, inplace=True)
        self.df = df[['latin_col', 'common_col']].dropna(how='any')

        def _clean(s):
            if s is None:
                return ''
            s = str(s)
            s = unicodedata.normalize('NFC', s)
            s = s.replace('\u200c', '').replace('\u200d', '')
            s = s.replace('\u00a0', ' ')
            s = re.sub(r'\s+', ' ', s).strip()
            return s

        self.df['common_norm'] = self.df['common_col'].astype(str).apply(_clean)
        self.df['latin_norm'] = self.df['latin_col'].astype(str).apply(lambda x: _clean(x).casefold())
        self.df['common_norm_cf'] = self.df['common_norm'].astype(str).apply(lambda x: unicodedata.normalize('NFC', x).casefold())

    def load_excel(self):
        if not PANDAS_AVAILABLE:
            self.status.setText('pandas not installed.')
            return
        path = 'remedies.xlsx'
        if not os.path.exists(path):
            self.status.setText('No remedies.xlsx found in the current folder.')
            return
        try:
            self._load_df(path)
            self.status.setText(f'Loaded {path} with {len(self.df)} entries.')
        except Exception as e:
            self.status.setText(f'Load error: {e}')

    # ============================================================
    # Add New Remedy
    # ============================================================
    def on_add_new(self):
        """Add new Common & Latin entry to remedies.xlsx"""
        if not PANDAS_AVAILABLE:
            self.status.setText('pandas not installed.')
            return

        path = 'remedies.xlsx'
        if not os.path.exists(path):
            QtWidgets.QMessageBox.warning(self, 'Missing File', 'remedies.xlsx not found. Please create it first.')
            return

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle('Add New Remedy')
        dlg.resize(400, 200)
        layout = QtWidgets.QVBoxLayout(dlg)

        form = QtWidgets.QFormLayout()
        common_input = QtWidgets.QLineEdit()
        latin_input = QtWidgets.QLineEdit()
        form.addRow('Common Name:', common_input)
        form.addRow('Latin Name:', latin_input)
        layout.addLayout(form)

        btns = QtWidgets.QHBoxLayout()
        ok_btn = QtWidgets.QPushButton('Add')
        cancel_btn = QtWidgets.QPushButton('Cancel')
        btns.addWidget(ok_btn)
        btns.addWidget(cancel_btn)
        layout.addLayout(btns)

        ok_btn.clicked.connect(dlg.accept)
        cancel_btn.clicked.connect(dlg.reject)

        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            common_name = common_input.text().strip()
            latin_name = latin_input.text().strip()

            if not common_name or not latin_name:
                QtWidgets.QMessageBox.warning(self, 'Input Error', 'Both Common and Latin names are required.')
                return

            try:
                df = pd.read_excel(path, engine='openpyxl')
                new_row = {'common_col': common_name, 'latin_col': latin_name}
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                df.to_excel(path, index=False, engine='openpyxl')
                self._load_df(path)
                self.status.setText(f'Added: {common_name} - {latin_name}')
                QtWidgets.QMessageBox.information(self, 'Success', 'New remedy added successfully!')
                self.on_search()
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, 'Error', f'Failed to add new entry:\n{e}')

    # ============================================================
    # Materia Medica
    # ============================================================
    def on_materia_search(self):
        import requests
        from bs4 import BeautifulSoup
        latin_name = None
        sel = self.table.selectedItems()
        if sel and len(sel) >= 2:
            latin_name = sel[1].text().strip()
        if not latin_name:
            latin_name = self.query.text().strip()
        if not latin_name:
            self.status.setText('Select a row or type a remedy name for Materia search.')
            return

        slug = latin_name.lower()
        slug = slug.split('(')[0].split(',')[0].strip()
        slug = slug.replace(' ', '-')
        slug = re.sub(r'[^a-z0-9\-]', '', slug)

        base = 'https://www.materiamedica.info/en/materia-medica/william-boericke/'
        url = base + slug
        try:
            r = requests.get(url, timeout=10)
            if r.status_code != 200:
                url, ok = QtWidgets.QInputDialog.getText(self, 'Materia URL', 'Could not find auto URL. Enter full URL:')
                if not ok or not url:
                    self.status.setText('Materia search cancelled.')
                    return
            else:
                soup = BeautifulSoup(r.text, 'html.parser')
                article = soup.find('article') or soup.find('main') or soup
                texts = [p.get_text(separator=' ', strip=True) for p in article.find_all(['h1', 'h2', 'h3', 'p', 'li']) if p.get_text()]
                content = '\n\n'.join(texts[:500])
                dlg = QtWidgets.QDialog(self)
                dlg.setWindowTitle(f'Materia: {latin_name}')
                dlg.resize(800, 600)
                lay = QtWidgets.QVBoxLayout(dlg)
                textw = QtWidgets.QTextEdit()
                textw.setReadOnly(True)
                textw.setPlainText(content)
                lay.addWidget(textw)
                btns = QtWidgets.QHBoxLayout()
                openb = QtWidgets.QPushButton('Open in browser')
                closeb = QtWidgets.QPushButton('Close')
                btns.addWidget(openb)
                btns.addWidget(closeb)
                lay.addLayout(btns)
                openb.clicked.connect(lambda: os.system(f'start {url}'))
                closeb.clicked.connect(dlg.accept)
                dlg.exec_()
        except Exception as e:
            self.status.setText(f'Materia fetch failed: {e}')

    # ============================================================
    # AI Suggestions (OpenAI or Local)
    # ============================================================
    def on_ai_suggest(self):
        q = self.query.text().strip()
        if not q:
            sel = self.table.selectedItems()
            if sel and len(sel) >= 1:
                q = sel[0].text().strip()
        if not q:
            self.status.setText('Type or select a remedy to get AI suggestions.')
            return

        openai_key = os.environ.get('OPENAI_API_KEY')
        used_openai = False
        if openai_key:
            try:
                import openai
                openai.api_key = openai_key
                prompt = (
                    "You are a homeopathy assistant. Given a remedy name or common name, "
                    "suggest possible alternative names, short indications, and closely related remedies. "
                    "Return a short bulleted list (max 6 bullets).\n\nInput:" + q + "\n\nOutput:\n"
                )
                resp = openai.ChatCompletion.create(
                    model=os.environ.get('OPENAI_MODEL', 'gpt-4o-mini'),
                    messages=[{'role': 'user', 'content': prompt}],
                    max_tokens=300,
                    temperature=0.3,
                )
                text = resp['choices'][0]['message']['content'].strip()
                used_openai = True
                dlg = QtWidgets.QDialog(self)
                dlg.setWindowTitle('AI Suggestions (OpenAI)')
                dlg.resize(600, 400)
                lay = QtWidgets.QVBoxLayout(dlg)
                te = QtWidgets.QTextEdit()
                te.setReadOnly(True)
                te.setPlainText(text)
                lay.addWidget(te)
                btn = QtWidgets.QPushButton('OK')
                btn.clicked.connect(dlg.accept)
                lay.addWidget(btn)
                dlg.exec_()
                self.status.setText('AI suggestions (OpenAI) shown.')
                return
            except Exception as e:
                self.status.setText(f'OpenAI request failed, falling back: {e}')

        try:
            names_common = self.df[self.common_col].astype(str).tolist() if self.df is not None else []
            names_latin = self.df[self.latin_col].astype(str).tolist() if self.df is not None else []
            combined = list(dict.fromkeys(names_common + names_latin))
            if not combined:
                self.status.setText('No local data available for suggestions.')
                return
            found = process.extract(q, combined, scorer=fuzz.WRatio, limit=10)
            bullets = [f"- {name}  ({score}%)" for name, score, _ in found]
            text = "Local suggestions:\n\n" + "\n".join(bullets)
            dlg = QtWidgets.QDialog(self)
            dlg.setWindowTitle('AI Suggestions (Local)')
            dlg.resize(600, 400)
            lay = QtWidgets.QVBoxLayout(dlg)
            te = QtWidgets.QTextEdit()
            te.setReadOnly(True)
            te.setPlainText(text)
            lay.addWidget(te)
            btn = QtWidgets.QPushButton('OK')
            btn.clicked.connect(dlg.accept)
            lay.addWidget(btn)
            dlg.exec_()
            self.status.setText('AI suggestions (local) shown.')
        except Exception as e:
            self.status.setText(f'AI suggest failed: {e}')

    # ============================================================
    # Search
    # ============================================================
    def on_search(self):
        query = self.query.text().strip()
        self.table.setRowCount(0)
        if not query:
            self.status.setText('Type a remedy name.')
            return
        if self.df is None:
            self.status.setText('No data loaded.')
            return

        mode = self.mode_combo.currentText()
        count = 0
        seen = set()

        def add_row(idx):
            nonlocal count
            if idx in seen:
                return
            seen.add(idx)
            common = str(self.df.iloc[idx][self.common_col])
            latin = str(self.df.iloc[idx][self.latin_col])
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(common))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(latin))
            count += 1

        if mode == 'Fuzzy':
            names_common = self.df[self.common_col].astype(str).tolist()
            found_c = process.extract(query, names_common, scorer=fuzz.WRatio, limit=50)
            for name, score, idx in found_c:
                if score >= 60:
                    add_row(idx)
            names_latin = self.df[self.latin_col].astype(str).tolist()
            found_l = process.extract(query, names_latin, scorer=fuzz.WRatio, limit=50)
            for name, score, idx in found_l:
                if score >= 60:
                    add_row(idx)
        else:
            common_list = self.df['common_norm_cf'].astype(str).tolist()
            latin_list = self.df['latin_norm'].astype(str).tolist()
            q_norm = unicodedata.normalize('NFC', query).strip()
            tokens = [t for t in q_norm.split() if t]
            tokens_cf = [unicodedata.normalize('NFC', t).casefold() for t in tokens]

            if mode == 'Contains (Excel)':
                for idx, (c, l) in enumerate(zip(common_list, latin_list)):
                    if all((t in c) or (t in l) for t in tokens_cf):
                        add_row(idx)
            elif mode == 'Starts-with':
                q_cf = q_norm.casefold()
                for idx, (c, l) in enumerate(zip(common_list, latin_list)):
                    if c.startswith(q_cf) or l.startswith(q_cf):
                        add_row(idx)
            elif mode == 'Word-prefix':
                first = tokens_cf[0] if tokens_cf else ''
                for idx, (c, l) in enumerate(zip(common_list, latin_list)):
                    matched = any(w.startswith(first) for w in c.split()) or any(w.startswith(first) for w in l.split())
                    if matched:
                        add_row(idx)

        self.status.setText(f'Found {count} results ({mode}).' if count else f'No matches found ({mode}).')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoWindow()
    w.show()
    sys.exit(app.exec_())
