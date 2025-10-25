from kivy.app import App
from kivy.lang import Builder
from kivy.uix.boxlayout import BoxLayout
from kivy.properties import StringProperty
from kivy.core.text import LabelBase
import os
import glob
import urllib.request
import shutil
import socket
from kivy.uix.popup import Popup
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout as UiBox

# Try to locate the Bengali font file in the project (or subfolders) and register it
FONT_FILE = 'NotoSansBengali-Regular.ttf'
BENGALI_FONT_AVAILABLE = False
FONT_NAME = None
FONT_PATH_USED = None

# Build candidate paths: same directory as this file and a recursive search
base_dir = os.path.dirname(os.path.abspath(__file__))
candidates = [os.path.join(base_dir, FONT_FILE)]
candidates += glob.glob(os.path.join(base_dir, '**', FONT_FILE), recursive=True)
candidates += glob.glob(os.path.join('.', '**', FONT_FILE), recursive=True)

font_path = None
for p in candidates:
    if p and os.path.exists(p):
        font_path = os.path.abspath(p)
        break

if font_path:
    try:
        LabelBase.register(name='Bengali', fn_regular=font_path)
        BENGALI_FONT_AVAILABLE = True
        FONT_NAME = 'Bengali'
        FONT_PATH_USED = font_path
    except Exception:
        BENGALI_FONT_AVAILABLE = False
        FONT_NAME = None
else:
    # No font found locally; attempt to download a known free Bengali font (Noto Sans Bengali)
    try:
        download_url = 'https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSansBengali/NotoSansBengali-Regular.ttf'
        target = os.path.join(base_dir, FONT_FILE)
        # short network timeout
        socket.setdefaulttimeout(10)
        urllib.request.urlretrieve(download_url, target)
        # If download succeeded, register it
        if os.path.exists(target):
            try:
                LabelBase.register(name='Bengali', fn_regular=target)
                BENGALI_FONT_AVAILABLE = True
                FONT_NAME = 'Bengali'
                FONT_PATH_USED = os.path.abspath(target)
            except Exception:
                BENGALI_FONT_AVAILABLE = False
                FONT_NAME = None
                FONT_PATH_USED = None
        else:
            BENGALI_FONT_AVAILABLE = False
            FONT_NAME = None
            FONT_PATH_USED = None
    except Exception:
        # Network/download failed — fall back to default font
        BENGALI_FONT_AVAILABLE = False
        FONT_NAME = None
        FONT_PATH_USED = None

# (Font registration handled above with a safe lookup; no unconditional register here.)

# 2. Fuzzy search package
from rapidfuzz import fuzz, process

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except Exception:
    PANDAS_AVAILABLE = False

KV = '''
<HomeBox>:
    orientation: 'vertical'
    padding: 16
    spacing: 10

    BoxLayout:
        size_hint_y: None
        height: '48dp'
        TextInput:
            id: query
            hint_text: 'Enter Common Name (English or Bengali)'
            multiline: False
            on_text_validate: root.on_search(query.text)
        Button:
            text: 'Search'
            size_hint_x: None
            width: '100dp'
            on_release: root.on_search(query.text)

    Label:
        id: status
        size_hint_y: None
        height: '24dp'
        text: root.status_text
        halign: 'left'
        valign: 'middle'

    Label:
        id: font_status
        size_hint_y: None
        height: '20dp'
        text: root.font_status_text
        halign: 'left'
        valign: 'middle'

    BoxLayout:
        id: results_box
        orientation: 'vertical'
        size_hint_y: 1

    BoxLayout:
        size_hint_y: None
        height: '40dp'
        spacing: 10
        Button:
            text: 'Load Excel'
            on_release: root.load_excel()
        Button:
            text: 'Load Font'
            on_release: root.open_font_chooser()
        Button:
            text: 'Clear'
            on_release: root.clear_results()
'''

class HomeBox(BoxLayout):
    status_text = StringProperty('Ready. Load remedies.xlsx (latin_col, common_col).' )
    font_status_text = StringProperty('')

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        # popup instance placeholder for font chooser
        self._font_popup = None
        self.df = None
        self.latin_col = 'latin_col'
        self.common_col = 'common_col'
        # config file to persist chosen font
        self._font_cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'font.cfg')

        # if a font was previously registered via automatic lookup, reflect it
        if BENGALI_FONT_AVAILABLE and FONT_PATH_USED:
            self.font_status_text = f'Font: {FONT_PATH_USED}'
        else:
            # try loading persisted font path
            try:
                if os.path.exists(self._font_cfg):
                    with open(self._font_cfg, 'r', encoding='utf-8') as f:
                        saved = f.read().strip()
                        if saved:
                            if self.register_font(saved):
                                # register_font updates font_status_text
                                pass
                            else:
                                self.font_status_text = 'No font loaded.'
                        else:
                            self.font_status_text = 'No font loaded.'
                else:
                    self.font_status_text = 'No font loaded.'
            except Exception:
                self.font_status_text = 'No font loaded.'
        if PANDAS_AVAILABLE:
            if os.path.exists('remedies.xlsx'):
                try:
                    self._load_df('remedies.xlsx')
                    self.status_text = 'Loaded remedies.xlsx.'
                except Exception as e:
                    self.status_text = f'Failed to load: {e}'
        else:
            self.status_text = 'pandas library not installed.'

    def _load_df(self, path):
        if not PANDAS_AVAILABLE:
            raise RuntimeError('pandas not available')
        df = pd.read_excel(path, engine='openpyxl')
        df.rename(columns={c: c.lower().strip() for c in df.columns}, inplace=True)
        self.df = df[['latin_col', 'common_col']].dropna(how='any')

    def load_excel(self):
        if not PANDAS_AVAILABLE:
            self.status_text = 'pandas not installed.'
            return
        path = 'remedies.xlsx'
        if not os.path.exists(path):
            self.status_text = 'No remedies.xlsx found.'
            return
        try:
            self._load_df(path)
            self.status_text = f'Loaded {path} with {len(self.df)} entries.'
        except Exception as e:
            self.status_text = f'Load error: {e}'

    def register_font(self, path):
        """Attempt to register a font file at runtime and update status."""
        global BENGALI_FONT_AVAILABLE, FONT_NAME, FONT_PATH_USED
        try:
            if not path or not os.path.exists(path):
                self.status_text = 'Font file not found.'
                return False
            LabelBase.register(name='Bengali', fn_regular=path)
            BENGALI_FONT_AVAILABLE = True
            FONT_NAME = 'Bengali'
            FONT_PATH_USED = os.path.abspath(path)
            self.status_text = f'Registered font: {FONT_PATH_USED}'
            self.font_status_text = f'Font: {FONT_PATH_USED}'
            # persist the font path
            try:
                with open(self._font_cfg, 'w', encoding='utf-8') as f:
                    f.write(FONT_PATH_USED)
            except Exception:
                pass
            return True
        except Exception as e:
            BENGALI_FONT_AVAILABLE = False
            FONT_NAME = None
            FONT_PATH_USED = None
            self.status_text = f'Font register failed: {e}'
            self.font_status_text = 'No font loaded.'
            return False

    def open_font_chooser(self):
        """Open a FileChooser popup to pick a .ttf font file and register it."""
        # prevent multiple popups
        if getattr(self, '_font_popup', None) and getattr(self, '_font_popup', 'open'):
            return
        content = UiBox(orientation='vertical', spacing=6)
        filechooser = FileChooserListView(filters=['*.ttf', '*.otf'], path='.')
        btn_box = UiBox(size_hint_y=None, height='40dp', spacing=6)
        btn_select = Button(text='Select', size_hint_x=None, width='120dp')
        btn_cancel = Button(text='Cancel', size_hint_x=None, width='120dp')
        btn_box.add_widget(btn_select)
        btn_box.add_widget(btn_cancel)
        content.add_widget(filechooser)
        content.add_widget(btn_box)

        popup = Popup(title='Choose font file', content=content, size_hint=(0.9, 0.9))
        self._font_popup = popup

        def _on_select(instance):
            selection = filechooser.selection
            if selection:
                picked = selection[0]
                ok = self.register_font(picked)
                popup.dismiss()

        def _on_cancel(instance):
            popup.dismiss()

        btn_select.bind(on_release=_on_select)
        btn_cancel.bind(on_release=_on_cancel)
        popup.open()

    def clear_results(self):
        self.ids.results_box.clear_widgets()
        self.status_text = 'Cleared.'

    def on_search(self, text):
        from kivy.uix.label import Label
        self.ids.results_box.clear_widgets()
        query = (text or '').strip()
        if not query:
            self.status_text = 'Type a remedy name.'
            return
        if self.df is not None:
            # AI fuzzy search on common_col (Bengali/English, all text)
            names = self.df[self.common_col].astype(str).tolist()
            # Get up to 5 best matches above 60% similarity
            found = process.extract(query, names, scorer=fuzz.WRatio, limit=5)
            count = 0
            for name, score, idx in found:
                if score >= 60:
                    common = str(self.df.iloc[idx][self.common_col])
                    latin = str(self.df.iloc[idx][self.latin_col])
                    # Build a two-column row: left = common (Bengali/English), right = latin
                    from kivy.uix.boxlayout import BoxLayout as KBox
                    row = KBox(orientation='horizontal', size_hint_y=None, height='32dp', spacing=6)

                    def contains_bengali(s):
                        if not s:
                            return False
                        for ch in s:
                            if '\u0980' <= ch <= '\u09FF':
                                return True
                        return False

                    # Common name label: use Bengali font when the text contains Bengali characters
                    common_kwargs = {'text': f'কমন: {common}', 'size_hint_x': 0.6}
                    if BENGALI_FONT_AVAILABLE and contains_bengali(common):
                        common_kwargs['font_name'] = FONT_NAME
                    common_lbl = Label(**common_kwargs)

                    # Latin name label: use default font (keeps Latin glyphs rendered as usual)
                    latin_lbl = Label(text=f'লাতিন: {latin}', size_hint_x=0.4)

                    row.add_widget(common_lbl)
                    row.add_widget(latin_lbl)
                    self.ids.results_box.add_widget(row)
                    count += 1
            if count:
                self.status_text = f'AI found {count} remedy/remedies.'
            else:
                self.status_text = 'No close matches found.'
        else:
            self.status_text = 'No data loaded.'

class HomeoSearchApp(App):
    def build(self):
        Builder.load_string(KV)
        return HomeBox()

if __name__ == '__main__':
    HomeoSearchApp().run()
