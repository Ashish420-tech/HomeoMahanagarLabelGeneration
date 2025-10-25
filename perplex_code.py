from kivy.app import App
from kivy.lang import Builder
from kivy.uix.boxlayout import BoxLayout
from kivy.properties import StringProperty
from kivy.core.window import Window
import os

# Try to import pandas for reading xlsx
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
            hint_text: 'Enter Common Name'
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
            text: 'Clear'
            on_release: root.clear_results()
'''

class HomeBox(BoxLayout):
    status_text = StringProperty('Ready. Load remedies.xlsx (with Common and Latin columns).')

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.df = None
        self.latin_col = 'Latin'
        self.common_col = 'Common'
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
        # Try to normalize and find column names flexibly
        columns = {c.lower().strip(): c for c in df.columns}
        lc = columns.get('latin', None)
        cc = columns.get('common', None)
        if not (lc and cc):
            raise ValueError("Excel file must have 'Latin' and 'Common' columns")
        self.df = df[[cc, lc]].dropna(how='any')
        self.latin_col = lc
        self.common_col = cc

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

    def clear_results(self):
        self.ids.results_box.clear_widgets()
        self.status_text = 'Cleared.'

    def on_search(self, text):
        from kivy.uix.label import Label
        self.ids.results_box.clear_widgets()
        query = (text or '').strip().lower()
        if not query:
            self.status_text = 'Type a common name.'
            return
        if self.df is not None:
            # Case-insensitive search in Common column
            results = self.df[self.df[self.common_col].astype(str).str.lower().str.contains(query, na=False)]
            if not results.empty:
                for _, row in results.iterrows():
                    common = str(row[self.common_col])
                    latin = str(row[self.latin_col])
                    label = Label(text=f'Common: {common}    Latin: {latin}', size_hint_y=None, height='32dp')
                    self.ids.results_box.add_widget(label)
                self.status_text = f'Found {len(results)} result(s).'
            else:
                self.status_text = 'No match found.'
        else:
            self.status_text = 'No data loaded.'

class HomeoSearchApp(App):
    def build(self):
        # Uncomment for mobile: Window.size = (360,640)
        Builder.load_string(KV)
        return HomeBox()

if __name__ == '__main__':
    HomeoSearchApp().run()
