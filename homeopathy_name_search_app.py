"""
Homeopathy Name Search - Kivy app (single-file)

How it works
- Put your Excel file (remedies.xlsx) in the same folder as this script.
- Excel must contain two columns with headers (case-insensitive): "Latin" and "Common".
  If your file has different headers, update the `latin_col` and `common_col` variables below.
- Run on desktop to test: `python homeopathy_name_search_app.py`
- To build an Android APK, use Buildozer. Note: pandas and openpyxl can bloat APK size; CSV is smaller.

Dependencies
- Python 3.8+
- kivy
- pandas
- openpyxl (if using .xlsx)

Install: pip install kivy pandas openpyxl

This file contains both the app and a small helper to load/parse the Excel.
"""

from kivy.app import App
from kivy.lang import Builder
from kivy.uix.boxlayout import BoxLayout
from kivy.properties import StringProperty, ListProperty
from kivy.core.window import Window
import os

# We'll try to import pandas; if not available the app will show an error message
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except Exception:
    PANDAS_AVAILABLE = False

KV = '''
<HomeBox>:
    orientation: 'vertical'
    padding: 12
    spacing: 10

    BoxLayout:
        size_hint_y: None
        height: '40dp'
        TextInput:
            id: query
            hint_text: 'Type Latin or common name (e.g. arsenicum album or arsenic)'
            multiline: False
            on_text_validate: root.on_search(query.text)
        Button:
            text: 'Search'
            size_hint_x: None
            width: '100dp'
            on_release: root.on_search(query.text)

    BoxLayout:
        size_hint_y: None
        height: '36dp'
        Label:
            text: 'Direction:'
            size_hint_x: None
            width: '80dp'
            halign: 'left'
            valign: 'middle'
        Spinner:
            id: direction
            text: 'Auto (detect)'
            values: ['Auto (detect)', 'Latin → Common', 'Common → Latin']
            size_hint_x: None
            width: '200dp'

    Label:
        id: status
        size_hint_y: None
        height: '22dp'
        text: root.status_text
        halign: 'left'
        valign: 'middle'

    ScrollView:
        GridLayout:
            id: results_box
            cols: 1
            size_hint_y: None
            height: self.minimum_height
            row_default_height: '40dp'
            row_force_default: False

    BoxLayout:
        size_hint_y: None
        height: '36dp'
        spacing: 10
        Button:
            text: 'Load Excel'
            on_release: root.load_excel()
        Button:
            text: 'Clear'
            on_release: root.clear_results()
'''

class HomeBox(BoxLayout):
    status_text = StringProperty('Ready. Load your Excel (remedies.xlsx) or click Load Excel.')
    last_query = StringProperty('')
    results = ListProperty([])

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.df = None
        self.latin_col = 'Latin'   # change if your Excel has different headers
        self.common_col = 'Common' # change if your Excel has different headers
        self.m_lat_to_common = {}
        self.m_common_to_lat = {}
        if PANDAS_AVAILABLE:
            # try to auto-load default file if present
            if os.path.exists('remedies.xlsx'):
                try:
                    self._load_df('remedies.xlsx')
                    self.status_text = 'Loaded remedies.xlsx successfully.'
                except Exception as e:
                    self.status_text = f'Failed to load remedies.xlsx: {e}'
        else:
            self.status_text = 'pandas not installed. Install pandas & openpyxl to load Excel.'

    def _load_df(self, path):
        # read excel into dataframe and normalize column names
        if not PANDAS_AVAILABLE:
            raise RuntimeError('pandas not available')
        df = pd.read_excel(path, engine='openpyxl')
        # normalize column names (strip & lower)
        cols = {c: c.strip() for c in df.columns}
        df.rename(columns=cols, inplace=True)

        # find columns by case-insensitive match
        lc = None
        cc = None
        for c in df.columns:
            if c.strip().lower() == self.latin_col.lower():
                lc = c
            if c.strip().lower() == self.common_col.lower():
                cc = c
        # If not found by exact name, try heuristics
        if lc is None:
            for c in df.columns:
                if 'latin' in c.strip().lower():
                    lc = c
                    break
        if cc is None:
            for c in df.columns:
                if 'common' in c.strip().lower() or 'english' in c.strip().lower() or 'name'==c.strip().lower():
                    cc = c
                    break
        if lc is None or cc is None:
            raise ValueError(f'Could not detect Latin/Common columns automatically. Found columns: {list(df.columns)}')

        self.df = df[[lc, cc]].copy()
        self.df.dropna(how='all', inplace=True)

        # build lowercase lookup mappings
        self.m_lat_to_common = {}
        self.m_common_to_lat = {}
        for _, row in self.df.iterrows():
            lat = str(row[lc]).strip()
            com = str(row[cc]).strip()
            if lat:
                self.m_lat_to_common[lat.lower()] = com
            if com:
                self.m_common_to_lat[com.lower()] = lat

    def load_excel(self):
        # simple loader: look for remedies.xlsx in current folder
        if not PANDAS_AVAILABLE:
            self.status_text = 'pandas not installed. Please install pandas & openpyxl.'
            return
        path = 'remedies.xlsx'
        if not os.path.exists(path):
            # also try remedies.csv
            if os.path.exists('remedies.csv'):
                path = 'remedies.csv'
            else:
                self.status_text = 'No remedies.xlsx or remedies.csv found in app folder.'
                return
        try:
            self._load_df(path)
            self.status_text = f'Loaded {path} with {len(self.df)} rows.'
        except Exception as e:
            self.status_text = f'Error loading file: {e}'

    def clear_results(self):
        self.ids.results_box.clear_widgets()
        self.status_text = 'Cleared results.'

    def on_search(self, text):
        query = (text or '').strip()
        self.last_query = query
        self.ids.results_box.clear_widgets()
        if not query:
            self.status_text = 'Type something to search.'
            return
        if not self.df is None:
            direction = self.ids.direction.text
            # detect direction
            use_lat_to_common = None
            ql = query.lower()
            if direction == 'Latin → Common':
                use_lat_to_common = True
            elif direction == 'Common → Latin':
                use_lat_to_common = False
            else:
                # auto-detect: if exact match in latin map -> treat as Latin
                if ql in self.m_lat_to_common:
                    use_lat_to_common = True
                elif ql in self.m_common_to_lat:
                    use_lat_to_common = False
                else:
                    # fallback to searching both
                    use_lat_to_common = None

            results = []
            if use_lat_to_common is True:
                # exact match
                if ql in self.m_lat_to_common:
                    results.append((query, self.m_lat_to_common[ql]))
                else:
                    # partial matches in latin column
                    matches = self.df[self.df.iloc[:,0].str.lower().str.contains(ql, na=False)]
                    for _, r in matches.iterrows():
                        results.append((str(r.iloc[0]), str(r.iloc[1])))
            elif use_lat_to_common is False:
                if ql in self.m_common_to_lat:
                    results.append((self.m_common_to_lat[ql], query))
                else:
                    matches = self.df[self.df.iloc[:,1].str.lower().str.contains(ql, na=False)]
                    for _, r in matches.iterrows():
                        results.append((str(r.iloc[0]), str(r.iloc[1])))
            else:
                # search both columns for partial matches
                matches_lat = self.df[self.df.iloc[:,0].str.lower().str.contains(ql, na=False)]
                for _, r in matches_lat.iterrows():
                    results.append((str(r.iloc[0]), str(r.iloc[1])))
                matches_com = self.df[self.df.iloc[:,1].str.lower().str.contains(ql, na=False)]
                for _, r in matches_com.iterrows():
                    tup = (str(r.iloc[0]), str(r.iloc[1]))
                    if tup not in results:
                        results.append(tup)

            if not results:
                # try fuzzy contains by splitting words
                parts = ql.split()
                for p in parts:
                    if not p: continue
                    matches = self.df[self.df.iloc[:,0].str.lower().str.contains(p, na=False)]
                    for _, r in matches.iterrows():
                        tup = (str(r.iloc[0]), str(r.iloc[1]))
                        if tup not in results:
                            results.append(tup)
                    matches = self.df[self.df.iloc[:,1].str.lower().str.contains(p, na=False)]
                    for _, r in matches.iterrows():
                        tup = (str(r.iloc[0]), str(r.iloc[1]))
                        if tup not in results:
                            results.append(tup)

            # show results
            if results:
                from kivy.uix.label import Label
                for lat, com in results:
                    lbl = Label(text=f'[b]Latin:[/b] {lat}    [b]Common:[/b] {com}', markup=True, size_hint_y=None, height='36dp')
                    self.ids.results_box.add_widget(lbl)
                self.status_text = f'Found {len(results)} result(s).'
            else:
                self.status_text = 'No matches found.'
        else:
            self.status_text = 'No data loaded. Click Load Excel and ensure remedies.xlsx is present.'

class HomeApp(App):
    def build(self):
        Window.size = (420, 720)  # good for desktop testing
        Builder.load_string(KV)
        return HomeBox()

if __name__ == '__main__':
    HomeApp().run()
