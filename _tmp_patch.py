# -*- coding: utf-8 -*-
path = 'app.py'
text = open(path, encoding='utf-8').read()
old = '  .st-key-cover_data_hero div[data-testid="stHorizontalBlock"] { max-width: 980px; margin: 0 auto; gap: 0.85rem; }\n  .st-key-cover_data_hero div[data-testid="stDownloadButton"] { display:flex; justify-content:center; }'
new = '  .st-key-cover_data_hero div[data-testid="stHorizontalBlock"] { max-width: 980px; margin: 0 auto; gap: 0.85rem; }\n  .st-key-cover_data_hero div[data-testid="stElementContainer"] { width: 100% !important; }\n  .st-key-cover_data_hero div[data-testid="stDownloadButton"] { display:flex; justify-content:center; width:100%; }'
assert old in text
open(path, 'w', encoding='utf-8').write(text.replace(old, new, 1))
print('ok')
