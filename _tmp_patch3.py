# -*- coding: utf-8 -*-
path = 'app.py'
text = open(path, encoding='utf-8').read()

# constant
old = "SHEET_RAW = 'Data_Raw'\n"
new = "SHEET_RAW = 'Data_Raw'\nPREPARED_DATA_DIR = 'data_tes'\n"
if 'PREPARED_DATA_DIR =' not in text:
    assert old in text
    text = text.replace(old, new, 1)

# helpers before get_active_data_path
marker = 'def get_active_data_path():'
helpers = '''def ensure_prepared_data_dir():
    os.makedirs(PREPARED_DATA_DIR, exist_ok=True)


def list_prepared_excel_files():
    files = []
    if os.path.isfile(DEFAULT_DATA_PATH):
        files.append(DEFAULT_DATA_PATH)
    ensure_prepared_data_dir()
    try:
        for name in sorted(os.listdir(PREPARED_DATA_DIR)):
            if name.lower().endswith(('.xlsx', '.xls')):
                files.append(os.path.join(PREPARED_DATA_DIR, name))
    except Exception:
        pass
    seen = set()
    out = []
    for f in files:
        key = os.path.abspath(f)
        if key not in seen:
            seen.add(key)
            out.append(f)
    return out


def save_uploaded_excel(uploaded_file):
    ensure_prepared_data_dir()
    safe_name = os.path.basename(uploaded_file.name).strip().replace(' ', '_')
    if not safe_name:
        safe_name = 'data_uji.xlsx'
    if not safe_name.lower().endswith(('.xlsx', '.xls')):
        safe_name += '.xlsx'
    dest = os.path.join(PREPARED_DATA_DIR, safe_name)
    with open(dest, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    return dest


''' + marker
if 'def ensure_prepared_data_dir' not in text:
    assert marker in text
    text = text.replace(marker, helpers, 1)

open(path, 'w', encoding='utf-8').write(text)
print('ok')
