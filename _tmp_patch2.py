# -*- coding: utf-8 -*-
path = 'app.py'
lines = open(path, encoding='utf-8').readlines()

# 1) Replace get_training_data_file with three provider functions
start = None
for i, l in enumerate(lines):
    if l.startswith('def get_training_data_file()'):
        start = i
        break
if start is None:
    raise SystemExit('get_training_data_file not found')
end = start
while end < len(lines) and not (end > start and lines[end].startswith('@st.cache_data')):
    end += 1
new_funcs = '''def get_training_data_file():
    """File Data Training yang sudah disiapkan (80% transaksi)."""
    p = os.path.join(PREPARED_DATA_DIR, 'Data_Training.xlsx')
    return p if os.path.isfile(p) else DEFAULT_DATA_PATH


def get_testing_data_file():
    """File Data Testing yang sudah disiapkan (20% transaksi)."""
    p = os.path.join(PREPARED_DATA_DIR, 'Data_Testing.xlsx')
    return p if os.path.isfile(p) else None


def get_guide_file():
    """Buku panduan (docx) untuk pengunjung."""
    p = os.path.join(PREPARED_DATA_DIR, 'Buku_Panduan_UTS_BI.docx')
    if os.path.isfile(p):
        return p
    p2 = 'P8. UTS_Business Intelligence.docx'
    return p2 if os.path.isfile(p2) else None


'''
lines[start:end] = [new_funcs]

# 2) Replace action bar block
text = ''.join(lines)
start_marker = '    # --- Action bar data training (gaya pill seperti referensi) ---'
end_marker = '    with st.expander("'
s = text.index(start_marker)
e = text.index(end_marker, s)
new_block = '''    # --- Action bar data training / testing / panduan (gaya pill) ---
    _bg = find_asset(ASSET_KAMPUS)
    _bg_ext = __import__("os").path.splitext(_bg)[1].lstrip(".").lower() if _bg else ""
    _bg_uri = img_to_data_uri(_bg, _bg_ext or "jpg") if _bg else None
    train_path = get_training_data_file()
    test_path = get_testing_data_file()
    guide_path = get_guide_file()

    if _bg_uri:
        st.markdown(
            f"<style>.st-key-cover_data_hero{{--cover-data-bg:url('{_bg_uri}');}}</style>",
            unsafe_allow_html=True,
        )

    with st.container(key="cover_data_hero"):
        b1, b2, b3 = st.columns(3, gap="medium")
        with b1:
            if train_path and os.path.isfile(train_path):
                st.download_button(
                    "Data Training",
                    data=read_file_bytes(train_path, os.path.getmtime(train_path)),
                    file_name="Data_Training.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key="dl_data_training",
                )
            else:
                st.caption("Data Training belum tersedia.")
        with b2:
            if test_path and os.path.isfile(test_path):
                st.download_button(
                    "Data Testing",
                    data=read_file_bytes(test_path, os.path.getmtime(test_path)),
                    file_name="Data_Testing.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key="dl_data_testing",
                )
            else:
                st.caption("Data Testing belum tersedia.")
        with b3:
            if guide_path and os.path.isfile(guide_path):
                st.download_button(
                    "Buku Panduan",
                    data=read_file_bytes(guide_path, os.path.getmtime(guide_path)),
                    file_name="Buku_Panduan_UTS_BI.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    icon=":material/menu_book:",
                    key="dl_buku_panduan",
                )
            else:
                st.caption("Buku Panduan belum tersedia.")

'''
text = text[:s] + new_block + text[e:]

# 3) CSS: give columns gap, shrink min-width
text = text.replace(
    '  .st-key-cover_data_hero div[data-testid="stHorizontalBlock"] { max-width: 100%; margin: 0; gap: 0; }',
    '  .st-key-cover_data_hero div[data-testid="stHorizontalBlock"] { max-width: 100%; margin: 0; gap: 0.8rem; }',
)
text = text.replace(
    '    width: auto !important; min-width: 250px; max-width: 100%;',
    '    width: auto !important; min-width: 210px; max-width: 100%;',
)

# 4) Fix sidebar garbled line and upload buttons
lines = text.splitlines(keepends=True)
for i, l in enumerate(lines):
    if '**Sumber data:**' in l and l.lstrip().startswith('st.sidebar.markdown'):
        indent = l[:len(l) - len(l.lstrip())]
        lines[i] = indent + 'st.sidebar.markdown(f"📄 **Sumber data:** `{os.path.basename(ACTIVE_DATA_PATH)}`  \\n📊 **Data awal:** {rows_before:,} baris  \\n✅ **Data bersih:** {len(df):,} baris  \\n🗓️ **Snapshot RFM:** {snapshot_date.date()}  \\n🧾 **Transaksi unik:** {df[\'Transaction_ID\'].nunique():,}  \\n👥 **Pelanggan:** {df[\'Customer_ID\'].nunique():,}")\n'
    if 'Simpan Data' in l and 'st.button(' in l:
        indent = l[:len(l) - len(l.lstrip())]
        lines[i] = indent + 'save_clicked = st.button("💾 Simpan Data", use_container_width=True)\n'
    if 'Simpan & Pakai' in l and 'st.button(' in l:
        indent = l[:len(l) - len(l.lstrip())]
        lines[i] = indent + 'use_clicked = st.button("✅ Simpan & Pakai", use_container_width=True, type="primary")\n'

open(path, 'w', encoding='utf-8').writelines(lines)
print('ok')
