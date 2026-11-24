# -*- coding: utf-8 -*-
import os, shutil
import numpy as np
import pandas as pd

src = 'P8. UTS_Business Intelligence.xlsx'
out_dir = 'data_tes'
os.makedirs(out_dir, exist_ok=True)
sheet = 'Data_Raw'

df = pd.read_excel(src, sheet_name=sheet)
print('rows', len(df))

if 'Transaction_ID' in df.columns:
    ids = pd.Series(df['Transaction_ID'].dropna().unique())
    rng = np.random.default_rng(42)
    ids_arr = ids.to_numpy().copy()
    rng.shuffle(ids_arr)
    n_train = int(len(ids_arr) * 0.8)
    train_ids = set(ids_arr[:n_train].tolist())
    train = df[df['Transaction_ID'].isin(train_ids)].copy()
    test = df[~df['Transaction_ID'].isin(train_ids)].copy()
else:
    train = df.sample(frac=0.8, random_state=42)
    test = df.drop(train.index)

train_path = os.path.join(out_dir, 'Data_Training.xlsx')
test_path = os.path.join(out_dir, 'Data_Testing.xlsx')
train.to_excel(train_path, sheet_name=sheet, index=False)
test.to_excel(test_path, sheet_name=sheet, index=False)

guide_src = 'P8. UTS_Business Intelligence.docx'
guide_dst = os.path.join(out_dir, 'Buku_Panduan_UTS_BI.docx')
if os.path.isfile(guide_src):
    shutil.copyfile(guide_src, guide_dst)

print('train rows', len(train), 'test rows', len(test))
print('train size', os.path.getsize(train_path), 'test size', os.path.getsize(test_path))
