# -*- coding: utf-8 -*-
import os
import zipfile

def make_zip(folder_path: str, zip_path: str) -> None:
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(folder_path):
            for f in files:
                if f.endswith(".zip"):
                    continue
                file_path = os.path.join(root, f)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname=arcname)
