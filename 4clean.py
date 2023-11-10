""" 清空
    data/1word
    data/2pdf
    data/3pdf_sign
    data/4pdf_sign_score
"""

import shutil
import os

dirpath_list = [
    'data/1word',
    'data/2pdf',
    'data/3pdf_sign',
    'data/4pdf_sign_score'
]

for dirpath in dirpath_list:
    shutil.rmtree(dirpath)
    os.mkdir(dirpath)
