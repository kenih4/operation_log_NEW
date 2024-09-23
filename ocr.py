from PIL import Image
import pyocr
import glob
import os
import sys


#   https://qiita.com/eiji-noguchi/items/c19c1e125eaa87c3616b
# 英語のみ
# python ocr.py


# tesseract.exeをインストールする必要あり　https://github.com/UB-Mannheim/tesseract/wiki
pyocr.tesseract.TESSERACT_CMD = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'


args = sys.argv
print("filedir:\t",args[1])


engines = pyocr.get_available_tools()
print(engines)
if len(engines) == 0:
    print("OCRツールが見つかりませんでした")
    sys.exit(1)
engine = engines[0]


currentDir = os.getcwd()
#filedir = 'C:/Users/kenichi/pictures/media'
files = glob.glob(args[1]+'/*.png', recursive=True)
for file in files:
    print(file)
    txt = engine.image_to_string(Image.open(file), lang="eng")
    print(txt.encode('cp932', "ignore"))
#input()

