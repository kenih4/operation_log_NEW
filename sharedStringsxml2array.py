import xml.etree.ElementTree as ET
import glob
import csv

#   sharedStrings.xml のsiタグの部分だけ配列に格納
# python sharedStringsxml2array.py

sslist = []
xmls = glob.glob('sharedStrings.xml', recursive=True)
for xml in xmls:
    print("xml file=",xml)
    # XMLファイルパース
    tree = ET.parse(xml)
    root = tree.getroot()
#    print("root = ",root)
    for ssl in root:
#        print("ssl = ",ssl)
        for child in ssl.iter():
                   
            if child.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si':   # 特定要素(si)の抽出
#                print("child.tag = ", child.tag)
                
                for child2 in child.iter():
                     if child2.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t':
                        #print("Hit child2.text= ",child2.text)
                        sslist.append(child2.text)
                        break

for index, item in enumerate(sslist):
    print("インデックス：" + str(index) + ", 値：" + item)
