import xml.etree.ElementTree as ET
import glob
import csv

# python sheet1xml2text.py




# XMLファイル一覧取得
# 前提:同一階層のxmlsフォルダにxmlファイルを配置する
#xmls = glob.glob('xmls/*.xml', recursive=True)
xmls = glob.glob('sheet1.xml', recursive=True)


# ファイルごとに解析する
for xml in xmls:
    print("xml file=",xml)
    # XMLファイルパース
    tree = ET.parse(xml)
    root = tree.getroot()
#    print("root = ",root)
    for sheetData in root:
#        print("sheetData = ",sheetData)
        for child in sheetData.iter():
                   
            if child.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row':   # 特定要素(row)の抽出
                print("")
                
                CELL = "-"
                for child2 in child.iter():
                    if child2.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c':
#                        print("Hit child2.attrib["r"] = ",child2.attrib["r"])
                        if not child2.attrib["r"].find('A') or not child2.attrib["r"].find('B') or not child2.attrib["r"].find('C'):
                            #print("Hit child2.attrib["r"] = ",child2.attrib["r"])
                            CELL = child2.attrib["r"]
                            VAL = "-"
                            for child3 in child2.iter():
                                if child3.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v':
                                    #print("Hit child3.tag =   ",child3.text)
                                    VAL = child3.text
                            print("\t",CELL,"\t",VAL,end='')
                    
                    
#                    print("Hit child.tag = ",child.tag)
#                print("Hit child.tag = ",child.tag)
                # ファイル名、国名、ランクを出力
#                row = child.text
#                print("row = ",row)
#                rank = child.text
#                print(f'{xml},{row}')
#                cdata = [xml, name, rank]
#                cdata_list.append(cdata)

# CSVファイルに保存
#with open('tmp.csv', 'w', newline="") as f:
#    writer = csv.writer(f)
#    writer.writerows(cdata_list)