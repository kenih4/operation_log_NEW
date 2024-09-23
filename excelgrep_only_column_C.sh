#!/bin/bash

#Usage:
# https://qiita.com/UKIUKI_ENGINEER/items/76d1ba94c2e210bc5f5d
#
## ./excelgrep_only_column_C.sh 'dcct' SP8/*.xlsm
#
#   TEST用
#   
# ./excelgrep_only_column_C.sh 'dcct' SP8/2024_06_SP8.xlsm
# 

# 引数からgrepの操作内容を取り出す
for arg in ${@}; do
  if [ $(echo ${arg} | grep -vE '\.xlsm$|\.xls$') ]; then
    # FIXME シングルクォーテーションが消えてしまう
    targetstr=${targetstr}" "${arg}
  fi
done
echo targetstr = ${targetstr}


# 一つ一つのxlsmファイルに対してgrepする
for arg in ${@}; do
  # *.xlsm以外の引数の場合次のループへ
  if [ $(echo ${arg} | grep -vE '\.xlsm$|\.xls$') ]; then
    continue
  fi

#メモ：　grep 指定した文字列を含まない行を抽出するためにはgrepの-vオプションを用います。
#メモ：　ハット（^）は「～で始まる」、ドル記号（$）は「～で終わる」を意味します

#  「~$2024_06_SP8.xlsm」のような一時ファイルは除く
  if [ $(echo ${arg} | grep '~') ]; then
    continue
  fi

  echo 
  echo excel.exe ${arg} '&   ---------------------------------------------------------------------------'


#  read -p "Hit enter: "
  
  # zip展開用一時ディレクトリ作成
  tmpdir=$(mktemp -d)
  # 一時ファイル作成
  tmp_sheet1A=$(mktemp)
  tmp_sheet1C=$(mktemp)
  tmp_sharedStrings=$(mktemp)

  echo tmpdir = ${tmpdir}
  echo mktemp = $(mktemp)
  echo tmp_sheet1A = ${tmp_sheet1A}
  echo tmp_sheet1C = ${tmp_sheet1C}
  echo tmp_sharedStrings = ${tmp_sharedStrings}

#read -p "Hit enter: "  

  # エクセルファイルを一時ディレクトリに解凍する
  # 標準出力とエラー出力は鬱陶しいので捨てる
  unzip ${arg} -d ${tmpdir} 1> /dev/null 2>&1

#メモ：　　find -type f:  検索対象のファイルタイプを指定して検索する  ファイルを検索するときf ディレクトリを検索するときd
#メモ：　　grep -E 　検索に拡張正規表現を使う
#メモ：　　xargs -I により、標準入力より渡されたデータを、 xargs引数コマンド の、任意の位置の引数に展開することが可能です。

<< COMMENTOUT   
find ${tmpdir} -type f |                               #対象のディレクトリのファイルを抽出 -type f はファイル
    grep -E 'xml$' |                              # .xmlファイルだけピックアップ
    xargs -I{} grep -iE ${targetstr} {} |             #  先ほどピックアップしたxmlファイルから検索対象文字列が含まれる部分だけ抽出
    perl -pe 's/<rPh.*?<\/rPh>//g' |                              # rPh タグで囲まれた所を削除 カタカナ部分
     perl -pe 's/<phoneticPr fontId=.*?\/>//g'  |                              # phoneticPr タグで囲まれた所を削除 カタカナ部分
      sed -e 's/>/>\n/g' |                             
       sed -e 's/<[^>]*>//g' |                              # タグ削除
        grep -v ^$ |                              #空行削除
         grep --color -n -3 -iE ${targetstr} 
COMMENTOUT






# sharedStrings.xmlの<si>タグだけピックアップ
# CRLFを削除、<sst>タグ内を抽出、　<rPh>タグを削除、<siで改行、タグを全て削除、1行目削除、指定した行のみ抽出
find ${tmpdir}  -type f | 
grep -E 'sharedStrings.xml$' | 
xargs -I{} perl -pe 's/\r\n//g'  {} | 
grep -oP '<sst.*?</sst>' | 
perl -pe 's/<rPh.*?<\/rPh>//g' | 
sed -e 's/<si/\n<si/g' | 
sed -e 's/<[^>]*>//g' | 
sed -e '1d' > ${tmp_sharedStrings}


#　作り掛け
# sheet1.xmlのA列（日付部分）だけ抽出
# $ find /tmp/tmp.leMDwt5SBY -type f | grep -E 'sheet1.xml$' | xargs -I{} grep -oP '<c r="A.*?</c>'  {} 
#find ${tmpdir} -type f | 
#grep -E 'sheet1.xml$' | 
#xargs -I{} grep -oP '<c r="A.*?</c>'  {} > ${tmp_sheet1A}
#perl Pickup_line_from_sheet1_A.pl ${tmp_sheet1A}


# sheet1.xmlのC列（内容部分）だけ抽出
find ${tmpdir} -type f | 
grep -E 'sheet1.xml$' | 
xargs -I{} grep -oP '<c r="C.*?</c>'  {} > ${tmp_sheet1C}


# sheet1.xmlのC列にsharedStrings.xmlの行番号が書かれているので探す
perl Pickup_line_from_sheet1.pl ${tmp_sheet1C} ${tmp_sharedStrings} | grep --color -n -A 2 -iE ${targetstr}



#read -p "Hit enter: "


  # 一時ディレクトリとファイルを削除
  rm -r ${tmpdir}
  rm ${tmp_sheet1A}
  rm ${tmp_sheet1C}  
  rm ${tmp_sharedStrings}

done