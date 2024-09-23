#!/bin/bash

#Usage:
# https://qiita.com/UKIUKI_ENGINEER/items/76d1ba94c2e210bc5f5d
#
#   TEST用
#   ./excelgrep.sh -iE 'dcct' SP8/2024_06_SP8.xlsm
#
# ./excelgrep.sh -iE '' SP8/2024_06_SP8.xlsm


# 引数からgrepの操作内容を取り出す
for arg in ${@}; do
  if [ $(echo ${arg} | grep -vE '\.xlsm$|\.xls$') ]; then
    # FIXME シングルクォーテーションが消えてしまう
    targetstr=${targetstr}" "${arg}
#    echo ${arg}
  fi
done

echo targetstr = ${targetstr}

#read -p "Hit enter: "

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
  echo excel.exe ${arg} '&'


#  read -p "Hit enter: "
  
  # zip展開用一時ディレクトリ作成
  tmpdir=$(mktemp -d)
  # grepの標準出力用一時ファイル作成
  tmpgrep=$(mktemp)
  
  echo tmpdir = ${tmpdir}
#  echo mktemp = $(mktemp)

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

#<< COMMENTOUT    #別方法
find ${tmpdir} -type f |   
grep -E 'xml$' | 
xargs -I{} grep -l -iE ${targetstr}  {} |      #先ほどピックアップしたxmlファイルから検索対象文字列が含まれるxmlファイルをさらにピックアップ
xargs -I{} perl -pe 's/<rPh.*?<\/rPh>//g' {}  |
perl -pe 's/<phoneticPr fontId=.*?\/>//g'  | 
sed -e 's/>/>\n/g' | 
sed -e 's/<[^>]*>//g' |  
grep -v ^$ | 
grep --color -n -3 -iE ${targetstr} 
#COMMENTOUT



<< COMMENTOUT    #TEST     ./excelgrep.sh '' SP8/2024_06_SP8.xlsm
find ${tmpdir} -type f |   
grep -E 'sheet1.xml$' | 
xargs -I{} grep -oP '<c r="C.*?</c>'  {}    #sheet1.xmlのC列だけ抽出
COMMENTOUT



<< COMMENTOUT #元
  find ${tmpdir} -type f |
    grep -E 'xml$' |
    xargs -I{} grep -iE ${targetstr} {} |
    tr '\"' '\n' |
    tr '>' '\n' |
    sed -e 's/<\/.*//' |
    grep ${targetstr} > ${tmpgrep}

echo tmpgrep = ${tmpgrep}
  # 検索文字列があればファイル名とその内容を出力する
  if [ $? = 0 ]; then
    # FIXME
    # awkを使って出力しているため、grepの--color=autoが効かない
    # echo tmpgrep = ${tmpgrep} 
    cat ${tmpgrep} | awk '{print filename": "$0}' filename=${arg}
  fi
COMMENTOUT

#read -p "Hit enter: "


  # 一時ディレクトリとファイルを削除
  rm -r ${tmpdir}
  rm ${tmpgrep}

done
