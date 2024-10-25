#!/bin/bash

#Usage:
# https://qiita.com/UKIUKI_ENGINEER/items/76d1ba94c2e210bc5f5d
#
## ./excelgrep_by_XMLparse.sh 'dcct' SP8/*.xlsm
#
#   TEST用
#   
# ./excelgrep_by_XMLparse.sh 'dcct' SP8/2024_06_SP8.xlsm
# ./excelgrep_by_XMLparse.sh 'CB0{0,1}8.1' SACLA/2024_06.xlsm
# 

echo Argument:   ${@}

# 引数からgrepの操作内容を取り出す
i=0
for arg in ${@}; do
  echo arg:   $arg
  
  if [ $(echo ${arg} | grep -vE '\.xlsm$|\.xls$') ]; then
    # FIXME シングルクォーテーションが消えてしまう
#    targetstr=${targetstr}" "${arg}
    targets[i]=${arg}
    i=`expr $i + 1`
  fi
done

for i in ${!targets[@]}
do
  echo target[$i] = ${targets[$i]}
  if [ $i -eq 0 ]; then
      targetstr=${targets[0]}
  else
      targetstr=${targetstr}"|"${targets[$i]}
  fi
done
echo targetstr = ${targetstr}

#exit


# 一つ一つのExcelファイルに対してgrepする
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
#  tmp_out=$(mktemp)


  echo tmpdir = ${tmpdir}
#  echo tmp_out = ${tmp_out}

#read -p "Hit enter: "  

  # エクセルファイルを一時ディレクトリに解凍する
  # 標準出力とエラー出力は鬱陶しいので捨てる
  unzip -t "${arg}" > error.log #  2> error.log標準エラー出力のみ
  if [ $? -ne 0 ]; then # $? は、直前に実行したコマンドの終了ステータス
      echo "ZIPファイルは異常です。"
      #cat error.log
      # 特定のエラーメッセージに基づく処理
      if grep -q "End-of-central-directory signature not found" error.log; then
          echo "指定されたファイルはZIP形式ではないか、壊れている可能性があります。"
      fi
      exit
#  else
#      echo "ZIPファイルは正常です。"
  fi
  unzip ${arg} -d ${tmpdir} 1> /dev/null 2>&1 # -d ディレクトリ	指定したディレクトリに展開する



ret=${tmpdir/\/tmp\//C:\\Users\\kenichi\\AppData\\Local\\Temp\\}
#echo start \"$ret\\xl\\media\"
if [ ! -e $ret\\xl\\media ]; then
  echo "Directory doesn't exists!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! May be unzip fail..."
  exit
fi
#   /tmp/tmp.KBjrD6k7Uq/xl/worksheets/sheet1.xml
#   /tmp/tmp.KBjrD6k7Uq/xl/sharedStrings.xml 

#python excelgrep_by_XMLparse.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml
#python excelgrep_by_XMLparse.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml | GREP_COLOR='0;33' grep -a --color -n -A 0 -iE ${targetstr}
#python excelgrep_by_XMLparse.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml > ${tmp_out}


#集計用テスト  色を付けるワードはpythonスクリプトにべた書き
python excelgrep_by_XMLparse_for_Untenshyukei.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml | grep -v -E '引渡し前|引渡し時|引渡し希望|引渡し後|引渡しが|引渡す|引渡て|引渡した事|引渡した旨|引渡しに|終了後|終了時|切替以降' | GREP_COLOR='1;4;33;41' grep -a --color -iE ${targetstr}
#python excelgrep_by_XMLparse_for_Untenshyukei.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml
#python excelgrep_by_XMLparse_for_Untenshyukei.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml | GREP_COLOR='0;33' grep -a --color -n -A 0 -iE ${targets[0]}'|'${targets[1]} 




#grep -a もしくは grep --text を使って「ちょっとバイナリファイルっぽくても諦めんなよ」という思い


#/tmp/tmp.XaXt8aTUVu/xl/media
#C:\Users\kenichi\AppData\Local\Temp\tmp.XaXt8aTUVu\xl\media

#画像フォルダを開くとき
#start $ret\\xl\\media
#画像内の文字も検索する時
#python ocr.py $ret\\xl\\media | grep -a --color -n -A 0 -iE ${targetstr}

#read -p "Hit enter: "

  # 一時ディレクトリとファイルを削除
  rm -r ${tmpdir}
#  rm ${tmp_out}

done