#!/bin/bash

#Usage:
# https://qiita.com/UKIUKI_ENGINEER/items/76d1ba94c2e210bc5f5d
#
#  Terminal:	GitBash
# ./excelgrep_by_XMLparse.sh -k='dcct' SP8/*.xlsm
#
# excelgrep_by_XMLparse.sh を検索モードか集計モードか、-k=検索ワードとすると検索モードで実行
#
# Formatter     Shift+Alt+F
# Ctrl + Shift + P (Windows)
# Formatter install方法　go install mvdan.cc/sh/v3/cmd/shfmt@latest
# shfmtの場所　C:\Users\kenic\go\bin

echo Argument: ${@}

# 引数からgrepの操作内容を取り出す^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
for arg in ${@}; do
	echo arg: $arg
	case "$arg" in
	-k=*)
		# "-k=" という文字より左側（-u=自体）を削除して、値だけ取り出す
		targetstr="${arg#*=}"
		FLG_K=true
		echo "💡 ログノート検索モード(ターミナルに色を付けて出力)です。検索ワード: 「$targetstr」"
		;;
	*)
		# その他の引数の処理（必要なら記述）
		;;
	esac
done

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# 一つ一つのExcelファイルに対してgrepする
files=("$@")
file_count=${#files[@]} # 2. 配列の要素数（ファイル数）を取得する
#for ((i = 0; i < file_count; i++)); do # 昇順ループ
for ((i = file_count - 1; i >= 0; i--)); do # 降順ループ
	# *.xlsm以外の引数の場合次のループへ
	if [ $(echo "${files[i]}" | grep -vE '\.xlsm$|\.xls$') ]; then
		continue
	fi

	#メモ：　grep 指定した文字列を含まない行を抽出するためにはgrepの-vオプションを用います。
	#メモ：　ハット（^）は「～で始まる」、ドル記号（$）は「～で終わる」を意味します

	#  「~$2024_06_SP8.xlsm」のような一時ファイルは除く
	if [ $(echo "${files[i]}" | grep '~') ]; then
		continue
	fi

	echo "📘 File: "${files[i]}"__________________________________________________________________________"

	#  read -p "Hit enter: "

	# zip展開用一時ディレクトリ作成
	tmpdir=$(mktemp -d)
	# 一時ファイル作成
	#  tmp_out=$(mktemp)

	#echo tmpdir = ${tmpdir}
	#  echo tmp_out = ${tmp_out}

	#read -p "Hit enter: "

	# エクセルファイルを一時ディレクトリに解凍する
	# 標準出力とエラー出力は鬱陶しいので捨てる
	unzip -t ""${files[i]}"" >error.log #  2> error.log標準エラー出力のみ　　「-t」オプション：正常に展開できるかテストする
	if [ $? -ne 0 ]; then               # $? は、直前に実行したコマンドの終了ステータス
		echo "❌ ZIPファイルは異常です。"
		#cat error.log
		# 特定のエラーメッセージに基づく処理
		if grep -q "End-of-central-directory signature not found" error.log; then
			echo "❌ 指定されたファイルはZIP形式ではないか、壊れている可能性があります。"
		fi
		continue
		#  else
		#      echo "ZIPファイルは正常です。"
	fi
	unzip "${files[i]}" -d ${tmpdir} 1>/dev/null 2>&1 # -d ディレクトリ	指定したディレクトリに展開する

	ret=${tmpdir/\/tmp\//C:\\Users\\kenic\\AppData\\Local\\Temp\\}
	#echo start \"$ret\\xl\\media\"
	if [ ! -e $ret\\xl\\media ]; then
		echo "❌ Directory doesn't exists!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! May be unzip fail..."
		exit
	fi
	#   /tmp/tmp.KBjrD6k7Uq/xl/worksheets/sheet1.xml
	#   /tmp/tmp.KBjrD6k7Uq/xl/sharedStrings.xml

	if [ "$FLG_K" = true ]; then
		#python excelgrep_by_XMLparse.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml
		python excelgrep_by_XMLparse.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml | GREP_COLOR='0;33' grep -a --color -n -A 0 -iE ${targetstr}
		#python excelgrep_by_XMLparse.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml > ${tmp_out}
	else
		echo "💡 通常処理（運転集計用にログノートとicalカレンダーをHTML出力）を実行します... 色を付けるワードはVBAの「Sub ログノートをHTML出力と調整時間がログノートに記載されてるか確認_ユニット月」の中に書いてある"
		#python excelgrep_by_XMLparse_for_Untenshyukei.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml | grep -v -E '引渡し前|引渡し時|引渡し希望|引渡し後|引渡しが|引渡す|引渡て|引渡した事|引渡した旨|引渡しに|終了後|終了時|切替以降' | GREP_COLOR='1;4;33;41' grep -a --color -iE ${targetstr}
		python excelgrep_by_XMLparse_for_Untenshyukei.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml
		#python excelgrep_by_XMLparse_for_Untenshyukei.py ${tmpdir}/xl/sharedStrings.xml ${tmpdir}/xl/worksheets/sheet1.xml | GREP_COLOR='0;33' grep -a --color -n -A 0 -iE ${targets[0]}'|'${targets[1]}
	fi

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
