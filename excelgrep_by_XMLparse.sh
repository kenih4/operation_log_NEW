#!/bin/bash

#Usage:
# https://qiita.com/UKIUKI_ENGINEER/items/76d1ba94c2e210bc5f5d
#
#  Terminal:	GitBash
# ./excelgrep_by_XMLparse.sh -k='dcct' SP8/*.xlsm
#
# excelgrep_by_XMLparse.sh を検索モードか集計モードか、-k=検索ワードとすると検索モードで実行
#
# xlsmの解凍は不要(Python側でzipのまま読む)。複数ファイルでもPythonは1回だけ起動する。
#
# Formatter     Shift+Alt+F
# Ctrl + Shift + P (Windows)
# Formatter install方法　go install mvdan.cc/sh/v3/cmd/shfmt@latest
# shfmtの場所　C:\Users\kenic\go\bin

echo Argument: ${@}

# 引数からgrepの操作内容を取り出す^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
for arg in "$@"; do
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

# 処理対象のExcelファイルを集める
files=("$@")
file_count=${#files[@]} # 2. 配列の要素数（ファイル数）を取得する
targets=()
#for ((i = 0; i < file_count; i++)); do # 昇順ループ
for ((i = file_count - 1; i >= 0; i--)); do # 降順ループ
	# *.xlsm, *.xls以外の引数の場合次のループへ
	case "${files[i]}" in
	*'~'*) continue ;; # 「~$2024_06_SP8.xlsm」のような一時ファイルは除く
	*.xlsm | *.xls) targets+=("${files[i]}") ;;
	*) continue ;;
	esac
done

if [ ${#targets[@]} -eq 0 ]; then
	echo "❌ 対象のExcelファイル(*.xlsm, *.xls)がありません。"
	exit 1
fi

# Pythonは1回だけ起動し、全ファイルを処理する(ファイルごとの見出し「📘 File:」もPythonが出力する)
if [ "$FLG_K" = true ]; then # ログノート検索モード
	echo "💡 ログノート検索モード（ターミナルに出力）を実行します"
	# MSYS2_ARG_CONV_EXCL: 検索ワードがパスと誤認されて変換されないようにする
	MSYS2_ARG_CONV_EXCL='--keyword=' python excelgrep_by_XMLparse.py --mode search --keyword="${targetstr}" "${targets[@]}"
else
	echo "💡 通常処理（運転集計用にログノートとicalカレンダーをHTML出力）を実行します... 色を付けるワードはVBAの「Sub ログノートをHTML出力と調整時間がログノートに記載されてるか確認_ユニット月」の中に書いてある"
	python excelgrep_by_XMLparse.py --mode summary "${targets[@]}"
fi

#画像内の文字も検索する時 (xlsmをzip展開して xl/media を ocr.py に渡す)
#python ocr.py $ret\\xl\\media | grep -a --color -n -A 0 -iE ${targetstr}
