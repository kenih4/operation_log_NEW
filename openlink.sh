#!/bin/bash

# リンクのリストを定義
declare -A links
links=(
  ["Google Drive"]="https://drive.google.com"
  ["Notion Documents"]="https://notion.so"
  # ここに他のリンクを追加
)

options=("${!links[@]}")

# ユーザーに選択させる
PS3='開きたいページを選択してください: '
select opt in "${options[@]}" "Quit"; do
  case $opt in
    "Quit")
      echo "終了します。"
      break
      ;;
    *)
      url="${links[$opt]}"
      if [[ -n $url ]]; then
        xdg-open "$url"
        break
      else
        echo "無効な選択肢です。"
      fi
      ;;
  esac
done
