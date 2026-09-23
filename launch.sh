#!/usr/bin/env bash

# systemd サービスを開始（既に起動中なら即座に復帰）
systemctl --user start csv-plus-web

URL="http://localhost:8765/"

# ブラウザを起動
if command -v garcon-url-handler >/dev/null 2>&1; then
    garcon-url-handler "$URL"
elif command -v xdg-open >/dev/null 2>&1; then
    xdg-open "$URL"
fi
