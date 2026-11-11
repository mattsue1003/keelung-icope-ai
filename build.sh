#!/bin/zsh
# deck.html（artifact 原始檔）→ index.html（GitHub Pages 用，加上 doctype 與 head）
cd "$(dirname "$0")"
{ echo '<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'; echo '<style>[hidden]{display:none!important}img{max-width:100%}</style>'; cat deck.html; echo '</html>'; } > index.html
echo built index.html
