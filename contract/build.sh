#!/bin/sh
# 確認事項へのご回答 — Word と PDF を両方書き出す
#   sh contract/build.sh
set -e
cd "$(dirname "$0")/.."
node contract/build_answer.js
soffice --headless --convert-to pdf --outdir contract \
        "contract/聖建様_確認事項へのご回答.docx" >/dev/null 2>&1
echo "書き出し: 聖建様_確認事項へのご回答.docx / .pdf"
