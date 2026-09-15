#!/bin/sh
# 契約まわりの書面 — Word と PDF を書き出す
#   sh contract/build.sh
set -e
cd "$(dirname "$0")/.."
node contract/build_answer.js
node contract/build_terms.js
soffice --headless --convert-to pdf --outdir contract \
        "contract/聖建様_確認事項へのご回答.docx" \
        "contract/聖建様_ご契約条件に関する申し入れ.docx" >/dev/null 2>&1
echo "書き出し: 確認事項へのご回答 / ご契約条件に関する申し入れ（各 docx・pdf）"
