# -*- coding: utf-8 -*-
"""聖建様 見積書ブック 仕上げスクリプト

  python3 quote/finalize.py [角印画像のパス]

・印刷設定（A4・1ページ収め）を各シートに適用
・角印画像を指定すると 01_見積書 の社名右に配置し quote/seal.png として保存
・すでに quote/seal.png があれば引数なしでもそれを使う
"""
import os, shutil, sys
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.units import cm_to_EMU
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D

HERE = os.path.dirname(os.path.abspath(__file__))
BOOK = os.path.join(HERE, "聖建様_概算見積書・注文書.xlsx")
SEAL = os.path.join(HERE, "seal.png")

# 印刷範囲は各シートの実データの下端まで（先方提出の3枚は特に切れると致命的）
PAGE = {
    "01_見積書":       ("A1:F38", "portrait"),
    "02_増減内訳":     ("A1:E46", "portrait"),
    "03_見積明細":     ("A1:F44", "landscape"),
    "04_教材科目":     ("A1:F24", "landscape"),
    "05_月額・別途費用": ("A1:D30", "landscape"),
    "06_注文書":       ("A1:F35", "portrait"),
    "07_前提条件":     ("A1:C26", "portrait"),
}

# 角印の配置（用紙左上からの距離と一辺の長さ）
SEAL_X_CM, SEAL_Y_CM, SEAL_CM = 17.3, 4.35, 2.3


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else None
    if src:
        if not os.path.exists(src):
            sys.exit("角印画像が見つかりません: " + src)
        if os.path.abspath(src) != os.path.abspath(SEAL):
            shutil.copy(src, SEAL)

    wb = openpyxl.load_workbook(BOOK)

    for name, (area, orient) in PAGE.items():
        ws = wb[name]
        ws.print_area = area
        ws.page_setup.orientation = orient
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top = ws.page_margins.bottom = 0.4
        ws.page_margins.header = ws.page_margins.footer = 0.2
        ws.print_options.horizontalCentered = True

    ws = wb["01_見積書"]
    if os.path.exists(SEAL):
        img = XLImage(SEAL)
        img.width = img.height = None
        side = cm_to_EMU(SEAL_CM)
        img.anchor = AbsoluteAnchor(
            pos=XDRPoint2D(cm_to_EMU(SEAL_X_CM), cm_to_EMU(SEAL_Y_CM)),
            ext=XDRPositiveSize2D(side, side),
        )
        ws.add_image(img)
        print("角印を配置しました:", SEAL)
    else:
        print("角印なし（quote/seal.png を置いて再実行すると押印されます）")

    wb.save(BOOK)
    print("保存:", BOOK)


if __name__ == "__main__":
    main()
