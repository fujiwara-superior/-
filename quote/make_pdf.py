# -*- coding: utf-8 -*-
"""聖建様 見積書ブック → PDF 書き出し

  python3 quote/make_pdf.py

LibreOffice でブック全体を1シート1ページのPDFにし、
提出用3点（見積書・前提条件・注文書）と社内用（全7シート）に切り分けます。
"""
import os, shutil, subprocess, tempfile
import pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
BOOK = os.path.join(HERE, "聖建様_概算見積書・注文書.xlsx")

# ブック内のシート順とページ番号の対応
P_QUOTE, P_TERMS, P_ORDER = 0, 6, 5


def main():
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", tmp, BOOK], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    src = os.path.join(tmp, os.path.splitext(os.path.basename(BOOK))[0] + ".pdf")
    d = pymupdf.open(src)
    assert d.page_count == 7, "シートが7ページになっていません: %d" % d.page_count

    def out(pages, name):
        n = pymupdf.open()
        for p in pages:
            n.insert_pdf(d, from_page=p, to_page=p)
        path = os.path.join(HERE, name)
        n.save(path)
        n.close()
        print("書き出し:", name)

    out([P_QUOTE], "聖建様_概算見積書.pdf")
    out([P_ORDER], "聖建様_注文書.pdf")
    out([P_TERMS], "聖建様_前提条件.pdf")
    out([P_QUOTE, P_TERMS, P_ORDER], "聖建様_提出用一式（見積書・前提条件・注文書）.pdf")
    d.close()
    shutil.copy(src, os.path.join(HERE, "聖建様_社内用_全シート.pdf"))
    print("書き出し: 聖建様_社内用_全シート.pdf")
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
