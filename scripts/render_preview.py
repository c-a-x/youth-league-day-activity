# -*- coding: utf-8 -*-
"""把 DOCX 渲染为逐页 PNG 预览，用于交付前的版式检查。

流程：Word COM（只读打开）→ PDF → PyMuPDF 逐页导出 PNG。
需要 Windows + Microsoft Word + PyMuPDF（pip install pymupdf）。

用法：
  python render_preview.py 文件1.docx 文件2.docx [--outdir 输出目录] [--dpi 110]

输出：<outdir>/<文件名>_pageNN.png（同时保留同名 PDF），并打印 JSON 清单。
渲染失败时以非零码退出并说明原因，调用方应如实告知用户，不得声称已检查。
"""
import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="DOCX 渲染预览")
    parser.add_argument("docx", nargs="+", help="待渲染的 DOCX 路径")
    parser.add_argument("--outdir", default=None, help="PNG 输出目录，默认为首个文件旁的 _preview/")
    parser.add_argument("--dpi", type=int, default=110)
    args = parser.parse_args()

    try:
        import fitz
    except ImportError:
        print("缺少 PyMuPDF，请先 pip install pymupdf", file=sys.stderr)
        return 2

    outdir = Path(args.outdir) if args.outdir else Path(args.docx[0]).parent / "_preview"
    outdir.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix="docx2pdf_"))

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("缺少 pywin32，无法使用 Word COM 渲染", file=sys.stderr)
        return 2

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    results = {}
    try:
        for src in map(Path, args.docx):
            pdf_tmp = tmp / (src.stem + ".pdf")
            doc = word.Documents.Open(str(src.resolve()), ReadOnly=True)
            try:
                doc.SaveAs2(str(pdf_tmp.resolve()), FileFormat=17)  # wdFormatPDF
            finally:
                doc.Close(False)
            pdf_out = outdir / (src.stem + ".pdf")
            shutil.copy(pdf_tmp, pdf_out)
            pages = []
            with fitz.open(pdf_out) as pdf_doc:
                for i, page in enumerate(pdf_doc, start=1):
                    png = outdir / f"{src.stem}_page{i:02d}.png"
                    page.get_pixmap(dpi=args.dpi).save(str(png))
                    pages.append(str(png))
            results[src.name] = pages
    except Exception as exc:  # Word 缺失、文件损坏、受保护视图等
        print(f"渲染失败：{exc}", file=sys.stderr)
        return 3
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
