from pathlib import Path
import traceback

import fitz
import win32com.client


BASE = Path(__file__).resolve().parent
DOCX = BASE / "Guia_practica_APIs_Python.docx"
OUT = BASE / "qa_api_guide"
PDF = OUT / "Guia_practica_APIs_Python.pdf"

OUT.mkdir(exist_ok=True)

try:
    print("starting Word", flush=True)
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        print(f"opening {DOCX}", flush=True)
        document = word.Documents.Open(str(DOCX))
        print("exporting PDF", flush=True)
        document.ExportAsFixedFormat(
            OutputFileName=str(PDF),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
        document.Close(False)
    finally:
        word.Quit()

    print("rasterizing", flush=True)
    pdf = fitz.open(PDF)
    matrix = fitz.Matrix(2.0, 2.0)
    for index, page in enumerate(pdf):
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        pix.save(OUT / f"page-{index + 1}.png")

    print(f"pages={len(pdf)}", flush=True)
    print(PDF, flush=True)
except Exception:
    traceback.print_exc()
    raise
