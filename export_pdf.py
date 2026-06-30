"""
Export de fichiers .xlsx en .pdf "prets a imprimer" via Microsoft Excel
(COM). Necessite Excel installe localement (Windows uniquement).
"""
from pathlib import Path

import win32com.client as win32

XL_TYPE_PDF = 0


def convert_to_pdf(xlsx_paths):
    """Convertit chaque fichier xlsx de xlsx_paths en pdf (meme dossier, meme
    nom). Reutilise une seule instance Excel pour tout le lot. Renvoie la
    liste des chemins .pdf generes, dans le meme ordre."""
    xlsx_paths = [Path(p).resolve() for p in xlsx_paths]
    pdf_paths = []

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        for xlsx_path in xlsx_paths:
            pdf_path = xlsx_path.with_suffix(".pdf")
            wb = excel.Workbooks.Open(str(xlsx_path))
            try:
                wb.ExportAsFixedFormat(XL_TYPE_PDF, str(pdf_path))
            finally:
                wb.Close(SaveChanges=False)
            pdf_paths.append(pdf_path)
    finally:
        excel.Quit()

    return pdf_paths


if __name__ == "__main__":
    import sys
    files = [Path(p) for p in sys.argv[1:]]
    if not files:
        print("Usage: python export_pdf.py fichier1.xlsx [fichier2.xlsx ...]")
        raise SystemExit(1)
    results = convert_to_pdf(files)
    for p in results:
        print("PDF genere:", p)
