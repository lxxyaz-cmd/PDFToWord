"""Build a mixed text/image-table PDF from the retained quotation scan."""
from pathlib import Path
import pymupdf

root = Path(__file__).resolve().parents[1]
target = root / 'tmp' / 'pdfs'
target.mkdir(parents=True, exist_ok=True)
with pymupdf.open(root / '报价单.pdf') as source:
    pix = source[0].get_pixmap(dpi=300, clip=pymupdf.Rect(38, 198, 564, 374))
    document = pymupdf.open()
    page = document.new_page(width=595, height=841)
    page.insert_text((32, 55), 'Native heading 12345', fontsize=14)
    page.insert_image(pymupdf.Rect(32, 90, 563, 268), stream=pix.tobytes('png'))
    page.insert_text((32, 320), 'Native footer 2026', fontsize=12)
    document.save(target / 'mixed-image-table.pdf')
    document.close()
print(target / 'mixed-image-table.pdf')
