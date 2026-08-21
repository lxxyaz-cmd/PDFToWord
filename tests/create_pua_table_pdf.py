from pathlib import Path

from fontTools.ttLib import TTFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont as ReportLabTTFont
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tmp" / "pdfs" / "pua-bank-table.pdf"
FONT = ROOT / "tmp" / "fonts" / "pua-arial.ttf"


def build_pua_font() -> None:
    FONT.parent.mkdir(parents=True, exist_ok=True)
    font = TTFont(r"C:\Windows\Fonts\arial.ttf")
    for table in font["cmap"].tables:
        if not table.isUnicode():
            continue
        for code in range(0x20, 0x7F):
            glyph = table.cmap.get(code)
            if glyph:
                table.cmap[0xF000 + code] = glyph
    font.save(FONT)


def pua(text: str) -> str:
    return "".join(chr(0xF000 + ord(char)) if 0x20 <= ord(char) <= 0x7E else char for char in text)


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    build_pua_font()
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    pdfmetrics.registerFont(ReportLabTTFont("PUAArial", str(FONT)))

    page_w, page_h = landscape(A4)
    c = canvas.Canvas(str(OUT), pagesize=(page_w, page_h))
    left = 34
    top = page_h - 34
    widths = [28, 108, 126, 126, 126, 126, 92, 108]
    scale = (page_w - 68) / sum(widths)
    widths = [value * scale for value in widths]

    c.setFont("STSong-Light", 15)
    c.setFillColor(colors.HexColor("#078BCC"))
    c.drawString(left, top, "个人银行服务收费一览表")

    headers = ["序号", "收费项目名称/内容描述", "收费标准/个人理财客户", "收费标准/优选理财客户",
               "收费标准/优先理财客户", "收费标准/优先私人理财客户", "收费依据", "备注"]
    y = top - 18
    header_h = 42
    x = left
    c.setStrokeColor(colors.HexColor("#5A6670"))
    for width, label in zip(widths, headers):
        c.setFillColor(colors.HexColor("#DDE8F2"))
        c.rect(x, y - header_h, width, header_h, fill=1, stroke=1)
        c.setFillColor(colors.black)
        c.setFont("STSong-Light", 6.8)
        parts = label.split("/")
        for idx, part in enumerate(parts):
            c.drawString(x + 3, y - 13 - idx * 10, part)
        x += width

    y -= header_h + 35
    row_h = 82
    values = ["1", "账户管理费", "免费", "每月人民币100元", "每月人民币500元", "每月人民币1500元",
              "市场调节价（固定金额定价），根据监管要求执行。", "本行提供个人理财服务，具体收费标准请参考相关文件。"]
    x = left
    for index, (width, value) in enumerate(zip(widths, values)):
        c.setFillColor(colors.HexColor("#EEF4F8"))
        c.rect(x, y - row_h, width, row_h, fill=1, stroke=1)
        c.setFillColor(colors.black)
        if index in (0,):
            c.setFont("PUAArial", 7)
            c.drawCentredString(x + width / 2, y - 43, pua(value))
        elif index in (3, 4, 5):
            c.setFont("STSong-Light", 6.2)
            c.drawString(x + 4, y - 43, "每月人民币")
            prefix = pdfmetrics.stringWidth("每月人民币", "STSong-Light", 6.2)
            c.setFont("PUAArial", 6.2)
            digits = value.replace("每月人民币", "").replace("元", "")
            c.drawString(x + 4 + prefix, y - 43, pua(digits))
            digit_w = pdfmetrics.stringWidth(pua(digits), "PUAArial", 6.2)
            c.setFont("STSong-Light", 6.2)
            c.drawString(x + 4 + prefix + digit_w, y - 43, "元")
        else:
            c.setFont("STSong-Light", 5.8 if index >= 6 else 6.5)
            line_width = max(4, int(width / 7))
            for line_no in range(0, len(value), line_width):
                c.drawString(x + 4, y - 16 - line_no // line_width * 9, value[line_no:line_no + line_width])
        x += width
    c.save()
    print(OUT)


if __name__ == "__main__":
    main()
