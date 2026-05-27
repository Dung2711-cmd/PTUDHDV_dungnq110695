from pathlib import Path
from docx import Document


ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "BC.docx"


def delete_paragraph(paragraph):
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def main():
    doc = Document(DOCX)
    previous_caption = None
    for paragraph in list(doc.paragraphs):
        text = paragraph.text.strip()
        if text.startswith("Hình ") or text.startswith("Bảng "):
            if text == previous_caption and not paragraph._p.xpath(".//w:drawing"):
                delete_paragraph(paragraph)
                continue
            previous_caption = text
        elif text:
            previous_caption = None
    doc.save(DOCX)
    print(DOCX)


if __name__ == "__main__":
    main()
