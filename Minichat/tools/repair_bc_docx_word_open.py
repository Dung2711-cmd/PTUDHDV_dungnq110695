from pathlib import Path
from shutil import copy2
from zipfile import ZipFile, ZIP_DEFLATED
import tempfile

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "BC.docx"
BACKUP = ROOT / "BC_backup_before_word_repair.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def remove_element(element):
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def make_text_run(text):
    r = etree.Element(f"{{{W_NS}}}r")
    t = etree.SubElement(r, f"{{{W_NS}}}t")
    t.text = text
    if text.startswith(" ") or text.endswith(" ") or "\t" in text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def paragraph_text(paragraph):
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)).strip()


def rewrite_list_entry(paragraph, label, page="1"):
    for child in list(paragraph):
        paragraph.remove(child)
    paragraph.append(make_text_run(f"{label}\t{page}"))


def repair_document_xml(xml_bytes):
    root = etree.fromstring(xml_bytes)

    # Word fields must be inside runs. Earlier edits inserted some field nodes
    # directly under paragraphs, which can make Microsoft Word refuse the file.
    for bad in root.xpath(".//w:p/w:fldChar | .//w:p/w:instrText", namespaces=NS):
        remove_element(bad)

    # Remove remaining field-code control nodes to make the document static and robust.
    for bad in root.xpath(".//w:fldChar | .//w:instrText", namespaces=NS):
        remove_element(bad)

    figure_labels = [
        "Hình 1. Kiến trúc client-server của MiniChat",
        "Hình 2. Mô hình giao tiếp client - server",
        "Hình 3. Luồng xử lý gửi và nhận tin nhắn trong MiniChat",
        "Hình 4. Mô phỏng giao diện sau khi kết nối và gửi tin nhắn",
    ]
    table_labels = [
        "Bảng 1. Các thành phần chính của ứng dụng MiniChat",
        "Bảng 2. Kết quả kiểm thử các chức năng chính",
    ]

    for p in root.xpath(".//w:p", namespaces=NS):
        text = paragraph_text(p)
        for label in figure_labels + table_labels:
            if text.startswith(label):
                rewrite_list_entry(p, label)
                break

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def repair_settings_xml(xml_bytes):
    root = etree.fromstring(xml_bytes)
    for update in root.xpath(".//w:updateFields", namespaces=NS):
        remove_element(update)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def main():
    if not DOCX.exists():
        raise FileNotFoundError(DOCX)
    if not BACKUP.exists():
        copy2(DOCX, BACKUP)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_docx = Path(tmp) / "BC_repaired.docx"
        with ZipFile(DOCX, "r") as zin, ZipFile(tmp_docx, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = repair_document_xml(data)
                elif item.filename == "word/settings.xml":
                    data = repair_settings_xml(data)
                zout.writestr(item, data)
        copy2(tmp_docx, DOCX)

    print(f"Repaired: {DOCX}")
    print(f"Backup: {BACKUP}")


if __name__ == "__main__":
    main()
