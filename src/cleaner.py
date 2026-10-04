"""
Text Sanitizer Utility
Ensures clean, readable, professional text without any markdown asterisks (*).
"""
import re


def bersihkan_markdown_dan_asterik(teks: str) -> str:
    """
    Membersihkan teks secara menyeluruh dari semua simbol Markdown dan tanda asterik (*).
    Menghasilkan teks yang rapi, profesional, dan mudah dibaca tanpa ada karakter bintang.
    """
    if not teks:
        return ""

    # Hapus code block
    teks = re.sub(r'```[\s\S]*?```', '', teks)
    teks = re.sub(r'`([^`]*)`', r'\1', teks)

    # Hapus tanda pagar header (### Judul -> Judul)
    def clean_header(match):
        h = match.group(1).strip()
        return f"\n{h}\n"
    teks = re.sub(r'^#{1,6}\s*(.+)$', clean_header, teks, flags=re.MULTILINE)

    # Hapus bold/italic markdown (***teks***, **teks**, *teks*)
    teks = re.sub(r'\*{3}(.+?)\*{3}', r'\1', teks)
    teks = re.sub(r'\*{2}(.+?)\*{2}', r'\1', teks)
    teks = re.sub(r'(?<!\w)\*([^\*\n]+?)\*(?!\w)', r'\1', teks)

    # Hapus underscore bold/italic (__teks__, _teks_)
    teks = re.sub(r'_{2}(.+?)_{2}', r'\1', teks)
    teks = re.sub(r'(?<!\w)_([^_\n]+?)_(?!\w)', r'\1', teks)

    # Ganti horizontal rule (--- atau ***) dengan garis pemisah bersih
    teks = re.sub(r'^[\-\*_]{3,}\s*$', '----------------------------------------', teks, flags=re.MULTILINE)

    # Format bullet point (* item) menjadi strip (- item)
    teks = re.sub(r'^\s*[\*]\s+', '- ', teks, flags=re.MULTILINE)

    # Hapus SEMUA sisa karakter bintang (*) tanpa kecuali
    teks = teks.replace('*', '')

    # Rapikan baris kosong berlebihan
    teks = re.sub(r'\n{3,}', '\n\n', teks)

    return teks.strip()
