import re
from pathlib import Path
from docxtpl import DocxTemplate


def sanitize_filename(text: str) -> str:
    text = re.sub(r'[\\/:*?"<>|]', "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def generate_docx(
    template_path: Path,
    output_root: Path,
    nama_mahasiswa: str,
    npm: str,
    context: dict,
) -> Path:
    if not template_path.exists():
        raise FileNotFoundError(f"Template tidak ditemukan: {template_path}")

    nama = sanitize_filename(nama_mahasiswa)
    npm_clean = sanitize_filename(npm)

    out_dir = Path(output_root) / f"{nama}_{npm_clean}"
    out_dir.mkdir(parents=True, exist_ok=True)

    out_file = out_dir / f"Berita Acara dan Nilai Ujian Skripsi - {nama}.docx"

    doc = DocxTemplate(str(template_path))
    doc.render(context)
    doc.save(str(out_file))
    return out_file
