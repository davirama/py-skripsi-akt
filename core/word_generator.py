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
    output_filename: str | None = None,
) -> Path:
    """
    Render template docx dengan context.

    - Jika output_filename diisi:
      pakai nama file tersebut (custom, tanpa diubah strukturnya)
    - Jika None:
      pakai default:
      "Berita Acara dan Nilai Ujian Skripsi_Nama_NPM.docx"

    Output disimpan ke:
    output_root / Nama_NPM / <filename>
    """
    if not template_path.exists():
        raise FileNotFoundError(f"Template tidak ditemukan: {template_path}")

    nama = sanitize_filename(nama_mahasiswa)
    npm_clean = sanitize_filename(npm)

    out_dir = Path(output_root) / f"{nama}_{npm_clean}"
    out_dir.mkdir(parents=True, exist_ok=True)

    if output_filename:
        filename = sanitize_filename(output_filename)
    else:
        filename = f"Berita Acara dan Nilai Ujian Skripsi_{nama}_{npm_clean}"

    out_file = out_dir / f"{filename}.docx"

    doc = DocxTemplate(str(template_path))
    doc.render(context)
    doc.save(str(out_file))
    return out_file
