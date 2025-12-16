from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit, QTextEdit,
    QPushButton, QFileDialog, QMessageBox, QGroupBox, QComboBox, QDateEdit,
    QTimeEdit, QSpinBox, QFrame, QCompleter
)


from core.paths import resource_path, app_root
from core.excel_loader import load_dosen_excel, Dosen
from core.date_formatter import format_tanggal_indonesia, nama_hari_indonesia
from core.validator import FormData, validate_form
from core.word_generator import generate_docx
from core.date_formatter import urutan_ke_kata


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Berita Acara & Nilai Ujian Skripsi (S1)")
        self.setMinimumWidth(860)

        self.template_path: Path | None = None
        self.excel_path: Path | None = None
        self.output_root: Path = app_root() / "output"

        self.dosen_by_id: dict[str, Dosen] = {}
        self.display_to_id: dict[str, str] = {}

        self._build_ui()
        self._apply_styles()
        self._load_defaults_if_any()

    # ---------- UI ----------
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        title = QLabel("Generate Berita Acara dan Nilai Ujian Skripsi (S1)")
        f = QFont()
        f.setPointSize(14)
        f.setBold(True)
        title.setFont(f)

        subtitle = QLabel("Pilih dosen dari Excel → isi form → generate .docx ke output/Nama_NPM/")
        subtitle.setStyleSheet("color: #666;")

        root.addWidget(title)
        root.addWidget(subtitle)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        root.addWidget(line)

        # --- Data Mahasiswa ---
        g_mhs = QGroupBox("Data Mahasiswa")
        lay_mhs = QGridLayout(g_mhs)
        lay_mhs.setHorizontalSpacing(10)
        lay_mhs.setVerticalSpacing(8)

        self.in_nama = QLineEdit()
        self.in_npm = QLineEdit()
        self.in_judul = QTextEdit()
        self.in_judul.setFixedHeight(90)
        self.in_urutan = QSpinBox()
        self.in_urutan.setRange(1, 20)
        self.in_urutan.setValue(1)

        lay_mhs.addWidget(QLabel("Nama Mahasiswa"), 0, 0)
        lay_mhs.addWidget(self.in_nama, 0, 1)
        lay_mhs.addWidget(QLabel("NPM"), 0, 2)
        lay_mhs.addWidget(self.in_npm, 0, 3)

        lay_mhs.addWidget(QLabel("Judul Skripsi"), 1, 0, Qt.AlignTop)
        lay_mhs.addWidget(self.in_judul, 1, 1, 1, 3)

        lay_mhs.addWidget(QLabel("Ujian ke-"), 2, 0)
        lay_mhs.addWidget(self.in_urutan, 2, 1)

        root.addWidget(g_mhs)

        # --- Waktu ---
        g_waktu = QGroupBox("Waktu Ujian")
        lay_waktu = QGridLayout(g_waktu)
        lay_waktu.setHorizontalSpacing(10)
        lay_waktu.setVerticalSpacing(8)

        self.in_tanggal = QDateEdit()
        self.in_tanggal.setCalendarPopup(True)
        self.in_tanggal.setDate(QDate.currentDate())
        self.in_tanggal.dateChanged.connect(self._on_date_changed)

        self.in_hari = QLineEdit()
        self.in_hari.setReadOnly(True)
        self.in_hari.setText(nama_hari_indonesia(self.in_tanggal.date()))

        self.in_mulai = QTimeEdit()
        self.in_mulai.setDisplayFormat("HH:mm")
        self.in_selesai = QTimeEdit()
        self.in_selesai.setDisplayFormat("HH:mm")

        lay_waktu.addWidget(QLabel("Tanggal"), 0, 0)
        lay_waktu.addWidget(self.in_tanggal, 0, 1)
        lay_waktu.addWidget(QLabel("Hari (auto)"), 0, 2)
        lay_waktu.addWidget(self.in_hari, 0, 3)

        lay_waktu.addWidget(QLabel("Jam Mulai"), 1, 0)
        lay_waktu.addWidget(self.in_mulai, 1, 1)
        lay_waktu.addWidget(QLabel("Jam Selesai"), 1, 2)
        lay_waktu.addWidget(self.in_selesai, 1, 3)

        root.addWidget(g_waktu)

        # --- Dosen ---
        g_dosen = QGroupBox("Dosen")
        lay_dosen = QGridLayout(g_dosen)
        lay_dosen.setHorizontalSpacing(10)
        lay_dosen.setVerticalSpacing(8)

        self.cb_pb1 = self._make_searchable_combo("Pilih Pembimbing 1 (wajib)")
        self.cb_pb2 = self._make_searchable_combo("Pilih Pembimbing 2 (opsional)")
        self.cb_pj1 = self._make_searchable_combo("Pilih Penguji 1 / Ketua (wajib)")
        self.cb_pj2 = self._make_searchable_combo("Pilih Penguji 2 / Sekretaris (wajib)")

        lay_dosen.addWidget(QLabel("Pembimbing 1"), 0, 0)
        lay_dosen.addWidget(self.cb_pb1, 0, 1)
        lay_dosen.addWidget(QLabel("Pembimbing 2"), 0, 2)
        lay_dosen.addWidget(self.cb_pb2, 0, 3)

        lay_dosen.addWidget(QLabel("Penguji 1 (Ketua)"), 1, 0)
        lay_dosen.addWidget(self.cb_pj1, 1, 1)
        lay_dosen.addWidget(QLabel("Penguji 2 (Sekretaris)"), 1, 2)
        lay_dosen.addWidget(self.cb_pj2, 1, 3)

        root.addWidget(g_dosen)

        # --- Buttons row ---
        row = QHBoxLayout()
        row.setSpacing(10)

        self.btn_excel = QPushButton("Load Excel Dosen")
        self.btn_template = QPushButton("Pilih Template .docx")
        self.btn_generate = QPushButton("Generate")
        self.btn_generate.setObjectName("primaryButton")

        row.addWidget(self.btn_excel)
        row.addWidget(self.btn_template)
        row.addStretch(1)
        row.addWidget(self.btn_generate)

        root.addLayout(row)

        self.lbl_status = QLabel("Status: siap. Load Excel & template (atau pakai default di resources/).")
        self.lbl_status.setStyleSheet("color: #444; margin-top: 4px;")
        root.addWidget(self.lbl_status)

        self.btn_excel.clicked.connect(self.on_pick_excel)
        self.btn_template.clicked.connect(self.on_pick_template)
        self.btn_generate.clicked.connect(self.on_generate)

    def _apply_styles(self):
        self.setStyleSheet("""
            QWidget { font-size: 12px; }
            QGroupBox {
                font-weight: 600;
                border: 1px solid #ddd;
                border-radius: 10px;
                margin-top: 10px;
                padding: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px 0 6px;
            }
            QLineEdit, QTextEdit, QComboBox, QDateEdit, QTimeEdit, QSpinBox {
                border: 1px solid #d6d6d6;
                border-radius: 8px;
                padding: 6px 8px;
                background: #fff;
            }
            QTextEdit { padding: 8px; }
            QPushButton {
                border: 1px solid #d0d0d0;
                border-radius: 10px;
                padding: 8px 12px;
                background: #fafafa;
            }
            QPushButton:hover { background: #f2f2f2; }
            QPushButton#primaryButton {
                background: #111;
                color: white;
                border: 1px solid #111;
                font-weight: 700;
            }
            QPushButton#primaryButton:hover { background: #000; }
        """)

    def _make_searchable_combo(self, placeholder: str) -> QComboBox:
        cb = QComboBox()
        cb.setEditable(True)
        cb.setInsertPolicy(QComboBox.NoInsert)
        cb.setPlaceholderText(placeholder)

        comp = cb.completer()
        comp.setCompletionMode(QCompleter.PopupCompletion)
        comp.setFilterMode(Qt.MatchContains)
        return cb



    def _on_date_changed(self, qdate: QDate):
        self.in_hari.setText(nama_hari_indonesia(qdate))

    # ---------- Defaults ----------
    def _load_defaults_if_any(self):
        # Default template/excel di resources/
        default_template = resource_path("resources/template_berita_acara_dan_nilai.docx")
        default_excel = resource_path("resources/dosen.xlsx")

        if default_template.exists():
            self.template_path = default_template
        if default_excel.exists():
            self.excel_path = default_excel
            try:
                self._load_excel_into_ui(default_excel)
                self.lbl_status.setText("Status: default Excel & template terdeteksi. Tinggal isi form dan Generate.")
            except Exception as e:
                self.lbl_status.setText(f"Status: default Excel ditemukan, tapi gagal load: {e}")

    # ---------- Excel ----------
    def on_pick_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Pilih Excel Dosen", str(app_root()), "Excel (*.xlsx)")
        if not path:
            return
        self.excel_path = Path(path)
        try:
            self._load_excel_into_ui(self.excel_path)
            self.lbl_status.setText(f"Status: Excel dosen loaded: {self.excel_path.name}")
        except Exception as e:
            QMessageBox.critical(self, "Gagal load Excel", str(e))

    def _load_excel_into_ui(self, path: Path):
        self.dosen_by_id, self.display_to_id = load_dosen_excel(path)
        self._refill_combos()

    def _refill_combos(self):
        items = sorted(self.display_to_id.keys())

        def refill(cb: QComboBox, keep_text: str | None = None):
            cb.blockSignals(True)
            cb.clear()
            cb.addItem("")  # allow empty for optional field
            cb.addItems(items)
            if keep_text:
                idx = cb.findText(keep_text)
                if idx >= 0:
                    cb.setCurrentIndex(idx)
                else:
                    cb.setEditText(keep_text)
            cb.blockSignals(False)

        refill(self.cb_pb1, self.cb_pb1.currentText())
        refill(self.cb_pb2, self.cb_pb2.currentText())
        refill(self.cb_pj1, self.cb_pj1.currentText())
        refill(self.cb_pj2, self.cb_pj2.currentText())

    def _selected_dosen(self, display_text: str) -> Dosen | None:
        if not display_text:
            return None
        dosen_id = self.display_to_id.get(display_text)
        if not dosen_id:
            return None
        return self.dosen_by_id.get(dosen_id)

    # ---------- Template ----------
    def on_pick_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Pilih Template Word", str(app_root()), "Word (*.docx)")
        if not path:
            return
        self.template_path = Path(path)
        self.lbl_status.setText(f"Status: template dipilih: {self.template_path.name}")

    # ---------- Generate ----------
    def on_generate(self):
        # hard checks
        if not self.template_path or not self.template_path.exists():
            QMessageBox.warning(self, "Belum siap", "Template .docx belum dipilih / tidak ditemukan.")
            return
        if not self.excel_path or not self.excel_path.exists():
            QMessageBox.warning(self, "Belum siap", "Excel dosen belum diload / tidak ditemukan.")
            return

        # gather
        nama_mhs = self.in_nama.text().strip()
        npm = self.in_npm.text().strip()
        judul = self.in_judul.toPlainText().strip()

        urutan_angka = int(self.in_urutan.value())
        urutan_kata = urutan_ke_kata(urutan_angka)

        hari = self.in_hari.text().strip()

        tanggal_str = format_tanggal_indonesia(self.in_tanggal.date())
        jam_mulai = self.in_mulai.time().toString("HH:mm")
        jam_selesai = self.in_selesai.time().toString("HH:mm")

        pb1 = self._selected_dosen(self.cb_pb1.currentText().strip())
        pb2 = self._selected_dosen(self.cb_pb2.currentText().strip())
        pj1 = self._selected_dosen(self.cb_pj1.currentText().strip())
        pj2 = self._selected_dosen(self.cb_pj2.currentText().strip())

        fd = FormData(
            nama_mahasiswa=nama_mhs,
            npm=npm,
            judul_skripsi=judul,
            urutan=urutan_angka,  # validasi tetap pakai angka (aman)
            hari=hari,
            jam_mulai=jam_mulai,
            jam_selesai=jam_selesai,
            pembimbing_1=pb1.nama if pb1 else "",
            pembimbing_2=pb2.nama if pb2 else "",
            penguji_1=pj1.nama if pj1 else "",
            penguji_2=pj2.nama if pj2 else "",
        )
        ok, msg = validate_form(fd)
        if not ok:
            QMessageBox.warning(self, "Input belum valid", msg)
            return

        # context sesuai placeholder template Word kamu
        context = {
            "hari": hari,
            "tanggal_bulan_tahun": tanggal_str,
            "jam_mulai": jam_mulai,
            "jam_selesai": jam_selesai,
            "urutan": urutan_kata,  # <- ini yang masuk ke Word (pertama, kedua, dst)

            "nama_mahasiswa": nama_mhs,
            "npm": npm,
            "judul_skripsi": judul,

            "pembimbing_1": pb1.nama if pb1 else "",
            "pembimbing_2": pb2.nama if pb2 else "",

            "penguji_1": pj1.nama if pj1 else "",
            "penguji_2": pj2.nama if pj2 else "",

            # penguji 1
            "nipnup_penguji1": pj1.jenis_id if pj1 else "",
            "nomor_nipnup_penguji1": pj1.id if pj1 else "",
            # penguji 2
            "nipnup_penguji2": pj2.jenis_id if pj2 else "",
            "nomor_nipnup_penguji2": pj2.id if pj2 else "",
        }

        try:
            out_path = generate_docx(
                template_path=self.template_path,
                output_root=self.output_root,
                nama_mahasiswa=nama_mhs,
                npm=npm,
                context=context,
            )
        except Exception as e:
            QMessageBox.critical(self, "Gagal generate", str(e))
            return

        self.lbl_status.setText(f"Status: sukses → {out_path}")
        QMessageBox.information(self, "Sukses", f"Dokumen berhasil dibuat:\n{out_path}")
