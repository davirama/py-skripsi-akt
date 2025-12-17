from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QDate, QUrl
from PySide6.QtGui import QFont, QDesktopServices
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit, QTextEdit,
    QPushButton, QFileDialog, QMessageBox, QGroupBox, QComboBox, QDateEdit,
    QTimeEdit, QSpinBox, QFrame, QCompleter, QSizePolicy
)

from core.paths import app_root, resource_path, pilih_template_berdasarkan_pembimbing
from core.excel_loader import load_dosen_excel, Dosen
from core.date_formatter import format_tanggal_indonesia, nama_hari_indonesia, urutan_ke_kata
from core.validator import FormData, validate_form
from core.word_generator import generate_docx


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Berita Acara & Nilai Ujian Skripsi (S1)")
        self.setMinimumWidth(920)

        self.excel_path: Path | None = None
        self.output_root: Path = app_root() / "output"

        self.dosen_by_id: dict[str, Dosen] = {}
        self.display_to_id: dict[str, str] = {}

        self._build_ui()
        self._apply_styles()
        self._load_defaults_if_any()
        self._refresh_output_label()

        # Kalau kamu MAU maksa dari sini (opsional), uncomment:
        # self.showMaximized()

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

        subtitle = QLabel("Pilih dosen dari Excel → isi form → generate .docx ke folder output/Nama_NPM/")
        subtitle.setStyleSheet("color: #666;")

        root.addWidget(title)
        root.addWidget(subtitle)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        root.addWidget(line)

        # =========================
        # Data Mahasiswa (FIX)
        # =========================
        g_mhs = QGroupBox("Data Mahasiswa")
        lay_mhs = QGridLayout(g_mhs)
        lay_mhs.setHorizontalSpacing(10)
        lay_mhs.setVerticalSpacing(8)

        # kunci proporsi kolom
        lay_mhs.setColumnStretch(0, 0)   # label kiri
        lay_mhs.setColumnStretch(1, 5)   # input kiri
        lay_mhs.setColumnStretch(2, 0)   # label kanan
        lay_mhs.setColumnStretch(3, 3)   # input kanan

        # kunci tinggi baris "Ujian ke-" biar spinbox gak ke-cut saat fullscreen
        lay_mhs.setRowMinimumHeight(2, 44)

        self.in_nama = QLineEdit()
        self.in_npm = QLineEdit()

        self.in_judul = QTextEdit()
        self.in_judul.setFixedHeight(90)

        self.in_urutan = QSpinBox()
        self.in_urutan.setRange(1, 20)
        self.in_urutan.setValue(1)
        self.in_urutan.setFixedWidth(120)
        self.in_urutan.setMinimumHeight(34)  # <- kunci tinggi supaya tidak ketiban
        self.in_urutan.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        # row 0 (SEMUA SATU BARIS, AMAN FULLSCREEN)
        lay_mhs.addWidget(QLabel("Nama Mahasiswa"), 0, 0)
        lay_mhs.addWidget(self.in_nama, 0, 1)

        lay_mhs.addWidget(QLabel("NPM"), 0, 2)
        lay_mhs.addWidget(self.in_npm, 0, 3)

        lay_mhs.addWidget(QLabel("Ujian ke-"), 0, 4, Qt.AlignRight)
        lay_mhs.addWidget(self.in_urutan, 0, 5, Qt.AlignLeft)


        # row 1
        lay_mhs.addWidget(QLabel("Judul Skripsi"), 1, 0, Qt.AlignTop)
        lay_mhs.addWidget(self.in_judul, 1, 1, 1, 5)


        # spacer untuk ngisi sisa kolom 2-3 (supaya layout rapi)
        sp = QWidget()
        sp.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        lay_mhs.addWidget(sp, 2, 2, 1, 2)

        root.addWidget(g_mhs)

        # =========================
        # Waktu Ujian (1 baris, RATA & STABIL)
        # =========================
        g_waktu = QGroupBox("Waktu Ujian")
        lay_waktu = QGridLayout(g_waktu)
        lay_waktu.setHorizontalSpacing(10)
        lay_waktu.setVerticalSpacing(6)

        # atur proporsi kolom
        lay_waktu.setColumnStretch(0, 0)  # label
        lay_waktu.setColumnStretch(1, 2)  # tanggal
        lay_waktu.setColumnStretch(2, 0)  # label
        lay_waktu.setColumnStretch(3, 3)  # hari (paling panjang)
        lay_waktu.setColumnStretch(4, 0)  # label
        lay_waktu.setColumnStretch(5, 1)  # jam mulai
        lay_waktu.setColumnStretch(6, 0)  # label
        lay_waktu.setColumnStretch(7, 1)  # jam selesai

        # widgets
        self.in_tanggal = QDateEdit()
        self.in_tanggal.setCalendarPopup(True)
        self.in_tanggal.setDate(QDate.currentDate())
        self.in_tanggal.dateChanged.connect(self._on_date_changed)
        self.in_tanggal.setMinimumHeight(34)

        self.in_hari = QLineEdit()
        self.in_hari.setReadOnly(True)
        self.in_hari.setText(nama_hari_indonesia(self.in_tanggal.date()))
        self.in_hari.setMinimumHeight(34)

        self.in_mulai = QTimeEdit()
        self.in_mulai.setDisplayFormat("HH:mm")
        self.in_mulai.setMinimumHeight(34)

        self.in_selesai = QTimeEdit()
        self.in_selesai.setDisplayFormat("HH:mm")
        self.in_selesai.setMinimumHeight(34)

        # row 0 (label)
        lay_waktu.addWidget(QLabel("Tanggal"), 0, 0)
        lay_waktu.addWidget(QLabel("Hari (auto)"), 0, 2)
        lay_waktu.addWidget(QLabel("Jam Mulai"), 0, 4)
        lay_waktu.addWidget(QLabel("Jam Selesai"), 0, 6)

        # row 1 (input)
        lay_waktu.addWidget(self.in_tanggal, 1, 0, 1, 2)
        lay_waktu.addWidget(self.in_hari, 1, 2, 1, 2)
        lay_waktu.addWidget(self.in_mulai, 1, 4, 1, 2)
        lay_waktu.addWidget(self.in_selesai, 1, 6, 1, 2)

        root.addWidget(g_waktu)

        # =========================
        # Dosen
        # =========================
        g_dosen = QGroupBox("Dosen")
        lay_dosen = QGridLayout(g_dosen)
        lay_dosen.setHorizontalSpacing(10)
        lay_dosen.setVerticalSpacing(8)

        lay_dosen.setColumnStretch(0, 0)
        lay_dosen.setColumnStretch(1, 1)
        lay_dosen.setColumnStretch(2, 0)
        lay_dosen.setColumnStretch(3, 1)

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

        # =========================
        # Folder Output
        # =========================
        g_out = QGroupBox("Folder Output")
        lay_out = QGridLayout(g_out)
        lay_out.setHorizontalSpacing(10)
        lay_out.setVerticalSpacing(8)

        lay_out.setColumnStretch(0, 0)
        lay_out.setColumnStretch(1, 1)
        lay_out.setColumnStretch(2, 0)
        lay_out.setColumnStretch(3, 0)

        self.out_path_view = QLineEdit()
        self.out_path_view.setReadOnly(True)

        self.btn_output = QPushButton("Pilih Folder Output")
        self.btn_open_output = QPushButton("Buka Folder Output")

        lay_out.addWidget(QLabel("Lokasi"), 0, 0)
        lay_out.addWidget(self.out_path_view, 0, 1, 1, 3)
        lay_out.addWidget(self.btn_output, 1, 2)
        lay_out.addWidget(self.btn_open_output, 1, 3)

        root.addWidget(g_out)

        # =========================
        # Buttons
        # =========================
        row_btn = QHBoxLayout()
        row_btn.setSpacing(10)

        self.btn_excel = QPushButton("Load Excel Dosen")
        self.btn_generate = QPushButton("Generate")
        self.btn_generate.setObjectName("primaryButton")
        self.btn_generate.setFixedWidth(140)

        row_btn.addWidget(self.btn_excel)
        row_btn.addStretch(1)
        row_btn.addWidget(self.btn_generate)

        root.addLayout(row_btn)

        self.lbl_status = QLabel("Status: siap. Template auto (1/2 pembimbing).")
        self.lbl_status.setStyleSheet("color: #444; margin-top: 4px;")
        root.addWidget(self.lbl_status)

        # signals
        self.btn_excel.clicked.connect(self.on_pick_excel)
        self.btn_output.clicked.connect(self.on_pick_output_folder)
        self.btn_open_output.clicked.connect(self.on_open_output_folder)
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

    # ---------- Output folder ----------
    def _refresh_output_label(self):
        self.output_root.mkdir(parents=True, exist_ok=True)
        self.out_path_view.setText(str(self.output_root))

    def on_pick_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Pilih Folder Output", str(self.output_root))
        if not folder:
            return
        self.output_root = Path(folder)
        self._refresh_output_label()
        self.lbl_status.setText(f"Status: folder output diubah → {self.output_root}")

    def on_open_output_folder(self):
        self._refresh_output_label()
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_root)))

    # ---------- Defaults ----------
    def _load_defaults_if_any(self):
        default_excel = resource_path("resources/dosen.xlsx")
        if default_excel.exists():
            self.excel_path = default_excel
            try:
                self._load_excel_into_ui(default_excel)
                self.lbl_status.setText("Status: default Excel terdeteksi. Template auto (1/2 pembimbing).")
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
            cb.addItem("")
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

    # ---------- Generate ----------
    def on_generate(self):
        if not self.excel_path or not self.excel_path.exists():
            QMessageBox.warning(self, "Belum siap", "Excel dosen belum diload / tidak ditemukan.")
            return

        pb1 = self._selected_dosen(self.cb_pb1.currentText().strip())
        pb2 = self._selected_dosen(self.cb_pb2.currentText().strip())
        pj1 = self._selected_dosen(self.cb_pj1.currentText().strip())
        pj2 = self._selected_dosen(self.cb_pj2.currentText().strip())

        if not pb1:
            QMessageBox.warning(self, "Input belum valid", "Pembimbing 1 wajib dipilih.")
            return

        jumlah_pembimbing = 2 if pb2 else 1
        template_path = pilih_template_berdasarkan_pembimbing(jumlah_pembimbing)

        if not template_path.exists():
            QMessageBox.critical(
                self,
                "Template tidak ditemukan",
                f"Template tidak ditemukan:\n{template_path}\n\nPastikan file ada di folder resources/."
            )
            return

        nama_mhs = self.in_nama.text().strip()
        npm = self.in_npm.text().strip()
        judul = self.in_judul.toPlainText().strip()

        urutan_angka = int(self.in_urutan.value())
        urutan_kata = urutan_ke_kata(urutan_angka)

        hari = self.in_hari.text().strip()
        tanggal_str = format_tanggal_indonesia(self.in_tanggal.date())
        jam_mulai = self.in_mulai.time().toString("HH:mm")
        jam_selesai = self.in_selesai.time().toString("HH:mm")

        fd = FormData(
            nama_mahasiswa=nama_mhs,
            npm=npm,
            judul_skripsi=judul,
            urutan=urutan_angka,
            hari=hari,
            jam_mulai=jam_mulai,
            jam_selesai=jam_selesai,
            pembimbing_1=pb1.nama,
            pembimbing_2=pb2.nama if pb2 else "",
            penguji_1=pj1.nama if pj1 else "",
            penguji_2=pj2.nama if pj2 else "",
        )

        ok, msg = validate_form(fd)
        if not ok:
            QMessageBox.warning(self, "Input belum valid", msg)
            return

        context = {
            "hari": hari,
            "tanggal_bulan_tahun": tanggal_str,
            "jam_mulai": jam_mulai,
            "jam_selesai": jam_selesai,
            "urutan": urutan_kata,

            "nama_mahasiswa": nama_mhs,
            "npm": npm,
            "judul_skripsi": judul,

            "pembimbing_1": pb1.nama,
            "pembimbing_2": pb2.nama if pb2 else "",

            "penguji_1": pj1.nama if pj1 else "",
            "penguji_2": pj2.nama if pj2 else "",

            "nipnup_penguji1": pj1.jenis_id if pj1 else "",
            "nomor_nipnup_penguji1": pj1.id if pj1 else "",
            "nipnup_penguji2": pj2.jenis_id if pj2 else "",
            "nomor_nipnup_penguji2": pj2.id if pj2 else "",
        }

        try:
            self._refresh_output_label()
            out_path = generate_docx(
                template_path=template_path,
                output_root=self.output_root,
                nama_mahasiswa=nama_mhs,
                npm=npm,
                context=context,
            )
        except Exception as e:
            QMessageBox.critical(self, "Gagal generate", str(e))
            return

        self.lbl_status.setText(f"Status: sukses (template {jumlah_pembimbing} pembimbing) → {out_path}")
        QMessageBox.information(self, "Sukses", f"Dokumen berhasil dibuat:\n{out_path}")
