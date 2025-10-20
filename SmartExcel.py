import sys
import shutil
import datetime
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox
import locale 
try:
    import pyperclip
    CLIP_AVAILABLE = True
except Exception:
    CLIP_AVAILABLE = False


def get_default_delimiter():
    try:

        locale.setlocale(locale.LC_ALL, '') 

        conv = locale.localeconv()
        if conv['decimal_point'] == ',':
            return ';' 
        else:
            return ',' 
    except Exception:
        return ',' 


ARG_DELIMITER = get_default_delimiter() 


# Database rumus 
RUMUS_DB = {
    # Lookup
    "VLOOKUP": {
        "kategori": "Lookup",
        "deskripsi": "Mencari nilai di kolom pertama tabel dan mengembalikan nilai dari kolom lain pada baris yang sama.",
        "inputs": [
            ("lookup_value", "Nilai/sel yang dicari (contoh: A2)", None),
            ("table_array", "Range tabel (contoh: Sheet1!A2:D100)", None),
            ("col_index", "Nomor kolom hasil (contoh: 3)", None),
            ("exact_match", "Cocokkan persis? TRUE atau FALSE (contoh: FALSE)", "FALSE"),
        ],
        "template": "=VLOOKUP({lookup_value}, {table_array}, {col_index}, {exact_match})",
        "contoh": "=VLOOKUP(A2, Data!A2:D100, 3, FALSE)"
    },
    "HLOOKUP": {
        "kategori": "Lookup",
        "deskripsi": "Mencari nilai di baris pertama tabel dan mengembalikan nilai dari baris lain pada kolom yang sama.",
        "inputs": [
            ("lookup_value", "Nilai/sel yang dicari (contoh: B1)", None),
            ("table_array", "Range tabel (contoh: Sheet1!A1:D10)", None),
            ("row_index", "Nomor baris hasil (contoh: 3)", None),
            ("exact_match", "TRUE atau FALSE (contoh: TRUE)", "TRUE"),
        ],
        "template": "=HLOOKUP({lookup_value}, {table_array}, {row_index}, {exact_match})",
        "contoh": "=HLOOKUP(B1, Data!A1:D10, 3, TRUE)"
    },
    "INDEX": {
        "kategori": "Lookup",
        "deskripsi": "Mengembalikan nilai dari range berdasarkan nomor baris & kolom.",
        "inputs": [
            ("array", "Range array (contoh: A2:C10)", None),
            ("row_num", "Nomor baris relatif (contoh: 2)", None),
            ("col_num", "Nomor kolom relatif (opsional: bisa kosong)", ""),
        ],
        "template": "=INDEX({array}, {row_num}{col_part})",
        "contoh": "=INDEX(A2:C10, 2, 3)"
    },
    "MATCH": {
        "kategori": "Lookup",
        "deskripsi": "Mengembalikan posisi relatif dari suatu nilai di dalam range.",
        "inputs": [
            ("lookup_value", "Nilai/sel yang dicari (contoh: \"Budi\" atau A2)", None),
            ("lookup_array", "Range pencarian (contoh: A2:A100)", None),
            ("match_type", "0 untuk exact match, 1/-1 untuk approximate (contoh: 0)", "0"),
        ],
        "template": "=MATCH({lookup_value}, {lookup_array}, {match_type})",
        "contoh": "=MATCH(\"Budi\", A2:A100, 0)"
    },

    # Logika
    "IF": {
        "kategori": "Logika",
        "deskripsi": "Jika kondisi TRUE maka hasil1, jika FALSE maka hasil2.",
        "inputs": [
            ("condition", "Kondisi (contoh: A1>70)", None),
            ("true_value", "Hasil jika benar (contoh: \"Lulus\")", None),
            ("false_value", "Hasil jika salah (contoh: \"Gagal\")", None),
        ],
        "template": "=IF({condition}, {true_value}, {false_value})",
        "contoh": "=IF(A1>70, \"Lulus\", \"Gagal\")"
    },
    "IFS": {
        "kategori": "Logika",
        "deskripsi": "Memeriksa beberapa kondisi berurutan; mengembalikan nilai pertama yang TRUE.",
        "inputs": [
            ("pairs", "Masukkan pasangan kondisi:hasil dipisah | (contoh: A1>90:\"A\"|A1>75:\"B\"|TRUE:\"C\")", None)
        ],
        "template": "=IFS({pairs})",
        "contoh": "=IFS(A1>90, \"A\", A1>75, \"B\", TRUE, \"C\")"
    },
    "AND": {
        "kategori": "Logika",
        "deskripsi": "Mengembalikan TRUE jika semua argumen TRUE.",
        "inputs": [
            ("conditions", "Kondisi dipisah koma (contoh: A1>50, B1>30)", None)
        ],
        "template": "=AND({conditions})",
        "contoh": "=AND(A1>50, B1>30)"
    },
    "OR": {
        "kategori": "Logika",
        "deskripsi": "Mengembalikan TRUE jika salah satu argumen TRUE.",
        "inputs": [
            ("conditions", "Kondisi dipisah koma (contoh: A1=\"X\", B1=\"Y\")", None)
        ],
        "template": "=OR({conditions})",
        "contoh": "=OR(A1=\"X\", B1=\"Y\")"
    },
    "NOT": {
        "kategori": "Logika",
        "deskripsi": "Membalikkan nilai logika (TRUE→FALSE, FALSE→TRUE).",
        "inputs": [
            ("condition", "Kondisi tunggal (contoh: A1>50)", None)
        ],
        "template": "=NOT({condition})",
        "contoh": "=NOT(A1>50)"
    },

    # Matematika & Statistik
    "SUM": {
        "kategori": "Matematika",
        "deskripsi": "Menjumlahkan angka atau range.",
        "inputs": [
            ("args", "Angka atau range (pisah koma, contoh: A2:A10, 10, B2)", None)
        ],
        "template": "=SUM({args})",
        "contoh": "=SUM(A2:A10)"
    },
    "SUMIF": {
        "kategori": "Matematika",
        "deskripsi": "Menjumlahkan sel yang memenuhi satu kriteria.",
        "inputs": [
            ("range_criteria", "Range kriteria (contoh: A2:A10)", None),
            ("criteria", "Kriteria (contoh: \"Penjualan\")", None),
            ("sum_range", "Range yang dijumlahkan (contoh: B2:B10)", None),
        ],
        "template": "=SUMIF({range_criteria}, {criteria}, {sum_range})",
        "contoh": "=SUMIF(A2:A10, \"Penjualan\", B2:B10)"
    },
    "SUMIFS": {
        "kategori": "Matematika",
        "deskripsi": "Menjumlahkan sel yang memenuhi beberapa kriteria.",
        "inputs": [
            ("sum_range", "Range yang dijumlahkan (contoh: B2:B10)", None),
            ("range1", "Range kriteria 1 (contoh: A2:A10)", None),
            ("criteria1", "Kriteria 1 (contoh: \"Penjualan\")", None),
            ("range2", "Range kriteria 2 (opsional)", ""),
            ("criteria2", "Kriteria 2 (opsional)", "")
        ],
        "template": "=SUMIFS({sum_range}, {range1}, {criteria1}{extra})",
        "contoh": "=SUMIFS(B2:B10, A2:A10, \"Penjualan\")"
    },
    "AVERAGE": {
        "kategori": "Matematika",
        "deskripsi": "Menghitung rata-rata dari sekumpulan angka.",
        "inputs": [("args", "Angka atau range (pisah koma)", None)],
        "template": "=AVERAGE({args})",
        "contoh": "=AVERAGE(A2:A10)"
    },
    "AVERAGEIF": {
        "kategori": "Matematika",
        "deskripsi": "Menghitung rata-rata berdasarkan satu kriteria.",
        "inputs": [
            ("range", "Range kriteria (contoh: A2:A10)", None),
            ("criteria", "Kriteria (contoh: \"Lulus\")", None),
            ("average_range", "Range untuk rata-rata (opsional)", "")
        ],
        "template": "=AVERAGEIF({range}, {criteria}{extra})",
        "contoh": "=AVERAGEIF(A2:A10, \"Lulus\", B2:B10)"
    },
    "COUNT": {
        "kategori": "Statistik",
        "deskripsi": "Menghitung jumlah sel yang berisi angka.",
        "inputs": [("args", "Range atau nilai (contoh: A2:A10)", None)],
        "template": "=COUNT({args})",
        "contoh": "=COUNT(A2:A10)"
    },
    "COUNTIF": {
        "kategori": "Statistik",
        "deskripsi": "Menghitung sel yang memenuhi satu kriteria.",
        "inputs": [
            ("range", "Range (contoh: A2:A10)", None),
            ("criteria", "Kriteria (contoh: \">=90\")", None)
        ],
        "template": "=COUNTIF({range}, {criteria})",
        "contoh": "=COUNTIF(A2:A10, \">=90\")"
    },
    "COUNTIFS": {
        "kategori": "Statistik",
        "deskripsi": "Menghitung sel yang memenuhi beberapa kriteria.",
        "inputs": [
            ("range1", "Range kriteria 1 (contoh: A2:A10)", None),
            ("criteria1", "Kriteria 1 (contoh: \"Penjualan\")", None),
            ("range2", "Range kriteria 2 (opsional)", ""),
            ("criteria2", "Kriteria 2 (opsional)", "")
        ],
        "template": "=COUNTIFS({range1}, {criteria1}{extra})",
        "contoh": "=COUNTIFS(A2:A10, \"Penjualan\")"
    },
    "MEDIAN": {
        "kategori": "Statistik",
        "deskripsi": "Mengembalikan nilai median.",
        "inputs": [("args", "Angka atau range", None)],
        "template": "=MEDIAN({args})",
        "contoh": "=MEDIAN(A2:A10)"
    },
    "STDEV.P": {
        "kategori": "Statistik",
        "deskripsi": "Standar deviasi populasi (Excel: STDEV.P).",
        "inputs": [("args", "Range atau angka (contoh: A2:A10)", None)],
        "template": "=STDEV.P({args})",
        "contoh": "=STDEV.P(A2:A10)"
    },

    # Teks
    "CONCAT": {
        "kategori": "Teks",
        "deskripsi": "Menggabungkan beberapa teks menjadi satu.",
        "inputs": [("args", "Teks / sel dipisah koma (contoh: A2, \" - \", B2)", None)],
        "template": "=CONCAT({args})",
        "contoh": "=CONCAT(A2, \" \", B2)"
    },
    "CONCATENATE": {
        "kategori": "Teks",
        "deskripsi": "Versi lama CONCAT; menggabungkan teks.",
        "inputs": [("args", "Teks / sel dipisah koma", None)],
        "template": "=CONCATENATE({args})",
        "contoh": "=CONCATENATE(\"Nama: \", A2)"
    },
    "TEXTJOIN": {
        "kategori": "Teks",
        "deskripsi": "Menggabungkan teks dengan delimiter dan opsi lewati kosong.",
        "inputs": [
            ("delimiter", "Delimiter (contoh: \", \" atau \"-\")", ", "),
            ("ignore_empty", "TRUE atau FALSE (contoh: TRUE)", "TRUE"),
            ("args", "Range atau teks yang digabung (contoh: A2:A10)", None)
        ],
        "template": "=TEXTJOIN({delimiter}, {ignore_empty}, {args})",
        "contoh": "=TEXTJOIN(\", \", TRUE, A2:A10)"
    },
    "LEFT": {
        "kategori": "Teks",
        "deskripsi": "Mengambil beberapa karakter dari kiri.",
        "inputs": [("text", "Teks atau sel (contoh: A2)", None), ("num_chars", "Jumlah karakter (contoh: 3)", None)],
        "template": "=LEFT({text}, {num_chars})",
        "contoh": "=LEFT(A2, 3)"
    },
    "RIGHT": {
        "kategori": "Teks",
        "deskripsi": "Mengambil beberapa karakter dari kanan.",
        "inputs": [("text", "Teks atau sel", None), ("num_chars", "Jumlah karakter", None)],
        "template": "=RIGHT({text}, {num_chars})",
        "contoh": "=RIGHT(A2, 4)"
    },
    "MID": {
        "kategori": "Teks",
        "deskripsi": "Mengambil bagian teks dari posisi tertentu.",
        "inputs": [("text", "Teks atau sel", None), ("start", "Posisi awal (1-based)", None), ("length", "Panjang substring", None)],
        "template": "=MID({text}, {start}, {length})",
        "contoh": "=MID(A2, 2, 3)"
    },
    "LEN": {
        "kategori": "Teks",
        "deskripsi": "Menghitung jumlah karakter dalam teks.",
        "inputs": [("text", "Teks atau sel", None)],
        "template": "=LEN({text})",
        "contoh": "=LEN(A2)"
    },
    "TRIM": {
        "kategori": "Teks",
        "deskripsi": "Menghapus spasi ekstra di teks (kecuali spasi tunggal antar kata).",
        "inputs": [("text", "Teks atau sel", None)],
        "template": "=TRIM({text})",
        "contoh": "=TRIM(A2)"
    },
    "UPPER": {
        "kategori": "Teks",
        "deskripsi": "Mengubah teks menjadi huruf kapital semua.",
        "inputs": [("text", "Teks atau sel", None)],
        "template": "=UPPER({text})",
        "contoh": "=UPPER(A2)"
    },
    "LOWER": {
        "kategori": "Teks",
        "deskripsi": "Mengubah teks menjadi huruf kecil semua.",
        "inputs": [("text", "Teks atau sel", None)],
        "template": "=LOWER({text})",
        "contoh": "=LOWER(A2)"
    },
    "PROPER": {
        "kategori": "Teks",
        "deskripsi": "Mengubah teks menjadi format Title Case (setiap kata kapital).",
        "inputs": [("text", "Teks atau sel", None)],
        "template": "=PROPER({text})",
        "contoh": "=PROPER(A2)"
    },

    # Tanggal & Waktu (beberapa contoh)
    "TODAY": {
        "kategori": "Tanggal",
        "deskripsi": "Menghasilkan tanggal hari ini.",
        "inputs": [],
        "template": "=TODAY()",
        "contoh": "=TODAY()"
    },
    "NOW": {
        "kategori": "Tanggal",
        "deskripsi": "Menghasilkan tanggal & waktu sekarang.",
        "inputs": [],
        "template": "=NOW()",
        "contoh": "=NOW()"
    },
    "YEAR": {
        "kategori": "Tanggal",
        "deskripsi": "Mengambil tahun dari tanggal.",
        "inputs": [("date", "Sel atau tanggal (contoh: A2)", None)],
        "template": "=YEAR({date})",
        "contoh": "=YEAR(A2)"
    },
    "MONTH": {
        "kategori": "Tanggal",
        "deskripsi": "Mengambil bulan dari tanggal.",
        "inputs": [("date", "Sel atau tanggal", None)],
        "template": "=MONTH({date})",
        "contoh": "=MONTH(A2)"
    },
    "DAY": {
        "kategori": "Tanggal",
        "deskripsi": "Mengambil hari (tanggal) dari tanggal.",
        "inputs": [("date", "Sel atau tanggal", None)],
        "template": "=DAY({date})",
        "contoh": "=DAY(A2)"
    },
    "DATEDIF": {
        "kategori": "Tanggal",
        "deskripsi": "Menghitung selisih antara dua tanggal (Excel punya fungsi tersembunyi DATEDIF).",
        "inputs": [
            ("start_date", "Tanggal mulai (contoh: A2)", None),
            ("end_date", "Tanggal akhir (contoh: B2)", None),
            ("unit", "Unit hasil (\"Y\",\"M\",\"D\",\"YM\",\"YD\",\"MD\")", "Y")
        ],
        "template": "=DATEDIF({start_date}, {end_date}, \"{unit}\")",
        "contoh": "=DATEDIF(A2, B2, \"Y\")"
    },
    "NETWORKDAYS": {
        "kategori": "Tanggal",
        "deskripsi": "Menghitung jumlah hari kerja (Mon-Fri) antara dua tanggal.",
        "inputs": [
            ("start_date", "Tanggal mulai", None),
            ("end_date", "Tanggal akhir", None),
            ("holidays", "Range hari libur (opsional)", "")
        ],
        "template": "=NETWORKDAYS({start_date}, {end_date}{extra})",
        "contoh": "=NETWORKDAYS(A2, B2, Holidays!A2:A5)"
    },

    # Keuangan (beberapa contoh)
    "PMT": {
        "kategori": "Keuangan",
        "deskripsi": "Menghitung pembayaran pinjaman periodik (pmt).",
        "inputs": [
            ("rate", "Suku bunga per periode (contoh: 0.01 untuk 1%)", None),
            ("nper", "Jumlah periode (contoh: 60)", None),
            ("pv", "Nilai sekarang / pokok (contoh: 1000000)", None),
            ("fv", "Nilai masa depan (opsional)", "0"),
            ("type", "0=akhir periode/1=awal periode (opsional)", "0"),
        ],
        "template": "=PMT({rate}, {nper}, {pv}, {fv}, {type})",
        "contoh": "=PMT(0.01, 60, 1000000)"
    },
    "FV": {
        "kategori": "Keuangan",
        "deskripsi": "Menghitung nilai masa depan investasi.",
        "inputs": [
            ("rate", "Suku bunga per periode", None),
            ("nper", "Jumlah periode", None),
            ("pmt", "Pembayaran periodik (negatif jika keluar)", None),
            ("pv", "Nilai sekarang (opsional)", "0"),
            ("type", "0 atau 1 (opsional)", "0")
        ],
        "template": "=FV({rate}, {nper}, {pmt}, {pv}, {type})",
        "contoh": "=FV(0.01, 12, -1000, 0, 0)"
    },
    "PV": {
        "kategori": "Keuangan",
        "deskripsi": "Menghitung nilai sekarang dari aliran kas.",
        "inputs": [
            ("rate", "Suku bunga per periode", None),
            ("nper", "Jumlah periode", None),
            ("pmt", "Pembayaran periodik", None),
            ("fv", "Nilai masa depan (opsional)", "0"),
            ("type", "0 atau 1 (opsional)", "0")
        ],
        "template": "=PV({rate}, {nper}, {pmt}, {fv}, {type})",
        "contoh": "=PV(0.01, 12, -1000)"
    },

    # Utility / Lainnya
    "IFERROR": {
        "kategori": "Utility",
        "deskripsi": "Mengembalikan nilai alternatif jika formula menghasilkan error.",
        "inputs": [("value", "Rumus atau nilai (contoh: VLOOKUP(...))", None), ("value_if_error", "Nilai jika error (contoh: \"-\")", " \"-\"")],
        "template": "=IFERROR({value}, {value_if_error})",
        "contoh": "=IFERROR(VLOOKUP(A2,Data!A2:C100,2,FALSE), \"Not found\")"
    },
    "TEXT": {
        "kategori": "Utility",
        "deskripsi": "Mengubah angka/tanggal menjadi teks dengan format.",
        "inputs": [("value", "Nilai atau sel (contoh: A2)", None), ("format_text", "Format teks Excel (contoh: \"dd-mm-yyyy\" atau \"#,##0.00\")", None)],
        "template": "=TEXT({value}, \"{format_text}\")",
        "contoh": "=TEXT(A2, \"dd-mm-yyyy\")"
    }
}

# Helper functions 
def build_formula(meta, collected):
    global ARG_DELIMITER
    tmpl = meta.get("template", "")
    formula = tmpl

    if "{col_part}" in formula:
        col_num = collected.get("col_num", "").strip() if collected else ""
        cp = f"{ARG_DELIMITER} {col_num}" if col_num else ""
        formula = formula.replace("{col_part}", cp)
        if "col_num" in (collected or {}):
            del collected["col_num"]

    if "{extra}" in formula:
        extras = []
        for k,v in (collected or {}).items():
            placeholder = "{" + k + "}"
            if placeholder not in tmpl and v and v.strip():

                extras.append(v.strip())
        
        extra_text = f"{ARG_DELIMITER} " + f"{ARG_DELIMITER} ".join(extras) if extras else ""
        formula = formula.replace("{extra}", extra_text)

    for key, val in (collected or {}).items():
        placeholder = "{" + key + "}"

        val_str = str(val).strip() if val is not None else ""
        formula = formula.replace(placeholder, val_str)

    def replace_commas_outside_quotes(text, new_delimiter):
        parts = []
        in_quote = False
        current_part = []
        
        for i, char in enumerate(text):
            if char == '"':
                in_quote = not in_quote
            
            if char == ',' and not in_quote:
                parts.append("".join(current_part))
                current_part = []
            else:
                current_part.append(char)
        
        parts.append("".join(current_part)) 
        

        return new_delimiter.join(parts)
        

    if ARG_DELIMITER != ',':


        formula = formula.replace(", ,", ",") 
        formula = formula.replace(",)", ")")   


        formula = replace_commas_outside_quotes(formula, ARG_DELIMITER)
        
        while f"{ARG_DELIMITER}{ARG_DELIMITER}" in formula:
            formula = formula.replace(f"{ARG_DELIMITER}{ARG_DELIMITER}", ARG_DELIMITER)
        while f"({ARG_DELIMITER}" in formula:
            formula = formula.replace(f"({ARG_DELIMITER}", "(")
        while f"{ARG_DELIMITER})" in formula:
            formula = formula.replace(f"{ARG_DELIMITER})", ")")
        while f" {ARG_DELIMITER} " in formula:
            formula = formula.replace(f" {ARG_DELIMITER} ", ARG_DELIMITER)
        formula = formula.replace(" ", "") 
    else:
        while ", ," in formula:
            formula = formula.replace(", ,", ",")
        while ",)" in formula:
            formula = formula.replace(",)", ")")
            
    if ARG_DELIMITER == ',':
        formula = formula.replace(" ", "")
        
    return formula.replace("  ", " ").strip()


def get_script_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def save_to_file(formula_text, filename="rumus_hasil.txt"):
    try:
        file_path = os.path.join(get_script_path(), filename)
        with open(file_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]  {formula_text}\n")
        return True, file_path
    except Exception as e:
        return False, str(e)
        
def set_delimiter(delimiter):
    global ARG_DELIMITER
    ARG_DELIMITER = delimiter

class ExcelRumusGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🧮 Asisten Rumus Excel - Sneijderlino")
        self.geometry("1000x720")
        self.configure(bg="#f7f9fb")
        self.resizable(True, True)

        self.selected_rumus = None
        self.form_entries = {}
        
        self.delimiter_var = tk.StringVar(value=ARG_DELIMITER) 
        
        self.create_widgets()

    def change_delimiter_and_notify(self, event=None):
        new_delimiter = self.delimiter_var.get()
        set_delimiter(new_delimiter)
        self.status.config(text=f"Pemisah diatur ke '{new_delimiter}' (Excel Regional Setting) ✅")

    def create_widgets(self):
        # Header
        header = tk.Label(self, text="🤖 Asisten Rumus Excel — By Sneijderlino",
                        font=("Segoe UI", 18, "bold"), bg="#f7f9fb", fg="#222")
        header.pack(pady=10)

        subtitle = tk.Label(self, text="Pilih rumus, isi parameter, lalu klik 'Buat Rumus' — hasil siap copy-paste ke Excel",
                            font=("Segoe UI", 11), bg="#f7f9fb", fg="#333")
        subtitle.pack()
        
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=18, pady=10)


        delimiter_frame = ttk.Frame(top_frame)
        delimiter_frame.pack(side="right", padx=6)
        
        ttk.Label(delimiter_frame, text="Pemisah Argumen:", font=("Segoe UI", 10, "bold")).pack(side="left", padx=6)
        

        self.delimiter_var = tk.StringVar(value=get_default_delimiter()) 
        delimiter_combo = ttk.Combobox(delimiter_frame, textvariable=self.delimiter_var, 
                                        values=[',', ';'], width=5, state="readonly")
        delimiter_combo.pack(side="left")
        delimiter_combo.bind("<<ComboboxSelected>>", self.change_delimiter_and_notify)
        
        # Tombol Semua Rumus
        ttk.Button(top_frame, text="📂 Semua Rumus", command=self.show_all_rumus).pack(side="right", padx=6)

        # 2. Kolom Pencarian
        ttk.Label(top_frame, text="Cari Rumus:", font=("Segoe UI", 10)).pack(side="left", padx=6)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(top_frame, textvariable=self.search_var, width=36)
        self.search_entry.pack(side="left")
        self.search_entry.bind("<KeyRelease>", self.update_listbox)
        ttk.Button(top_frame, text="🔍 Cari", command=self.update_listbox).pack(side="left", padx=6)
        
        # --- Frame Kiri (Listbox) ---
        left_frame = ttk.Frame(self)
        left_frame.pack(side="left", fill="y", padx=12, pady=8)

        ttk.Label(left_frame, text="📘 Daftar Rumus:", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.listbox = tk.Listbox(left_frame, height=32, width=36, font=("Consolas", 10),
                                selectbackground="#0078d4", selectforeground="white")
        self.listbox.pack(fill="y", pady=6)
        self.listbox.bind("<<ListboxSelect>>", self.on_rumus_select)

        self.show_all_rumus()

        # --- Frame Kanan (Detail, Form, Output) ---
        right_frame = ttk.Frame(self)
        right_frame.pack(side="right", expand=True, fill="both", padx=12, pady=8)

        # Detail Box
        self.detail_text = tk.Text(right_frame, height=7, wrap="word", font=("Segoe UI", 10), 
                                bg="#eef3f7", fg="#333", padx=5, pady=5)
        self.detail_text.pack(fill="x", pady=6)
        self.detail_text.config(state="disabled")

        # Form Parameter (Scrollable)
        ttk.Label(right_frame, text="Masukkan Parameter:", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(8,0))
        
        # Canvas and Scrollbar setup
        canvas = tk.Canvas(right_frame, borderwidth=0, background="#ffffff")
        v_scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=v_scrollbar.set)

        v_scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True, pady=6)

        self.form_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.form_frame, anchor="nw", width=canvas.winfo_reqwidth())
        
        self.form_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas.find_all()[-1], width=e.width))

        # Output Box
        ttk.Label(right_frame, text="Hasil Rumus (siap copy-paste):", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(8,0))
        self.output_box = tk.Text(right_frame, height=4, font=("Consolas", 12), fg="#0078d4", bg="#f0f0f0", padx=5, pady=5)
        self.output_box.pack(fill="x", pady=6)

        # Tombol Aksi
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(pady=8)
        ttk.Button(btn_frame, text="✨ Buat Rumus", command=self.generate_formula, style='Accent.TButton').pack(side="left", padx=8)
        ttk.Button(btn_frame, text="📋 Salin ke Clipboard", command=self.copy_to_clipboard).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="💾 Simpan ke File", command=self.save_formula).pack(side="left", padx=8)

        # Status Bar
        self.status = tk.Label(self, text=f"Siap ✅. Pemisah default: '{ARG_DELIMITER}'", bd=1, relief="sunken", anchor="w",
                            bg="#eef3f7", font=("Segoe UI", 9))
        self.status.pack(side="bottom", fill="x")
        
        # Style
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('Accent.TButton', background='#0078d4', foreground='white', font=('Segoe UI', 10, 'bold'))
        style.map('Accent.TButton', background=[('active', '#005a9e')])


    def show_all_rumus(self):
        self.search_var.set("")
        self.listbox.delete(0, tk.END)
        kategori_dict = {}
        for name, meta in RUMUS_DB.items():
            kategori = meta.get("kategori", "Lainnya")
            if kategori not in kategori_dict:
                kategori_dict[kategori] = []
            kategori_dict[kategori].append(name)

        for kategori in sorted(kategori_dict.keys()):
            self.listbox.insert(tk.END, f"--- {kategori.upper()} ---")
            for name in sorted(kategori_dict[kategori]):
                self.listbox.insert(tk.END, name)

    def update_listbox(self, event=None):
        query = self.search_var.get().strip().upper()
        self.listbox.delete(0, tk.END)
        
        filtered_names = [name for name in RUMUS_DB.keys() if query in name.upper()]
        
        for name in sorted(filtered_names):
            self.listbox.insert(tk.END, name)
            
        if not filtered_names:
            self.listbox.insert(tk.END, "(Tidak ada hasil ditemukan)")


    def on_rumus_select(self, event):
        sel = self.listbox.curselection()
        if not sel:
            return
        
        name = self.listbox.get(sel[0])
        if name.startswith("---"): 
            return 
            
        self.selected_rumus = name
        meta = RUMUS_DB.get(name)

        if not meta:
            return
            

        self.detail_text.config(state="normal")
        self.detail_text.delete(1.0, tk.END)
        detail = f"🧾 {name}\nKategori : {meta.get('kategori','-')}\nDeskripsi: {meta.get('deskripsi','-')}\n\nContoh: {meta.get('contoh','-')}"
        self.detail_text.insert(tk.END, detail)
        self.detail_text.config(state="disabled")


        for w in self.form_frame.winfo_children():
            w.destroy()
        self.form_entries.clear()
        self.form_frame.grid_columnconfigure(0, weight=0) 
        self.form_frame.grid_columnconfigure(1, weight=1) 
        
        if meta.get("inputs"):
            for i, (field, prompt_text, default) in enumerate(meta["inputs"]):

                ttk.Label(self.form_frame, text=f"{field} :", font=("Segoe UI", 10, "bold")).grid(row=i, column=0, sticky="w", padx=5, pady=3)
                ttk.Label(self.form_frame, text=f"({prompt_text})", font=("Segoe UI", 9)).grid(row=i, column=0, sticky="w", padx=(60,0), pady=3)
                

                var = tk.StringVar(value=default if default is not None else "")
                ent = ttk.Entry(self.form_frame, textvariable=var)
                ent.grid(row=i, column=1, sticky="ew", padx=5, pady=3)
                self.form_entries[field] = var
                

        self.form_frame.update_idletasks()
        self.form_frame.master.config(scrollregion=self.form_frame.master.bbox("all"))

    def generate_formula(self):
        global ARG_DELIMITER
        
        if not self.selected_rumus:
            messagebox.showwarning("Peringatan", "Pilih rumus terlebih dahulu.")
            return
            
        meta = RUMUS_DB[self.selected_rumus]
        collected = {k: v.get().strip() for k,v in self.form_entries.items()}
        
        if self.selected_rumus == "IFS" and "pairs" in collected and collected["pairs"]:
            collected["pairs"] = collected["pairs"].replace("|", ARG_DELIMITER)
            
        formula = build_formula(meta, collected)
        
        self.output_box.config(state="normal")
        self.output_box.delete(1.0, tk.END)
        self.output_box.insert(tk.END, formula)
        self.output_box.config(state="disabled") 
        
        self.status.config(text=f"Rumus {self.selected_rumus} dibuat dengan pemisah '{ARG_DELIMITER}' ✅")
        
        if CLIP_AVAILABLE:
            try:
                pyperclip.copy(formula)
                self.status.config(text=f"Rumus disalin otomatis ke clipboard (Pemisah: '{ARG_DELIMITER}') 📋")
            except Exception:
                pass

    def copy_to_clipboard(self):

        text = self.output_box.get(1.0, tk.END).strip()
        if not text:
            messagebox.showinfo("Info", "Tidak ada rumus untuk disalin.")
            return
            
        if CLIP_AVAILABLE:
            try:
                pyperclip.copy(text)
                self.status.config(text=f"Rumus disalin ke clipboard (Pemisah: '{ARG_DELIMITER}') 📋")
            except Exception as e:
                try:
                    self.clipboard_clear()
                    self.clipboard_append(text)
                    self.status.config(text="Rumus disalin ke clipboard (tk fallback) 📋")
                except Exception as e_tk:
                    messagebox.showerror("Error", f"Gagal menyalin: {e_tk}")
        else:
            try:
                self.clipboard_clear()
                self.clipboard_append(text)
                self.status.config(text="Rumus disalin ke clipboard (tk fallback) 📋")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal menyalin: {e}")

    def save_formula(self):
        text = self.output_box.get(1.0, tk.END).strip()
        if not text:
            messagebox.showinfo("Info", "Tidak ada rumus untuk disimpan.")
            return
            
        ok, info = save_to_file(text)
        if ok:
            messagebox.showinfo("Sukses", f"Rumus disimpan ke file:\n{info}")
            self.status.config(text=f"Disimpan ke {os.path.basename(info)} 💾")
        else:
            messagebox.showerror("Error", info)

if __name__ == "__main__":
    try:
        app = ExcelRumusGUI()
        app.mainloop()
    except KeyboardInterrupt:
        print("\nDihentikan oleh user.")
        sys.exit(0)