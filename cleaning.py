"""
cleaning.py
-----------
Tugas file ini: membaca "DATA UNIT LAYANAN.xlsx" lalu membersihkannya menjadi
tabel-tabel rapi dengan kolom yang SERAGAM, supaya app.py (dashboard) cukup
memakai satu cara untuk semua sheet.

Kolom seragam di setiap tabel bersih:
    tanggal      tanggal kegiatan (kosong/NaT kalau tanggalnya tidak valid)
    kanwil_kode  kode kanwil (901-911, 900 = Kantor Pusat)
    kanwil       nama kanwil baku (dari kode)
    cabang       nama kantor cabang pairing (baku)
    ul           nama Unit Layanan (baku)
    nilai        angka utama ("jumlah layanan"; 1 per baris, kecuali sheet 2)
    nilai2       angka pendamping (nominal klaim / jumlah peserta), boleh kosong

File Excel diunggah lewat dashboard dan diproses di memori. Tidak ada file Excel atau
hasil olahan yang disimpan ke disk. Kolom data pribadi (nama, NIK, HP, email, dll)
tidak pernah dibaca, karena hanya kolom yang dibutuhkan yang diambil.
"""

import re

import pandas as pd

# ---------------------------------------------------------------------------
# 1. REFERENSI (bagian yang boleh kamu edit kalau ada perubahan di lapangan)
# ---------------------------------------------------------------------------

# Legenda kode kanwil -> nama baku (diambil dari sheet 1 yang bersumber dari sistem)
KANWIL_NAMA = {
    900: "KANTOR PUSAT",
    901: "SUMBAGUT",
    902: "SUMBAGSEL",
    903: "DKI JAKARTA",
    904: "JAWA BARAT",
    905: "JATENG DAN DIY",
    906: "JATIM",
    907: "KALIMANTAN",
    908: "SULAWESI MALUKU",
    909: "SUMBARIAU",
    910: "BANTEN",
    911: "BANUSPA",
}

# Penulisan nama kanwil yang beda-beda antar sheet -> kode
KANWIL_ALIAS = {
    "SULAMA": 908,
    "JAWA TIMUR": 906,
    "JAWA TENGAH": 905,
    **{nama: kode for kode, nama in KANWIL_NAMA.items()},
}

# Nama UL yang ditulis beda tapi sebenarnya UL yang sama (nama beda -> nama baku)
UL_ALIAS = {
    "UL MAKASSAR": "UL KOTA MAKASSAR",
    "UL PINRANG": "UL KAB PINRANG",
    "UL MALL MPP TEBO": "UL KAB TEBO",
    "UL MPP KOTA TANGERANG SELATAN": "UL KOTA TANGERANG SELATAN",
}

# Ejaan cabang yang beda antar sumber -> ejaan baku
CABANG_ALIAS = {
    "GRHA BPJAMSOSTEK": "GRAHA BPJAMSOSTEK",
    "LHOKSUMAWE": "LHOKSEUMAWE",
    "PEMATANG SIANTAR": "PEMATANGSIANTAR",
}

UL_TIDAK_TERPETAKAN = "(UL tidak terpetakan)"
CABANG_TIDAK_DIKETAHUI = "(cabang tidak diketahui)"
KANWIL_TIDAK_DIKETAHUI = "(kanwil tidak diketahui)"

# Tanggal dianggap valid bila jatuh antara TANGGAL_MULAI_VALID dan tanggal
# rekam terakhir di sheet itu (tanggal masa depan / tahun ngawur = tidak valid).
TANGGAL_MULAI_VALID = pd.Timestamp("2025-01-01")

# Batas wajar "jumlah peserta" sosialisasi (ada nilai 2.147.483.647 = error)
MAKS_PESERTA_SOSIALISASI = 100_000

# Konfigurasi tampilan tiap sheet (dipakai app.py)
SHEETS = {
    "1": {
        "judul": "Klaim JHT Biometrik",
        "grup": "JMO-Mitra",
        "satuan": "klaim",
        "tanggal_acuan": "tgl_klaim",
        "nilai2_label": "Total manfaat dibayar (Rp)",
        "nilai2_rupiah": True,
    },
    "2": {
        "judul": "Klaim JHT Lapak Asik/Antol",
        "grup": "JMO-Mitra",
        "satuan": "klaim",
        "tanggal_acuan": "tanggal",
        "nilai2_label": None,
        "nilai2_rupiah": False,
    },
    "3": {
        "judul": "Asistensi Klaim",
        "grup": "Asistensi Klaim",
        "satuan": "asistensi",
        "tanggal_acuan": "Tgl Asistensi",
        "nilai2_label": None,
        "nilai2_rupiah": False,
    },
    "4": {
        "judul": "Layanan Informasi & Pengaduan",
        "grup": "Informasi & Pengaduan",
        "satuan": "layanan",
        "tanggal_acuan": "Tgl Layanan",
        "nilai2_label": None,
        "nilai2_rupiah": False,
    },
    "5": {
        "judul": "Sosialisasi",
        "grup": "Sosialisasi",
        "satuan": "kegiatan",
        "tanggal_acuan": "Tgl Sosialisasi",
        "nilai2_label": "Total peserta sosialisasi",
        "nilai2_rupiah": False,
    },
    "6": {
        "judul": "Pendaftaran TK",
        "grup": "Perekaman",
        "satuan": "pendaftaran",
        "tanggal_acuan": "Tgl Daftar",
        "nilai2_label": None,
        "nilai2_rupiah": False,
    },
    "7": {
        "judul": "Pendaftaran Perusahaan",
        "grup": "Perekaman",
        "satuan": "pendaftaran",
        "tanggal_acuan": "Tgl Daftar",
        "nilai2_label": None,
        "nilai2_rupiah": False,
    },
}

# ---------------------------------------------------------------------------
# 2. FUNGSI PEMBERSIH KECIL
# ---------------------------------------------------------------------------

_PREFIX = "Unit Layanan BPJS Ketenagakerjaan"


def bersihkan_nama_ul(nama):
    """Rapikan nama UL jadi bentuk baku, mis. 'UL KAB BANGKA'.

    - Awalan "Unit Layanan BPJS Ketenagakerjaan" dipendekkan jadi "UL".
    - Perbaiki nama rusak akibat find-and-replace: kalau frasa itu muncul
      lagi di tengah kata (mis. "KEPUnit Layanan BPJS KetenagakerjaanAUAN"),
      itu seharusnya huruf "UL" (-> "KEPULAUAN").
    - KABUPATEN -> KAB, huruf besar semua, spasi rapi.
    - Terakhir, cek tabel UL_ALIAS untuk nama yang ditulis beda.
    """
    if pd.isna(nama):
        return pd.NA
    teks = str(nama)
    teks = re.sub(r"^\s*" + re.escape(_PREFIX) + r"\s+", "UL ", teks, flags=re.I)
    teks = teks.replace(_PREFIX, "UL")
    teks = teks.upper()
    teks = re.sub(r"^UNIT LAYANAN\s+", "UL ", teks)
    teks = re.sub(r"\bKABUPATEN\b", "KAB", teks)
    teks = re.sub(r"\bKAB\.\s*", "KAB ", teks)
    teks = re.sub(r"\s+", " ", teks).strip()
    return UL_ALIAS.get(teks, teks)


def bersihkan_cabang(nama):
    if pd.isna(nama):
        return pd.NA
    teks = re.sub(r"\s+", " ", str(nama).upper()).strip()
    return CABANG_ALIAS.get(teks, teks)


def kode_ke_int(seri):
    """Ubah kolom kode (bisa float/teks) jadi angka bulat; gagal -> kosong."""
    return pd.to_numeric(seri, errors="coerce").astype("Int64")


def nama_kanwil_dari_kode(seri_kode):
    return seri_kode.map(KANWIL_NAMA).fillna(KANWIL_TIDAK_DIKETAHUI)


def kode_kanwil_dari_nama(seri_nama):
    kunci = seri_nama.astype("string").str.upper().str.strip()
    return kunci.map(KANWIL_ALIAS).astype("Int64")


def tanggal_valid(seri, batas_akhir):
    """Ubah ke tanggal; yang rusak (0000-00-00, tahun 0006, masa depan) -> NaT."""
    t = pd.to_datetime(seri, errors="coerce")
    ok = (t >= TANGGAL_MULAI_VALID) & (t <= batas_akhir)
    return t.where(ok).dt.normalize()


def tanggal_dengan_cadangan(seri_kegiatan, seri_rekam, batas_akhir):
    """Tanggal kegiatan; kalau rusak, pakai Tgl Rekam (tanggal input di sistem).

    Dengan begitu baris yang salah ketik tanggalnya tidak hilang dari total.
    Mengembalikan (tanggal, jumlah baris yang diganti dengan tanggal rekam).
    """
    utama = tanggal_valid(seri_kegiatan, batas_akhir)
    cadangan = tanggal_valid(seri_rekam, batas_akhir)
    diganti = int((utama.isna() & cadangan.notna()).sum())
    return utama.fillna(cadangan), diganti


def _koordinat(seri):
    """'-6.26, 106.52' -> (lat, lon) sebagai angka; yang gagal -> kosong."""
    pecah = seri.astype("string").str.extract(r"(-?\d+(?:\.\d+)?)\s*[,;]\s*(-?\d+(?:\.\d+)?)")
    return pd.to_numeric(pecah[0]), pd.to_numeric(pecah[1])


# ---------------------------------------------------------------------------
# 3. MEMBACA EXCEL
# ---------------------------------------------------------------------------


def _baca(xl, sheet, wajib, opsional=()):
    """Baca satu sheet dari pd.ExcelFile, hanya kolom yang dibutuhkan (lebih cepat).

    Baris judul kolom dicari otomatis di 6 baris pertama (ada sheet yang punya
    baris judul laporan di atasnya, ada yang tidak). Kolom `wajib` harus ada;
    kolom `opsional` dibaca kalau ada.
    """
    if sheet not in xl.sheet_names:
        raise ValueError(f"Sheet '{sheet}' tidak ditemukan di file. Sheet yang ada: {xl.sheet_names}")
    wajib, opsional = set(wajib), set(opsional)
    atas = xl.parse(sheet, header=None, nrows=6)
    baris_judul = None
    for i in range(len(atas)):
        isi = {str(v).strip() for v in atas.iloc[i].tolist() if pd.notna(v)}
        if wajib <= isi:
            baris_judul = i
            break
    if baris_judul is None:
        ditemukan = [
            [str(v).strip() for v in atas.iloc[i].tolist() if pd.notna(v)] for i in range(min(3, len(atas)))
        ]
        raise ValueError(
            f"Sheet '{sheet}': kolom {sorted(wajib)} tidak ditemukan bersamaan di 6 baris pertama. "
            f"Isi 3 baris pertama yang terbaca: {ditemukan}"
        )
    diambil = wajib | opsional
    df = xl.parse(sheet, header=baris_judul, usecols=lambda c: str(c).strip() in diambil)
    df.columns = [str(c).strip() for c in df.columns]
    for kolom in opsional - set(df.columns):  # kolom opsional yang tidak ada -> kosong
        df[kolom] = pd.NA
    return df


# ---------------------------------------------------------------------------
# 4. MASTER UNIT LAYANAN & PEMETAAN
# ---------------------------------------------------------------------------


def buat_master_ul(s01, s20, s1, s2):
    """Gabungkan semua sumber nama UL jadi satu master (union).

    Prioritas atribut kanwil/cabang untuk satu UL:
      1) sheet 1 (dari sistem, pairing terbaru)
      2) sheet 0.1 (master lama; satu-satunya sumber koordinat)
      3) sheet 2.0 / sheet 2 (hanya kanwil)
    """
    kosong = pd.NA
    # sumber 1: sheet 1
    a = pd.DataFrame(
        {
            "ul": s1["nama_mitra"].map(bersihkan_nama_ul),
            "kanwil_kode": kode_ke_int(s1["kode_wilayah_pairing"]),
            "cabang": s1["nama_kantor_pairing"].map(bersihkan_cabang),
            "lat": kosong, "lon": kosong,
            "prioritas": 1,
        }
    ).drop_duplicates("ul")

    # sumber 2: sheet 0.1
    lat, lon = _koordinat(s01["LATITUDE, LONGTITUDE PENEMPATAN"])
    b = pd.DataFrame(
        {
            "ul": s01["Unit Layanan BPJS Ketenagakerjaan NAME"].map(bersihkan_nama_ul),
            "kanwil_kode": kode_kanwil_dari_nama(s01["KANWIL"]),
            "cabang": s01["NAMA KANTOR CABANG PAIRING"].map(bersihkan_cabang),
            "lat": lat, "lon": lon,
            "prioritas": 2,
        }
    )

    # sumber 3: sheet 2.0 dan sheet 2 (kanwil saja)
    c = pd.DataFrame(
        {
            "ul": s20["UL NAME"].map(bersihkan_nama_ul),
            "kanwil_kode": kode_kanwil_dari_nama(s20["KANWIL"]),
            "cabang": kosong, "lat": kosong, "lon": kosong,
            "prioritas": 3,
        }
    )
    d = pd.DataFrame(
        {
            "ul": s2["unit"].map(bersihkan_nama_ul),
            "kanwil_kode": kode_kanwil_dari_nama(s2["kanwil"]),
            "cabang": kosong, "lat": kosong, "lon": kosong,
            "prioritas": 4,
        }
    ).drop_duplicates("ul")

    gabung = pd.concat([a, b, c, d], ignore_index=True).dropna(subset=["ul"])
    gabung = gabung.sort_values("prioritas")
    gabung[["lat", "lon"]] = gabung[["lat", "lon"]].apply(pd.to_numeric, errors="coerce")

    # per UL: ambil nilai pertama yang TIDAK kosong menurut urutan prioritas
    master = gabung.groupby("ul", as_index=False).agg(
        kanwil_kode=("kanwil_kode", "first"),
        cabang=("cabang", "first"),
        lat=("lat", "first"),
        lon=("lon", "first"),
    )
    master["cabang"] = master["cabang"].fillna(CABANG_TIDAK_DIKETAHUI)
    master["kanwil"] = nama_kanwil_dari_kode(master["kanwil_kode"].astype("Int64"))
    return master


def buat_peta_cabang(s1):
    """Kode cabang (IDKC) -> nama cabang, diambil dari sheet 1."""
    p = s1[["kode_kantor_pairing", "nama_kantor_pairing"]].drop_duplicates("kode_kantor_pairing")
    return dict(
        zip(
            p["kode_kantor_pairing"].astype(str).str.strip(),
            p["nama_kantor_pairing"].map(bersihkan_cabang),
        )
    )


def buat_peta_user(s01, s20):
    """Kode akun petugas (ULBPJSTK###) -> nama UL.

    Sumber: kolom 'KODE USER' di sheet 0.1 dan 'USER ID' di sheet 2.0.
    Satu sel boleh berisi lebih dari satu kode; semuanya diambil.
    """
    peta = {}
    pasangan = [
        (s01["KODE USER"], s01["Unit Layanan BPJS Ketenagakerjaan NAME"]),
        (s20["USER ID"], s20["UL NAME"]),
    ]
    for kolom_kode, kolom_nama in pasangan:
        for kode, nama in zip(kolom_kode, kolom_nama):
            if pd.isna(kode) or pd.isna(nama):
                continue
            for k in re.findall(r"ULBPJSTK\d+", str(kode).upper()):
                peta.setdefault(k, bersihkan_nama_ul(nama))
    return peta


# ---------------------------------------------------------------------------
# 5. MEMBANGUN TABEL BERSIH UNTUK TIAP SHEET
# ---------------------------------------------------------------------------


def _seragamkan(df):
    kolom = ["tanggal", "kanwil_kode", "kanwil", "cabang", "ul", "nilai", "nilai2"]
    if "status" in df.columns:  # hanya sheet 1 yang punya kolom status
        kolom.append("status")
    return df[kolom].reset_index(drop=True)


def _tambah_kanwil_nama(df):
    df["kanwil"] = nama_kanwil_dari_kode(df["kanwil_kode"])
    return df


def _tabel_dengan_kode(df, kol_idkw, kol_idkc, kol_user, kol_tanggal, peta_cabang, peta_user):
    """Sheet 3/4/5/6/7: punya kode kanwil + kode cabang + akun petugas."""
    out = pd.DataFrame(index=df.index)
    out["tanggal"], diganti = tanggal_dengan_cadangan(df[kol_tanggal], df["Tgl Rekam"], df["_batas"].iloc[0] if len(df) else pd.Timestamp.max)
    out["kanwil_kode"] = kode_ke_int(df[kol_idkw])
    out = _tambah_kanwil_nama(out)
    kode_cabang = df[kol_idkc].astype("string").str.strip()
    out["cabang"] = (
        kode_cabang.map(peta_cabang)
        .fillna(kode_cabang.map(lambda k: f"kode {k}" if pd.notna(k) else pd.NA))
        .fillna(CABANG_TIDAK_DIKETAHUI)
    )
    user = df[kol_user].astype("string").str.upper().str.strip()
    out["ul"] = user.map(peta_user).fillna(UL_TIDAK_TERPETAKAN)
    out["nilai"] = 1
    out["nilai2"] = pd.NA
    out.attrs["tanggal_diganti"] = diganti
    return out


def proses_excel(sumber, nama_file="", progress=None):
    """Baca Excel (path atau file hasil unggahan) -> bersihkan.

    Mengembalikan (tabel, master, legenda_cabang, catatan):
      tabel          dict sheet -> DataFrame bersih (kolom seragam)
      master         tabel UL (kanwil, cabang, koordinat)
      legenda_cabang kode cabang -> nama cabang
      catatan        ringkasan kualitas data
    Tidak ada yang ditulis ke disk.
    """

    def lapor(teks):
        if progress:
            progress(teks)

    xl = pd.ExcelFile(sumber)

    # --- referensi
    lapor("Membaca sheet referensi (0.1, 2.0)...")
    s01 = _baca(
        xl, "0.1",
        {"KANWIL", "Unit Layanan BPJS Ketenagakerjaan NAME", "NAMA KANTOR CABANG PAIRING"},
        {"KODE USER", "LATITUDE, LONGTITUDE PENEMPATAN"},
    )
    s20 = _baca(xl, "2.0", {"KANWIL", "UL NAME"}, {"USER ID"})

    # --- sheet 1 (klaim biometrik) dan sheet 2 (Lapak Asik)
    lapor("Membaca sheet 1 (klaim biometrik)...")
    s1 = _baca(
        xl, "1",
        {
            "kode_pengajuan", "nama_mitra", "kode_wilayah_pairing", "kode_kantor_pairing",
            "nama_kantor_pairing", "tgl_klaim", "tgl_rekam", "status_pengajuan", "nom_manfaat_netto",
        },
    )
    lapor("Membaca sheet 2 (Lapak Asik)...")
    s2 = _baca(xl, "2", {"tanggal", "kanwil", "unit", "tgl_rekam", "jmlh_data"})

    master = buat_master_ul(s01, s20, s1, s2)
    peta_cabang = buat_peta_cabang(s1)
    peta_user = buat_peta_user(s01, s20)
    master_idx = master.set_index("ul")

    catatan = {"sheet": {}, "peta_user_jumlah": len(peta_user), "excel": nama_file}
    hasil = {}

    def simpan(kunci, df, jumlah_mentah, diganti=0, ekstra=None):
        df = _seragamkan(df)
        total = len(df)
        ul_ok = int((df["ul"] != UL_TIDAK_TERPETAKAN).sum())
        ada_tanggal = df["tanggal"].notna().any()
        catatan["sheet"][kunci] = {
            "baris_mentah": int(jumlah_mentah),
            "baris_dipakai": int(total),
            "tanggal_diganti_tgl_rekam": int(diganti),
            "tanggal_tidak_valid": int(df["tanggal"].isna().sum()),
            "persen_ul_terpetakan": round(100 * ul_ok / total, 1) if total else 0,
            "tgl_min": str(df["tanggal"].min().date()) if ada_tanggal else "",
            "tgl_maks": str(df["tanggal"].max().date()) if ada_tanggal else "",
            **(ekstra or {}),
        }
        hasil[kunci] = df

    # ------------------ sheet 1
    # Semua status ikut dihitung di halaman Layanan (sama dengan pivot 1.1/1.2 dari user);
    # ada pilihan "hanya DIBAYAR" lewat kolom status. Halaman Home membuang DIBATALKAN.
    lapor("Membersihkan sheet 1...")
    batas = pd.to_datetime(s1["tgl_rekam"], errors="coerce").max().normalize()
    t = pd.DataFrame(index=s1.index)
    t["tanggal"], diganti = tanggal_dengan_cadangan(s1["tgl_klaim"], s1["tgl_rekam"], batas)
    t["kanwil_kode"] = kode_ke_int(s1["kode_wilayah_pairing"])
    t = _tambah_kanwil_nama(t)
    t["cabang"] = s1["nama_kantor_pairing"].map(bersihkan_cabang)
    t["ul"] = s1["nama_mitra"].map(bersihkan_nama_ul)
    t["nilai"] = 1
    t["status"] = s1["status_pengajuan"].astype("string").str.upper().str.strip()
    nominal = pd.to_numeric(s1["nom_manfaat_netto"], errors="coerce")
    t["nilai2"] = nominal.where(t["status"] == "DIBAYAR")  # nominal hanya untuk yang dibayar
    bukan_bayar = int((t["status"] != "DIBAYAR").sum())
    simpan(
        "1", t, len(s1), diganti,
        {"catatan": f"Semua status pengajuan dihitung (sama dengan pivot 1.1/1.2); {bukan_bayar} baris berstatus bukan DIBAYAR."},
    )

    # ------------------ sheet 2
    lapor("Membersihkan sheet 2...")
    batas = pd.to_datetime(s2["tgl_rekam"], errors="coerce").max().normalize()
    t = pd.DataFrame(index=s2.index)
    t["tanggal"], diganti = tanggal_dengan_cadangan(s2["tanggal"], s2["tgl_rekam"], batas)
    t["ul"] = s2["unit"].map(bersihkan_nama_ul)
    kode_dari_master = t["ul"].map(master_idx["kanwil_kode"])
    kode_dari_sheet = kode_kanwil_dari_nama(s2["kanwil"])
    t["kanwil_kode"] = kode_dari_master.fillna(kode_dari_sheet).astype("Int64")
    t = _tambah_kanwil_nama(t)
    t["cabang"] = t["ul"].map(master_idx["cabang"]).fillna(CABANG_TIDAK_DIKETAHUI)
    t["ul"] = t["ul"].fillna(UL_TIDAK_TERPETAKAN)
    t["nilai"] = pd.to_numeric(s2["jmlh_data"], errors="coerce").fillna(0)
    t["nilai2"] = pd.NA
    simpan("2", t, len(s2), diganti, {"catatan": "Jumlah klaim = jumlah kolom jmlh_data (bukan jumlah baris)."})
    del s1, s2

    # ------------------ sheet 3, 4, 6, 7 (berkode)
    spesifik = {
        "3": ("Tgl Asistensi", {"IDKW", "IDKC", "User Petugas", "Tgl Rekam", "Tgl Asistensi"}),
        "4": ("Tgl Layanan", {"IDKW", "IDKC", "User Petugas", "Tgl Rekam", "Tgl Layanan"}),
        "6": ("Tgl Daftar", {"IDKW", "IDKC", "User Petugas", "Tgl Rekam", "Tgl Daftar"}),
        "7": ("Tgl Daftar", {"IDKW", "IDKC", "User Petugas", "Tgl Rekam", "Tgl Daftar"}),
    }
    for kunci, (kol_tgl, kolom) in spesifik.items():
        lapor(f"Membaca dan membersihkan sheet {kunci} ({SHEETS[kunci]['judul']})...")
        df = _baca(xl, kunci, kolom)
        df["_batas"] = pd.to_datetime(df["Tgl Rekam"], errors="coerce").max().normalize()
        t = _tabel_dengan_kode(df, "IDKW", "IDKC", "User Petugas", kol_tgl, peta_cabang, peta_user)
        simpan(kunci, t, len(df), t.attrs.get("tanggal_diganti", 0))
        del df

    # ------------------ sheet 5 (sosialisasi): kolomnya beda sendiri
    lapor("Membaca dan membersihkan sheet 5 (Sosialisasi)...")
    df = _baca(xl, "5", {"Kanwil", "Pairing", "User Petugas", "Tgl Rekam", "Tgl Sosialisasi", "Jml Peserta"})
    df["_batas"] = pd.to_datetime(df["Tgl Rekam"], errors="coerce").max().normalize()
    t = _tabel_dengan_kode(df, "Kanwil", "Pairing", "User Petugas", "Tgl Sosialisasi", peta_cabang, peta_user)
    peserta = pd.to_numeric(df["Jml Peserta"], errors="coerce")
    tidak_wajar = (peserta < 0) | (peserta > MAKS_PESERTA_SOSIALISASI)
    t["nilai2"] = peserta.where(~tidak_wajar)
    simpan(
        "5", t, len(df), t.attrs.get("tanggal_diganti", 0),
        {"catatan": f"{int(tidak_wajar.sum())} baris jumlah peserta tidak wajar (negatif / > {MAKS_PESERTA_SOSIALISASI:,}) tidak ikut dijumlahkan."},
    )

    # ------------------ legenda & catatan
    legenda_cabang = pd.DataFrame({"kode_cabang": list(peta_cabang.keys()), "nama_cabang": list(peta_cabang.values())})
    legenda_cabang = legenda_cabang.sort_values("kode_cabang").reset_index(drop=True)

    # UL yang muncul di data transaksi tetapi tidak ada di master 0.1 (master lama)
    nama_di_01 = set(s01["Unit Layanan BPJS Ketenagakerjaan NAME"].map(bersihkan_nama_ul).dropna())
    catatan["ul_tidak_ada_di_sheet_0_1"] = sorted(set(master["ul"]) - nama_di_01)
    catatan["jumlah_ul_master"] = int(len(master))
    catatan["ul_dengan_koordinat"] = int(master["lat"].notna().sum())
    lapor("Selesai.")
    return hasil, master, legenda_cabang, catatan
