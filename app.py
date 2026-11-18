"""
app.py
------
Dashboard monitoring Unit Layanan (UL) BPJS Ketenagakerjaan, dibuat dengan Streamlit.

Cara menjalankan (di Terminal, dari folder dashboard_ul, dengan .venv aktif):
    streamlit run app.py

Cara memakai:
    1. Unggah file Excel (DATA UNIT LAYANAN.xlsx) di bagian atas halaman.
       Dashboard baru muncul setelah file diunggah. File hanya diproses di memori,
       tidak disimpan ke disk.
    2. Pilih halaman:
         Home     score card per kelompok layanan + peta lokasi Unit Layanan
         Layanan  per data (sheet): jumlah layanan, top 10, grafik per bulan
    3. Filter (melintang di atas, berlaku untuk kedua halaman):
         Rentang waktu, Kantor wilayah pairing, Kantor cabang pairing, Unit Layanan

File yang dibutuhkan di folder yang sama: cleaning.py dan peta_indonesia.svg
"""

import io
import json
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

import cleaning

BASE_DIR = Path(__file__).resolve().parent
FILE_PETA = BASE_DIR / "peta_indonesia.svg"

WARNA = "#2F6FB2"
BULAN_ID = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
DIMENSI = {"Unit Layanan": "ul", "Cabang pairing": "cabang", "Kanwil pairing": "kanwil"}
KELOMPOK = ["Informasi & Pengaduan", "Perekaman", "Asistensi Klaim", "Sosialisasi", "JMO-Mitra"]

st.set_page_config(page_title="Monitoring Unit Layanan", layout="wide", initial_sidebar_state="collapsed")


# ---------------------------------------------------------------------------
# Format angka gaya Indonesia (titik ribuan, koma desimal)
# ---------------------------------------------------------------------------
def fmt_int(n):
    return f"{int(round(float(n))):,}".replace(",", ".")


def fmt_rp(n):
    n = float(n)
    if n >= 1e12:
        return "Rp " + f"{n / 1e12:.2f}".replace(".", ",") + " triliun"
    if n >= 1e9:
        return "Rp " + f"{n / 1e9:.2f}".replace(".", ",") + " miliar"
    if n >= 1e6:
        return "Rp " + f"{n / 1e6:.1f}".replace(".", ",") + " juta"
    return "Rp " + fmt_int(n)


def label_bulan(periode):
    return f"{BULAN_ID[periode.month - 1]} {periode.year}"


# ---------------------------------------------------------------------------
# Mengolah Excel yang diunggah (hasilnya diingat di memori selama file yang sama)
# ---------------------------------------------------------------------------
def gabung_home(tabel):
    """Satukan semua sheet jadi satu tabel untuk halaman Home (klaim DIBATALKAN tidak dihitung)."""
    bagian = []
    for kunci, df in tabel.items():
        d = df[["tanggal", "kanwil", "cabang", "ul", "nilai"]].copy()
        if kunci == "1":
            d = d[~df["status"].fillna("").str.contains("BATAL")]
        d["grup"] = cleaning.SHEETS[kunci]["grup"]
        bagian.append(d)
    return pd.concat(bagian, ignore_index=True)


@st.cache_resource(show_spinner=False, max_entries=1)
def olah(isi: bytes, nama: str):
    tabel, master, legenda_cabang, catatan = cleaning.proses_excel(io.BytesIO(isi), nama)
    return tabel, master, legenda_cabang, catatan, gabung_home(tabel)


@st.cache_resource(show_spinner=False)
def baca_peta():
    return FILE_PETA.read_text(encoding="utf-8") if FILE_PETA.exists() else None


# ---------------------------------------------------------------------------
# Bagian atas halaman: judul dan unggah Excel
# ---------------------------------------------------------------------------
st.title("Monitoring Unit Layanan BPJS Ketenagakerjaan")

sudah_unggah = st.session_state.get("berkas_excel") is not None
with st.expander("File Excel data", expanded=not sudah_unggah):
    berkas = st.file_uploader(
        "Unggah file DATA UNIT LAYANAN.xlsx",
        type=["xlsx"],
        key="berkas_excel",
        help="File hanya diproses di memori dan tidak disimpan. Unggah file baru kapan saja untuk memperbarui dashboard.",
    )

if berkas is None:
    st.info(
        "Unggah file Excel di atas untuk menampilkan dashboard. "
        "File hanya diproses di memori dan tidak disimpan ke disk."
    )
    st.stop()

try:
    with st.spinner("Membaca dan membersihkan Excel. Bisa sampai sekitar 1 menit untuk file baru..."):
        tabel, master, legenda_cabang, catatan, home = olah(berkas.getvalue(), berkas.name)
except Exception as err:  # file bukan Excel yang benar, sheet atau kolom tidak ada, dll
    st.error(f"Gagal membaca file: {err}")
    st.stop()

semua_tgl = pd.concat([d["tanggal"] for d in tabel.values()]).dropna()
tgl_min, tgl_maks = semua_tgl.min().date(), semua_tgl.max().date()
tgl_awal_default = max(tgl_min, pd.Timestamp(year=tgl_maks.year, month=1, day=1).date())

st.caption(
    f"Sumber: {catatan['excel']}  ·  {catatan['jumlah_ul_master']} Unit Layanan dalam master  ·  "
    f"rentang data {tgl_min:%d-%m-%Y} s/d {tgl_maks:%d-%m-%Y}"
)

# --- Pilihan halaman
halaman = st.segmented_control("Halaman", ["Home", "Layanan"], default="Home", key="halaman", label_visibility="collapsed")
halaman = halaman or "Home"

# --- Dropdown pilihan data (hanya di halaman Layanan)
PILIHAN_CATATAN = "Catatan & Legenda"
judul_ke_kunci = {cfg["judul"]: k for k, cfg in cleaning.SHEETS.items()}
pilihan_data = None
if halaman == "Layanan":
    kolom_pilih, _kosong = st.columns([2, 3])
    pilihan_data = kolom_pilih.selectbox("Pilih data", list(judul_ke_kunci) + [PILIHAN_CATATAN], key="pilihan_data")

# Daftar pilihan filter dibuat dari gabungan master UL dan semua sheet
dims = pd.concat(
    [master[["kanwil", "cabang", "ul"]]] + [d[["kanwil", "cabang", "ul"]].drop_duplicates() for d in tabel.values()],
    ignore_index=True,
).drop_duplicates()


def urut(seri):
    """Urutkan alfabet, pilihan '(tidak diketahui)' ditaruh paling bawah."""
    return sorted(set(seri.dropna()), key=lambda x: (x.startswith("("), x))


def pilih(wadah, label, opsi, kunci):
    # buang pilihan lama yang sudah tidak ada di daftar (karena filter di sebelah kirinya berubah)
    if kunci in st.session_state:
        st.session_state[kunci] = [x for x in st.session_state[kunci] if x in opsi]
    return wadah.multiselect(label, opsi, key=kunci, placeholder="Semua")


# --- Filter melintang (tidak tampil untuk Catatan & Legenda)
mulai, akhir = pd.Timestamp(tgl_awal_default), pd.Timestamp(tgl_maks)
f_kanwil, f_cabang, f_ul = [], [], []

if pilihan_data != PILIHAN_CATATAN:
    with st.container(border=True):
        k_waktu, k_kanwil, k_cabang, k_ul = st.columns([1.3, 1, 1, 1])
        rentang = k_waktu.date_input(
            "Rentang waktu",
            value=(tgl_awal_default, tgl_maks),
            min_value=tgl_min,
            max_value=tgl_maks,
            key="f_waktu",
            help="Tanggal acuan tiap data berbeda (lihat keterangan di halaman Layanan).",
        )
        if not isinstance(rentang, (tuple, list)) or len(rentang) != 2:
            k_waktu.info("Pilih tanggal akhir.")
            st.stop()
        mulai, akhir = pd.Timestamp(rentang[0]), pd.Timestamp(rentang[1])

        f_kanwil = pilih(k_kanwil, "Kantor wilayah pairing", urut(dims["kanwil"]), "f_kanwil")
        d1 = dims[dims["kanwil"].isin(f_kanwil)] if f_kanwil else dims
        f_cabang = pilih(k_cabang, "Kantor cabang pairing", urut(d1["cabang"]), "f_cabang")
        d2 = d1[d1["cabang"].isin(f_cabang)] if f_cabang else d1
        opsi_ul = urut(d2.loc[d2["ul"] != cleaning.UL_TIDAK_TERPETAKAN, "ul"])
        f_ul = pilih(k_ul, "Unit Layanan", opsi_ul, "f_ul")


# ---------------------------------------------------------------------------
# Fungsi bantu tampilan
# ---------------------------------------------------------------------------
def terapkan_filter(df, punya_ul):
    m = df["tanggal"].between(mulai, akhir)  # tanggal kosong otomatis tidak lolos
    if f_kanwil:
        m &= df["kanwil"].isin(f_kanwil)
    if f_cabang:
        m &= df["cabang"].isin(f_cabang)
    if f_ul and punya_ul:
        m &= df["ul"].isin(f_ul)
    return df[m]


def tabel_ranking(df, kolom):
    r = df.groupby(kolom, as_index=False)["nilai"].sum().sort_values("nilai", ascending=False)
    r.columns = ["nama", "jumlah"]
    return r.reset_index(drop=True)


def grafik_top10(top):
    top = top.copy()
    top["teks"] = top["jumlah"].map(fmt_int)
    dasar = alt.Chart(top).encode(
        y=alt.Y("nama:N", sort=alt.EncodingSortField("jumlah", order="descending"), title=None,
                axis=alt.Axis(labelLimit=280)),
        x=alt.X("jumlah:Q", axis=None),
        tooltip=[alt.Tooltip("nama:N", title="Nama"), alt.Tooltip("teks:N", title="Jumlah")],
    )
    batang = dasar.mark_bar(color=WARNA, cornerRadiusEnd=3)
    angka = dasar.mark_text(align="left", dx=4, color="#444").encode(text="teks:N")
    return (batang + angka).properties(height=max(60, 30 * len(top))).configure_view(strokeWidth=0)


def grafik_bulanan(df):
    per_bulan = df.groupby(df["tanggal"].dt.to_period("M"))["nilai"].sum()
    semua_bulan = pd.period_range(mulai.to_period("M"), akhir.to_period("M"), freq="M")
    per_bulan = per_bulan.reindex(semua_bulan, fill_value=0)
    data = pd.DataFrame(
        {"bulan": [label_bulan(p) for p in per_bulan.index], "jumlah": per_bulan.values}
    )
    data["teks"] = data["jumlah"].map(fmt_int)
    urutan = list(data["bulan"])
    batas_atas = max(1, float(data["jumlah"].max()) * 1.15)  # ruang di atas untuk angka
    dasar = alt.Chart(data).encode(
        x=alt.X("bulan:N", sort=urutan, title=None, axis=alt.Axis(labelAngle=-45)),
        y=alt.Y("jumlah:Q", title=None, scale=alt.Scale(domain=[0, batas_atas]),
                axis=alt.Axis(format="~s", grid=True)),
        tooltip=[alt.Tooltip("bulan:N", title="Bulan"), alt.Tooltip("teks:N", title="Jumlah")],
    )
    garis = dasar.mark_line(color=WARNA, strokeWidth=2.5, point=alt.OverlayMarkDef(
        filled=True, fill=WARNA, size=80, stroke="white", strokeWidth=1.5))
    chart = garis
    if len(data) <= 14:
        chart = garis + dasar.mark_text(dy=-14, color="#444").encode(text="teks:N")
    return chart.properties(height=320).configure_view(strokeWidth=0), bool((data["jumlah"] == 0).any())


# ---------------------------------------------------------------------------
# HALAMAN HOME: score card + peta
# ---------------------------------------------------------------------------
def html_peta(titik, svg):
    """Peta Indonesia (SVG) + titik per UL. titik = [nama, kanwil, cabang, nilai, cx, cy, radius]."""
    data_json = json.dumps(titik, ensure_ascii=False).replace("</", "<\\/")
    svg = svg.replace("</svg>", '<g id="titik"></g></svg>')
    return f"""
<style>
  body {{ margin:0; font:13px system-ui,-apple-system,"Segoe UI",Roboto,sans-serif; }}
  svg {{ width:100%; height:auto; display:block; }}
  .land {{ fill:#dde5ee; stroke:#ffffff; stroke-width:.6; }}
  circle {{ fill:{WARNA}; fill-opacity:.65; stroke:#ffffff; stroke-width:.6; cursor:pointer; }}
  circle.nol {{ fill-opacity:.2; }}
  #tip {{ position:fixed; pointer-events:none; background:#1a2230; color:#fff; padding:6px 9px;
          border-radius:7px; font-size:12px; display:none; z-index:9; max-width:240px; }}
</style>
{svg}
<div id="tip"></div>
<script>
const P = {data_json};
const NS = "http://www.w3.org/2000/svg";
const g = document.getElementById("titik"), tip = document.getElementById("tip");
const esc = s => String(s).replace(/[&<>"]/g, c => ({{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}}[c]));
P.forEach(p => {{
  const c = document.createElementNS(NS, "circle");
  c.setAttribute("cx", p[4]); c.setAttribute("cy", p[5]); c.setAttribute("r", p[6]);
  if (!p[3]) c.setAttribute("class", "nol");
  c.addEventListener("mousemove", e => {{
    tip.innerHTML = "<b>" + esc(p[0]) + "</b><br>" + esc(p[1]) + " · " + esc(p[2]) + "<br>" +
                    p[3].toLocaleString("id-ID") + " layanan";
    tip.style.display = "block";
    tip.style.left = Math.min(e.clientX + 12, innerWidth - 250) + "px";
    tip.style.top = (e.clientY + 12) + "px";
  }});
  c.addEventListener("mouseleave", () => tip.style.display = "none");
  g.appendChild(c);
}});
</script>
"""


def tampilkan_peta(html):
    if hasattr(st, "iframe"):  # Streamlit baru: ukuran mengikuti isi
        st.iframe(html, height="content")
    else:  # Streamlit lama
        import streamlit.components.v1 as components

        components.html(html, height=520)


def tampilkan_home():
    # --- data setelah filter
    m = home["tanggal"].between(mulai, akhir)
    if f_kanwil:
        m &= home["kanwil"].isin(f_kanwil)
    if f_cabang:
        m &= home["cabang"].isin(f_cabang)
    if f_ul:
        m &= home["ul"].isin(f_ul)
    dh = home[m]
    dh_ul = dh[dh["ul"] != cleaning.UL_TIDAK_TERPETAKAN]
    nilai_ul = dh_ul.groupby("ul")["nilai"].sum()
    nilai_ul = nilai_ul[nilai_ul > 0]

    total = dh["nilai"].sum()
    aktif = len(nilai_ul)
    rata = nilai_ul.sum() / aktif if aktif else 0
    per_grup = dh.groupby("grup")["nilai"].sum().reindex(KELOMPOK).fillna(0)

    st.caption(f"Periode {mulai:%d-%m-%Y} s/d {akhir:%d-%m-%Y}  ·  jumlah layanan per Unit Layanan")

    # --- score card
    baris1, baris2 = st.columns(4), st.columns(4)
    baris1[0].metric("Total layanan", fmt_int(total), help="Gabungan semua data. Klaim berstatus DIBATALKAN tidak dihitung.")
    baris1[1].metric("Unit Layanan aktif", fmt_int(aktif), help="UL yang punya minimal 1 layanan pada filter ini.")
    baris1[2].metric("Rata-rata per UL aktif", fmt_int(rata), help="Total layanan yang terhubung ke UL, dibagi jumlah UL aktif.")
    baris1[3].metric(KELOMPOK[0], fmt_int(per_grup[KELOMPOK[0]]))
    for kolom, nama in zip(baris2, KELOMPOK[1:]):
        kolom.metric(nama, fmt_int(per_grup[nama]))

    tanpa_ul = total - dh_ul["nilai"].sum()
    if tanpa_ul > 0:
        st.caption(
            f"{fmt_int(tanpa_ul)} layanan ({tanpa_ul / total:.1%}) belum bisa dikaitkan ke Unit Layanan "
            "(akun petugas belum terpetakan). Angka itu ikut di total, tapi tidak di UL aktif, rata-rata, dan peta."
        )

    # --- peta
    st.write("")
    with st.container(border=True):
        st.markdown("**Peta lokasi Unit Layanan**  ·  ukuran titik = jumlah layanan")
        svg = baca_peta()
        if svg is None:
            st.warning("File peta_indonesia.svg tidak ditemukan di folder aplikasi, jadi peta tidak bisa ditampilkan.")
            return
        mst = master
        if f_kanwil:
            mst = mst[mst["kanwil"].isin(f_kanwil)]
        if f_cabang:
            mst = mst[mst["cabang"].isin(f_cabang)]
        if f_ul:
            mst = mst[mst["ul"].isin(f_ul)]
        mst = mst.assign(v=mst["ul"].map(nilai_ul).fillna(0))
        mst = mst.dropna(subset=["lat", "lon"])
        mst = mst.assign(cx=(mst["lon"] - 94.5) * 20, cy=(6.5 - mst["lat"]) * 20)
        mst = mst[mst["cx"].between(0, 940) & mst["cy"].between(0, 360)]  # di luar Indonesia tidak digambar
        terbesar = max(1.0, float(mst["v"].max())) if len(mst) else 1.0
        titik = [
            [r.ul, r.kanwil, r.cabang, int(r.v), round(r.cx, 1), round(r.cy, 1), round(2.2 + (r.v / terbesar) ** 0.5 * 7, 1)]
            for r in mst.itertuples()
        ]
        tampilkan_peta(html_peta(titik, svg))
        jumlah_filter = len(master) if not (f_kanwil or f_cabang or f_ul) else int(
            master["kanwil"].isin(f_kanwil or master["kanwil"]).__and__(master["cabang"].isin(f_cabang or master["cabang"])).__and__(master["ul"].isin(f_ul or master["ul"])).sum()
        )
        tidak_tampil = jumlah_filter - len(titik)
        st.caption(
            f"{len(titik)} UL tampil di peta."
            + (f" {tidak_tampil} UL tanpa koordinat atau berada di luar Indonesia tidak ditampilkan." if tidak_tampil else "")
        )

    st.caption(
        "Layanan = pelayanan informasi dan pengaduan, perekaman tenaga kerja dan perusahaan, asistensi klaim, "
        "sosialisasi, antrean online (ANTOL), dan klaim JMO-Mitra (tidak termasuk yang dibatalkan)."
    )


# ---------------------------------------------------------------------------
# HALAMAN LAYANAN: isi satu data (dipakai berulang untuk 7 sheet)
# ---------------------------------------------------------------------------
def tampilkan_sheet(kunci):
    cfg = cleaning.SHEETS[kunci]
    q = catatan["sheet"][kunci]
    punya_ul = q["persen_ul_terpetakan"] >= 50
    df = tabel[kunci]

    st.subheader(cfg["judul"])
    st.caption(f"Tanggal acuan: **{cfg['tanggal_acuan']}**  ·  {fmt_int(q['baris_dipakai'])} baris data sumber")

    if kunci == "1":
        if st.checkbox("Hanya klaim berstatus DIBAYAR", key="hanya_dibayar",
                       help="Kalau tidak dicentang, semua pengajuan dihitung (sama dengan pivot user)."):
            df = df[df["status"] == "DIBAYAR"]

    if f_ul and not punya_ul:
        st.warning(
            "Filter **Unit Layanan** belum berlaku di sheet ini, karena akun petugas belum bisa "
            "dipetakan ke nama UL. Filter kanwil, cabang, dan waktu tetap berlaku."
        )

    df = terapkan_filter(df, punya_ul)
    if df.empty:
        st.info("Tidak ada data untuk kombinasi filter ini.")
        return

    # --- 1. Kartu angka
    total = df["nilai"].sum()
    jumlah_bulan = (akhir.year - mulai.year) * 12 + akhir.month - mulai.month + 1
    kartu = st.columns(4)
    kartu[0].metric(f"Jumlah {cfg['satuan']}", fmt_int(total))
    kartu[1].metric("Rata-rata per bulan", fmt_int(total / jumlah_bulan))
    if punya_ul:
        aktif_ul = df.loc[df["ul"] != cleaning.UL_TIDAK_TERPETAKAN, "ul"].nunique()
        kartu[2].metric("UL aktif", fmt_int(aktif_ul))
    else:
        kartu[2].metric("Cabang aktif", fmt_int(df["cabang"].nunique()))
    if cfg["nilai2_label"]:
        jumlah2 = df["nilai2"].sum()
        kartu[3].metric(cfg["nilai2_label"], fmt_rp(jumlah2) if cfg["nilai2_rupiah"] else fmt_int(jumlah2))
    elif punya_ul:
        kartu[3].metric("Cabang aktif", fmt_int(df["cabang"].nunique()))
    else:
        kartu[3].metric("Kanwil aktif", fmt_int(df["kanwil"].nunique()))

    st.write("")
    kiri, kanan = st.columns([5, 6], gap="large")

    # --- 2. Top 10
    with kiri:
        pilihan = list(DIMENSI) if punya_ul else ["Cabang pairing", "Kanwil pairing"]
        dim = st.radio("Top 10 berdasarkan", pilihan, horizontal=True, key=f"dim_{kunci}")
        kolom = DIMENSI[dim]
        sumber = df
        if kolom == "ul":
            sumber = df[df["ul"] != cleaning.UL_TIDAK_TERPETAKAN]
        ranking = tabel_ranking(sumber, kolom)
        st.altair_chart(grafik_top10(ranking.head(10)), width="stretch")
        if kolom == "ul" and len(sumber) < len(df):
            st.caption(f"{fmt_int(df['nilai'].sum() - sumber['nilai'].sum())} {cfg['satuan']} tanpa UL tidak ikut ranking.")

    # --- 3. Grafik per bulan
    with kanan:
        st.markdown(f"**Jumlah {cfg['satuan']} per bulan**")
        chart, ada_nol = grafik_bulanan(df)
        st.altair_chart(chart, width="stretch")
        if ada_nol:
            st.caption("Ada bulan bernilai 0 pada rentang ini. Cek dulu apakah memang tidak ada data di sumber.")

    with st.expander("Lihat tabel lengkap dan unduh"):
        ranking_tampil = ranking.rename(columns={"nama": dim, "jumlah": f"jumlah_{cfg['satuan']}"})
        st.dataframe(ranking_tampil, width="stretch", hide_index=True)
        st.download_button(
            "Unduh CSV",
            ranking_tampil.to_csv(index=False).encode("utf-8"),
            file_name=f"ranking_{cfg['judul'].lower().replace(' ', '_').replace('/', '-')}.csv",
            mime="text/csv",
            key=f"unduh_{kunci}",
        )


def tampilkan_catatan():
    st.subheader("Catatan data dan legenda")
    st.markdown(
        """
**Aturan pembersihan yang dipakai sistem**
- Nama UL dirapikan: awalan "Unit Layanan BPJS Ketenagakerjaan" jadi "UL", nama rusak di sheet 0.1 diperbaiki, KABUPATEN jadi KAB.
- Nama UL yang ditulis beda dipadankan lewat tabel alias (`UL_ALIAS` di `cleaning.py`).
- Master UL = gabungan semua sumber (sheet 1, 0.1, 2.0, 2). Koordinat peta diambil dari sheet 0.1.
- Kanwil dan cabang memakai **kode** (IDKW/IDKC) lalu diberi nama dari legenda di bawah.
- Tanggal kegiatan yang rusak (mis. 0000-00-00, tahun ngawur, masa depan) diganti dengan Tgl Rekam supaya barisnya tidak hilang dari total.
- Sheet 0.2 tidak dipakai (sudah tercakup di sheet 4).
- Hanya kolom yang dibutuhkan yang dibaca. Kolom data pribadi (nama, NIK, HP, email) tidak pernah dibaca atau disimpan.
        """
    )
    baris = []
    for k, cfg in cleaning.SHEETS.items():
        q = catatan["sheet"][k]
        baris.append(
            {
                "Sheet": f"{k}. {cfg['judul']}",
                "Baris sumber": q["baris_dipakai"],
                "Tanggal diganti Tgl Rekam": q["tanggal_diganti_tgl_rekam"],
                "Tanggal tetap tidak valid": q["tanggal_tidak_valid"],
                "% baris dengan UL terpetakan": q["persen_ul_terpetakan"],
                "Rentang tanggal": f"{q['tgl_min']} s/d {q['tgl_maks']}",
                "Catatan": q.get("catatan", ""),
            }
        )
    st.markdown("**Kualitas data per sheet**")
    st.dataframe(pd.DataFrame(baris), width="stretch", hide_index=True)
    st.caption(f"{catatan['ul_dengan_koordinat']} dari {catatan['jumlah_ul_master']} UL punya koordinat untuk peta.")

    ul_baru = catatan.get("ul_tidak_ada_di_sheet_0_1", [])
    if ul_baru:
        st.markdown(f"**{len(ul_baru)} UL ada di data transaksi tetapi tidak ada di sheet 0.1 (master lama)**")
        st.write(", ".join(ul_baru))

    kolom_a, kolom_b = st.columns(2)
    with kolom_a:
        st.markdown("**Legenda kode kanwil**")
        st.dataframe(
            pd.DataFrame({"kode": list(cleaning.KANWIL_NAMA), "kanwil": list(cleaning.KANWIL_NAMA.values())}),
            hide_index=True, width="stretch",
        )
    with kolom_b:
        st.markdown("**Legenda kode cabang**")
        st.dataframe(legenda_cabang.rename(columns={"kode_cabang": "kode", "nama_cabang": "cabang"}),
                     hide_index=True, width="stretch")
        st.caption("Kode cabang yang belum ada namanya ditampilkan sebagai 'kode XXX' di dashboard.")


# ---------------------------------------------------------------------------
# Tampilkan halaman yang dipilih
# ---------------------------------------------------------------------------
if halaman == "Home":
    tampilkan_home()
elif pilihan_data == PILIHAN_CATATAN:
    tampilkan_catatan()
else:
    tampilkan_sheet(judul_ke_kunci[pilihan_data])
