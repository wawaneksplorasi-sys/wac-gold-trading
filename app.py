# -*- coding: utf-8 -*-
# =====================================================================
#   APLIKASI MANAJEMEN USAHA EMAS  —  SINGLE ENTRY SYSTEM
#   -----------------------------------------------------------------
#   Menu user reguler : Dashboard, Keuangan, Rekening, Stok Emas,
#                       Tracking Admin
#   Menu khusus ADMIN (tambahan) : Input Data, Pengaturan
#                       - Input Data  = sekali simpan -> semua menu
#                                       otomatis ter-update
#                       - Pengaturan = profil, ganti password,
#                                       tambah user baru
#
#   Database : Google Sheets (tab: Transaksi, Rekening, Users —
#              dibuat OTOMATIS saat aplikasi pertama kali dijalankan)
#   Login pertama : admin / admin123  dan  user1 / user123
# =====================================================================

import os
import io
import re
import uuid
import base64
import hashlib
import datetime as dt
from datetime import timedelta, timezone

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from PIL import Image, ImageOps

st.set_page_config(page_title="Manajemen Usaha Emas",
                   page_icon="🪙", layout="wide")


# =====================================================================
#  KONFIGURASI — SATU-SATUNYA BAGIAN YANG PERLU ANDA UBAH
#  Ganti dengan ID Google Sheets Anda.
#  Contoh: https://docs.google.com/spreadsheets/d/1AbCdEfGhIjK/edit
#  ID-nya = 1AbCdEfGhIjK
# =====================================================================
SPREADSHEET_ID = "1FYaxiDSoJQe_dwiWOY2YIkBISP_NugYKbK6o4W2wv6E"

FILE_KUNCI    = "service_account.json"     # file kunci Google Cloud
SALT_PASSWORD = "aplikasi-emas-2025"       # jangan diubah

SCOPES    = ["https://www.googleapis.com/auth/spreadsheets"]
WIB       = timezone(timedelta(hours=7))
NAMA_HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]

JENIS_MASUK  = ["Jual Emas", "Pemasukan Lain"]
JENIS_KELUAR = ["Beli Emas", "Pengeluaran Operasional"]
JENIS_SEMUA  = ["Beli Emas", "Jual Emas", "Pemasukan Lain",
                "Pengeluaran Operasional", "Transfer Antar Rekening"]

KOLOM_TRANSAKSI = ["ID", "Tanggal", "Hari", "Jam", "Jenis", "Berat_Gram",
                   "Harga_Per_Gram", "Nominal", "Rekening_Utama",
                   "Rekening_Tujuan", "Arah", "Lokasi_Nama", "Latitude",
                   "Longitude", "Aktivitas", "Admin", "Waktu_Input", "Foto"]
KOLOM_REKENING = ["Nama_Rekening", "Jenis", "No_Rekening", "Atas_Nama",
                  "Saldo_Awal", "Keterangan"]
KOLOM_USERS    = ["Username", "PasswordHash", "Nama", "Role", "Email"]

MENU_DASHBOARD = "📊 Dashboard"
MENU_INPUT     = "📝 Input Data"
MENU_KEUANGAN  = "💰 Keuangan"
MENU_REKENING  = "🏦 Rekening"
MENU_STOK      = "🪙 Stok Emas"
MENU_TRACKING  = "📍 Tracking Admin"
MENU_SETTING   = "⚙️ Pengaturan"


# =====================================================================
#  FUNGSI BANTU
# =====================================================================
def sekarang_wib():
    return dt.datetime.now(WIB)


def format_rupiah(nilai):
    try:
        nilai = float(nilai)
    except (TypeError, ValueError):
        return "-"
    tanda = "-" if nilai < 0 else ""
    return f"{tanda}Rp {abs(nilai):,.0f}".replace(",", ".")


def format_gram(nilai):
    try:
        nilai = float(nilai)
    except (TypeError, ValueError):
        return "-"
    s = f"{nilai:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} gram"


def hash_password(pw):
    return hashlib.sha256((SALT_PASSWORD + str(pw)).encode("utf-8")).hexdigest()


def parse_koordinat(teks):
    """Mengubah teks seperti '-6.2146, 106.8451' menjadi (lat, lon)."""
    if not teks:
        return None
    s = str(teks).strip()
    m = re.match(r"^\s*(-?\d+,\d+)\s*[,;]?\s*(-?\d+,\d+)\s*$", s)
    if m:  # format koma desimal, mis. -6,2146, 106,8451
        return float(m.group(1).replace(",", ".")), float(m.group(2).replace(",", "."))
    angka = re.findall(r"-?\d+(?:\.\d+)?", s)
    if len(angka) >= 2:
        try:
            return float(angka[0]), float(angka[1])
        except ValueError:
            return None
    return None


def kompres_foto(byte_foto):
    """Memperkecil foto agar muat disimpan di Google Sheets (base64)."""
    try:
        img = ImageOps.exif_transpose(Image.open(io.BytesIO(byte_foto)))
        img = img.convert("RGB")
        hasil = ""
        for sisi, kualitas in [(480, 60), (380, 55), (300, 50), (220, 45)]:
            salinan = img.copy()
            salinan.thumbnail((sisi, sisi))
            buf = io.BytesIO()
            salinan.save(buf, format="JPEG", quality=kualitas)
            hasil = base64.b64encode(buf.getvalue()).decode("utf-8")
            if len(hasil) <= 45000:
                return hasil
        return hasil
    except Exception:
        return ""


def deteksi_gps():
    """Mencoba deteksi GPS otomatis dari perangkat. Bila gagal, kembalikan None."""
    try:
        from streamlit_geolocation import write_streamlit_geolocation
        hasil = write_streamlit_geolocation()
        if isinstance(hasil, dict) and hasil.get("latitude") not in (None, ""):
            return float(hasil["latitude"]), float(hasil["longitude"])
    except Exception:
        pass
    return None


# =====================================================================
#  KONEKSI & DATA (GOOGLE SHEETS)
# =====================================================================
@st.cache_resource(show_spinner=False)
def buat_koneksi():
    # 1) Kumpulkan semua file yang mengandung "json" pada namanya
    kandidat = []
    try:
        for nama_file in os.listdir("."):
            if "json" in nama_file.lower():
                kandidat.append(nama_file)
    except Exception:
        pass

    # 2) Pakai file yang benar-benar berisi kunci Google
    #    (ditandai adanya 'client_email' dan 'private_key' di isinya)
    for nama in kandidat:
        try:
            with open(nama, "r", encoding="utf-8") as f:
                isi = f.read()
            if "client_email" in isi and "private_key" in isi:
                creds = Credentials.from_service_account_file(nama, scopes=SCOPES)
                return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)
        except Exception:
            continue

    # 3) Bila tidak ada, coba Streamlit secrets (dipakai nanti saat hosting)
    try:
        info = dict(st.secrets["service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)
    except Exception:
        pass

    # 4) Pesan error yang lebih jelas (menunjukkan isi folder Anda)
    folder = os.getcwd()
    if kandidat:
        raise RuntimeError(
            f"Ada file JSON di {folder}: {', '.join(kandidat)} — tetapi BUKAN "
            "file kunci Google yang valid (tidak berisi 'client_email'). "
            "Unduh ulang: Google Cloud Console → IAM & Admin → Service Accounts "
            "→ klik akun → tab KEYS → ADD KEY → Create new key → JSON.")
    raise RuntimeError(
        f"File kunci Google (JSON) TIDAK ADA di folder: {folder}. Salin file "
        "JSON hasil unduhan dari Google Cloud ke folder tersebut (nama file "
        "apa pun boleh, asal berekstensi .json), lalu jalankan ulang aplikasi.")


def pastikan_setup():
    """Membuat tab Users, Rekening, Transaksi bila belum ada (sekali saja).
    VERSI RINGAN: hanya membaca baris judul + 1 kolom per tab — tidak lagi
    mengunduh seluruh isi sheet (termasuk foto), sehingga tetap cepat
    walaupun data sudah banyak."""
    sh = buat_koneksi()
    judul = [w.title for w in sh.worksheets()]

    def buat_tab(nama, kolom, jml_baris, jml_kolom):
        if nama not in judul:
            w = sh.add_worksheet(nama, jml_baris, jml_kolom)
            w.append_row(kolom)
            return w
        w = sh.worksheet(nama)
        header = [str(x).strip() for x in w.row_values(1)]
        if not header:
            w.append_row(kolom)
            return w
        if header[:len(kolom)] != kolom:
            # migrasi otomatis: tab Users versi lama (tanpa kolom Email)
            if nama == "Users" and header[:4] == kolom[:4]:
                w.update_acell("E1", "Email")
                return w
            raise RuntimeError(
                f"Tab '{nama}' sudah ada tetapi format kolomnya tidak sesuai. "
                "Gunakan Google Sheets BARU yang kosong khusus aplikasi ini.")
        return w

    wu = buat_tab("Users", KOLOM_USERS, 100, 10)
    wr = buat_tab("Rekening", KOLOM_REKENING, 100, 10)
    buat_tab("Transaksi", KOLOM_TRANSAKSI, 3000, 20)

    # cek isi Users & Rekening cukup lewat kolom pertama (ringan)
    try:
        ada_user = len(wu.col_values(1)) > 1
    except Exception:
        ada_user = True
    if not ada_user:
        wu.append_row(["admin", hash_password("admin123"),
                       "Administrator", "admin", ""])
        wu.append_row(["user1", hash_password("user123"),
                       "User Demo", "reguler", ""])

    try:
        ada_rek = len(wr.col_values(1)) > 1
    except Exception:
        ada_rek = True
    if not ada_rek:
        wr.append_row(["Kas Toko", "Kas Tunai", "", "", 0,
                       "Contoh rekening — silakan ganti"])


@st.cache_resource(show_spinner=False)
def ambil_sheet(nama_tab):
    """Objek tab Google Sheets versi cache — menghemat 1 panggilan API
    setiap kali membaca data (semua halaman jadi lebih cepat)."""
    sh = buat_koneksi()
    return sh.worksheet(nama_tab)


@st.cache_data(ttl=60, show_spinner=False)
def ambil_transaksi(dengan_foto=False):
    sh = buat_koneksi()
    akhir = "R" if dengan_foto else "Q"   # kolom R = Foto (diambil hanya jika perlu)
    resp = sh.values_get(f"Transaksi!A2:{akhir}")
    baris = resp.get("values", [])
    n = 18 if dengan_foto else 17
    rapi = []
    for r in baris:
        r = list(r)
        if len(r) < n:
            r = r + [""] * (n - len(r))
        rapi.append(r[:n])
    df = pd.DataFrame(rapi, columns=KOLOM_TRANSAKSI[:n])
    if df.empty:
        return df
    for c in ("Berat_Gram", "Harga_Per_Gram", "Nominal", "Latitude", "Longitude"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    for c in ("ID", "Tanggal", "Hari", "Jam", "Jenis", "Rekening_Utama",
              "Rekening_Tujuan", "Arah", "Lokasi_Nama", "Aktivitas",
              "Admin", "Waktu_Input"):
        df[c] = df[c].astype(str).str.strip()
    df["_baris"] = list(range(2, 2 + len(df)))
    df["_urut"] = df["Waktu_Input"].astype(str)
    df = df.sort_values(["Tanggal", "Jam", "_urut"]).drop(columns="_urut")
    return df.reset_index(drop=True)


@st.cache_data(ttl=60, show_spinner=False)
def ambil_rekening():
    baris = ambil_sheet("Rekening").get_all_records()
    df = pd.DataFrame(baris)
    if df.empty:
        df = pd.DataFrame(columns=KOLOM_REKENING)
    for c in KOLOM_REKENING:
        if c not in df.columns:
            df[c] = ""
    df["Saldo_Awal"] = pd.to_numeric(df["Saldo_Awal"], errors="coerce").fillna(0)
    df["Nama_Rekening"] = df["Nama_Rekening"].astype(str).str.strip()
    return df


@st.cache_data(ttl=60, show_spinner=False)
def ambil_users():
    w = ambil_sheet("Users")
    values = w.get_all_values()
    if not values:
        return pd.DataFrame(columns=KOLOM_USERS)
    header = [str(x).strip() for x in values[0]]
    # jaring pengaman: pastikan kolom Email ada di baris judul
    if len(header) < len(KOLOM_USERS):
        for i in range(len(header), len(KOLOM_USERS)):
            w.update_acell(f"{chr(65 + i)}1", KOLOM_USERS[i])
        header = header + KOLOM_USERS[len(header):]
    baris = []
    for r in values[1:]:
        r = list(r)
        if not any(str(x).strip() for x in r):
            continue  # lewati baris kosong
        if len(r) < len(KOLOM_USERS):
            r = r + [""] * (len(KOLOM_USERS) - len(r))
        baris.append(r[:len(KOLOM_USERS)])
    df = pd.DataFrame(baris, columns=KOLOM_USERS)
    for c in KOLOM_USERS:
        df[c] = df[c].astype(str).str.strip()
    return df


@st.cache_data(ttl=60, show_spinner=False)
def ambil_foto_terbatas(baris_awal, baris_akhir):
    """Unduh kolom Foto HANYA untuk rentang baris yang ditampilkan
    (bukan seluruh database) — halaman Tracking jadi jauh lebih ringan."""
    sh = buat_koneksi()
    resp = sh.values_get(f"Transaksi!R{baris_awal}:R{baris_akhir}")
    hasil = {}
    for i, r in enumerate(resp.get("values", [])):
        if r and str(r[0]).strip():
            hasil[baris_awal + i] = str(r[0])
    return hasil


def tambah_baris(nama_tab, baris):
    sh = buat_koneksi()
    sh.worksheet(nama_tab).append_row(baris)
    st.cache_data.clear()


def cari_baris_user(username):
    """Nomor baris seorang user di sheet Users (atau None)."""
    sh = buat_koneksi()
    wu = sh.worksheet("Users")
    semua = wu.get_all_values()
    for i, r in enumerate(semua):
        if i > 0 and str(r[0]).strip().lower() == str(username).strip().lower():
            return i + 1
    return None


# =====================================================================
#  PERHITUNGAN OTOMATIS (inti Single Entry System)
# =====================================================================
def hitung_saldo(df_trx, df_rek):
    saldo = {str(r["Nama_Rekening"]).strip(): float(r["Saldo_Awal"] or 0)
             for _, r in df_rek.iterrows()}
    for _, t in df_trx.iterrows():
        j = t["Jenis"]
        n = float(t["Nominal"] or 0)
        ru = str(t["Rekening_Utama"]).strip()
        rt = str(t["Rekening_Tujuan"]).strip()
        if j in JENIS_MASUK:
            saldo[ru] = saldo.get(ru, 0) + n
        elif j in JENIS_KELUAR:
            saldo[ru] = saldo.get(ru, 0) - n
        elif j == "Transfer Antar Rekening":
            saldo[ru] = saldo.get(ru, 0) - n
            if rt:
                saldo[rt] = saldo.get(rt, 0) + n
    return saldo


def bangun_buku(df_trx):
    """Buku kas: Debit = uang masuk, Kredit = uang keluar."""
    data = []
    for _, t in df_trx.iterrows():
        j = t["Jenis"]
        n = float(t["Nominal"] or 0)
        if j in JENIS_MASUK:
            debit, kredit, rek = n, 0.0, t["Rekening_Utama"]
        elif j in JENIS_KELUAR:
            debit, kredit, rek = 0.0, n, t["Rekening_Utama"]
        else:  # transfer
            debit, kredit = n, n
            rek = f"{t['Rekening_Utama']} → {t['Rekening_Tujuan']}"
        ket = str(t["Aktivitas"] or "")
        if t["Lokasi_Nama"]:
            ket = (ket + " | " if ket else "") + f"📍{t['Lokasi_Nama']}"
        data.append({"Tanggal": t["Tanggal"], "Hari": t["Hari"], "Jam": t["Jam"],
                     "Jenis": j, "Rekening": rek,
                     "Debit (Masuk)": debit, "Kredit (Keluar)": kredit,
                     "Keterangan": ket, "Admin": t["Admin"], "ID": t["ID"]})
    return pd.DataFrame(data)


def mutasi_rekening(df_trx, df_rek, akun):
    saldo_awal = 0.0
    for _, r in df_rek.iterrows():
        if str(r["Nama_Rekening"]).strip() == akun:
            saldo_awal = float(r["Saldo_Awal"] or 0)
            break
    sub = df_trx[(df_trx["Rekening_Utama"].astype(str).str.strip() == akun) |
                 (df_trx["Rekening_Tujuan"].astype(str).str.strip() == akun)].copy()
    efek = []
    for _, t in sub.iterrows():
        j = t["Jenis"]
        n = float(t["Nominal"] or 0)
        if j == "Transfer Antar Rekening" and str(t["Rekening_Tujuan"]).strip() == akun:
            efek.append(n)
        else:
            efek.append(n if j in JENIS_MASUK else -n)
    sub["Efek"] = efek
    sub["Saldo"] = saldo_awal + sub["Efek"].cumsum()
    sub["Masuk"] = sub["Efek"].clip(lower=0)
    sub["Keluar"] = (-sub["Efek"]).clip(lower=0)
    hasil = sub[["Tanggal", "Hari", "Jam", "Jenis", "Aktivitas", "Lokasi_Nama",
                 "Masuk", "Keluar", "Saldo", "ID"]].copy()
    return hasil, saldo_awal


def data_emas(df_trx):
    e = df_trx[df_trx["Jenis"].isin(["Beli Emas", "Jual Emas"])].copy()
    if e.empty:
        return e
    e["Arah_Gram"] = e["Berat_Gram"].where(e["Jenis"].eq("Beli Emas"),
                                           -e["Berat_Gram"])
    e["Sisa_Stok"] = e["Arah_Gram"].cumsum()
    return e


# =====================================================================
#  HALAMAN LOGIN
# =====================================================================
def halaman_login():
    k1, k2, k3 = st.columns([1, 1.3, 1])
    with k2:
        st.markdown("<h2 style='text-align:center'>🪙 Manajemen Usaha Emas</h2>",
                    unsafe_allow_html=True)
        st.markdown("<p style='text-align:center;color:gray'>Single Entry System · "
                    "Data tersimpan di Google Sheets</p>", unsafe_allow_html=True)
        with st.form("form_login", border=True):
            user = st.text_input("Username")
            pw = st.text_input("Password", type="password")
            masuk = st.form_submit_button("🔐 Masuk", type="primary",
                                          use_container_width=True)
        if masuk:
            user = str(user).strip().lower()
            if not user or not pw:
                st.error("Isi username dan password.")
            else:
                df_u = ambil_users()
                cocok = df_u[df_u["Username"].astype(str).str.strip().str.lower() == user]
                if (not cocok.empty and
                        str(cocok.iloc[0]["PasswordHash"]).strip() == hash_password(pw)):
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = str(cocok.iloc[0]["Username"]).strip()
                    st.session_state["nama"] = str(cocok.iloc[0]["Nama"] or
                                                  cocok.iloc[0]["Username"]).strip()
                    st.session_state["role"] = str(cocok.iloc[0]["Role"]).strip().lower()
                    st.rerun()
                else:
                    st.error("Username atau password salah.")
        st.caption("🔐 Login pertama: **admin / admin123** — segera ganti password "
                   "di menu ⚙️ Pengaturan.")


# =====================================================================
#  MENU (SIDEBAR)
# =====================================================================
def menu_utama():
    peran = st.session_state.get("role", "user")
    nama = st.session_state.get("nama", "")
    with st.sidebar:
        st.markdown("## 🪙 Usaha Emas")
        st.caption(f"👤 {nama} · "
                   f"{'🛡️ Admin' if peran == 'admin' else '👥 Reguler'}")
        daftar = [MENU_DASHBOARD, MENU_KEUANGAN, MENU_REKENING,
                  MENU_STOK, MENU_TRACKING]
        if peran == "admin":
            daftar.insert(1, MENU_INPUT)
            daftar.append(MENU_SETTING)
        pilihan = st.radio("MENU", daftar, label_visibility="collapsed")
        st.markdown("---")
        if st.button("🔄 Muat Ulang Data (terbaru)", use_container_width=True,
                     help="Ambil data terbaru dari Google Sheets sekarang juga, "
                          "tanpa menunggu cache 60 detik. Gunakan bila admin "
                          "lain baru saja menginput dari perangkat lain."):
            st.cache_data.clear()
            st.rerun()
        if st.button("🚪 Keluar", use_container_width=True):
            for k in ("logged_in", "username", "nama", "role", "setup_ok"):
                st.session_state.pop(k, None)
            st.rerun()
        st.caption("💾 Database: Google Sheets (cache 60 detik)")

    if pilihan == MENU_DASHBOARD:
        halaman_dashboard()
    elif pilihan == MENU_INPUT and peran == "admin":
        halaman_input()
    elif pilihan == MENU_KEUANGAN:
        halaman_keuangan()
    elif pilihan == MENU_REKENING:
        halaman_rekening()
    elif pilihan == MENU_STOK:
        halaman_stok()
    elif pilihan == MENU_TRACKING:
        halaman_tracking()
    elif pilihan == MENU_SETTING and peran == "admin":
        halaman_setting()
    else:
        halaman_dashboard()


# =====================================================================
#  HALAMAN: DASHBOARD
# =====================================================================
def grafik_batang_laba(label_bulan, nilai):
    """Grafik batang laba/rugi per bulan: hijau = laba (+),
    merah = rugi (-). Bulan bernilai None (belum ada data) tidak
    digambar batangnya. Fungsi ini MANDIRI (tidak bergantung pada
    fungsi lain di dalam halaman dashboard)."""
    import altair as alt

    def _tanda_rp(v):
        """Format Rupiah bertanda: + untuk laba, - untuk rugi."""
        if v is None:
            return "-"
        if v > 0:
            return "+" + format_rupiah(v)
        if v < 0:
            return "-" + format_rupiah(abs(v))
        return format_rupiah(0)

    df_c = pd.DataFrame({
        "Bulan": list(label_bulan),
        "Nilai": [None if v is None else float(v) for v in nilai]})
    df_c["Warna"] = df_c["Nilai"].map(
        lambda v: None if v is None or v == 0 else
        ("Laba" if v > 0 else "Rugi"))
    df_c["Tampil"] = df_c["Nilai"].map(_tanda_rp)
    skala = alt.Scale(domain=["Laba", "Rugi"],
                      range=["#2ca02c", "#d62728"])   # hijau, merah
    chart = alt.Chart(df_c).mark_bar().encode(
        x=alt.X("Bulan:N", sort=list(label_bulan), title=None),
        y=alt.Y("Nilai:Q", title="Laba/Rugi (Rp)"),
        color=alt.Color("Warna:N", scale=skala, legend=alt.Legend(
            title=None, orient="top")),
        tooltip=[alt.Tooltip("Bulan:N", title="Bulan"),
                 alt.Tooltip("Tampil:N", title="Laba/Rugi")])
    st.altair_chart(chart, use_container_width=True)

def grafik_batang_bulan(label_bulan, nilai, judul, warna="#f0b429"):
    """Grafik batang per bulan dengan urutan bulan yang BENAR
    (sesuai kalender / bulan data pertama s.d. Desember), bukan abjad.
    Tooltip menampilkan nilai dalam format Rupiah."""
    import altair as alt
    df_c = pd.DataFrame({"Bulan": list(label_bulan),
                         "Nilai": [float(v) for v in nilai]})
    df_c["Tampil"] = [format_rupiah(v) for v in df_c["Nilai"]]
    chart = alt.Chart(df_c).mark_bar(color=warna).encode(
        x=alt.X("Bulan:N", sort=list(label_bulan), title=None),
        y=alt.Y("Nilai:Q", title=judul, axis=alt.Axis(format="~s")),
        tooltip=[alt.Tooltip("Bulan:N", title="Bulan"),
                 alt.Tooltip("Tampil:N", title=judul)])
    st.altair_chart(chart, use_container_width=True)
    
def halaman_dashboard():
    st.header("📊 Dashboard")

    BULAN_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
                "Agustus", "September", "Oktober", "November", "Desember"]

    df_trx = ambil_transaksi()
    df_rek = ambil_rekening()
    now = sekarang_wib()
    ada_trx = not df_trx.empty

    # pemformat data kosong -> tanda "-"
    def rp(v):
        return "-" if v is None else format_rupiah(v)

    def gr(v):
        return "-" if v is None else format_gram(v)

    # pemformat nilai bertanda: + = profit / bertambah, - = loss / berkurang
    def tanda_rp(v):
        if v is None:
            return "-"
        if v > 0:
            return "+" + format_rupiah(v)
        if v < 0:
            return "-" + format_rupiah(abs(v))
        return format_rupiah(0)

    def tanda_gr(v):
        if v is None:
            return "-"
        if v > 0:
            return "+" + format_gram(v)
        if v < 0:
            return "-" + format_gram(abs(v))
        return format_gram(0)

    if ada_trx:
        df_trx = df_trx.copy()
        df_trx["_t"] = pd.to_datetime(df_trx["Tanggal"], errors="coerce")
        df_tahun = df_trx[df_trx["_t"].dt.year == now.year]
    else:
        df_tahun = pd.DataFrame()
    ada_tahun = not df_tahun.empty

    # jendela bulan tahun berjalan: bulan transaksi pertama s.d. Desember
    bulan_awal = int(df_tahun["_t"].dt.month.min()) if ada_tahun else now.month
    bulan_list = list(range(bulan_awal, 13))
    label_bulan = [BULAN_ID[b - 1] for b in bulan_list]

    # --- agregat bulanan (semua tahun) untuk perhitungan profit/loss ---
    penj_bulan = {}          # total penjualan emas per bulan (Rp)
    emas_all = pd.DataFrame()
    if ada_trx:
        jual_all = df_trx[df_trx["Jenis"].eq("Jual Emas")].copy()
        if not jual_all.empty:
            jual_all["_ym"] = (jual_all["_t"].dt.year * 12
                               + jual_all["_t"].dt.month - 1)
            penj_bulan = {int(k): float(v) for k, v in
                          jual_all.groupby("_ym")["Nominal"].sum().items()}
        emas_all = df_trx[df_trx["Jenis"].isin(["Beli Emas", "Jual Emas"])].copy()
        if not emas_all.empty:
            emas_all["_ym"] = (emas_all["_t"].dt.year * 12
                               + emas_all["_t"].dt.month - 1)
            emas_all["_arah"] = emas_all["Berat_Gram"].where(
                emas_all["Jenis"].eq("Beli Emas"), -emas_all["Berat_Gram"])

    def stok_sampai(ym):
        """Berat emas (gram) yang menjadi stok pada akhir bulan 'ym'."""
        if emas_all.empty or "_ym" not in emas_all.columns:
            return 0.0
        try:
            return float(emas_all.loc[emas_all["_ym"] <= ym, "_arah"].sum())
        except Exception:
            return 0.0

    if not ada_trx:
        st.info("📭 Belum ada data transaksi — Dashboard menampilkan template "
                "dengan tanda ( **-** ). Masuk sebagai **admin**, buka menu "
                "**📝 Input Data** untuk mencatat transaksi pertama — semua "
                "angka, tabel, dan grafik akan terisi otomatis.")

    # -----------------------------------------------------------------
    # 1) REKENING — tabel semua rekening + saldo terakhir + total
    # -----------------------------------------------------------------
    st.subheader("🏦 Rekening")
    saldo = hitung_saldo(df_trx, df_rek)
    terdaftar = set()
    baris_rek = []
    for _, r in df_rek.iterrows():
        nama = str(r["Nama_Rekening"]).strip()
        if not nama:
            continue
        terdaftar.add(nama)
        baris_rek.append({"Nama Rekening": nama,
                          "Jenis": str(r["Jenis"] or "").strip() or "-",
                          "No. Rekening": str(r["No_Rekening"] or "").strip() or "-",
                          "Saldo Terakhir": format_rupiah(saldo.get(nama, 0))})
    for k, v in saldo.items():   # rekening lama yang tidak terdaftar lagi
        if k and k not in terdaftar:
            baris_rek.append({"Nama Rekening": k, "Jenis": "(tidak terdaftar)",
                              "No. Rekening": "-",
                              "Saldo Terakhir": format_rupiah(v)})
    if baris_rek:
        st.dataframe(pd.DataFrame(baris_rek), use_container_width=True,
                     hide_index=True)
        st.markdown("<div style='text-align:right;font-size:1.05rem'>"
                    f"<b>💰 TOTAL SALDO: {format_rupiah(sum(saldo.values()))}"
                    f"</b></div>", unsafe_allow_html=True)
    else:
        st.dataframe(pd.DataFrame([
            {"Nama Rekening": "-", "Jenis": "-", "No. Rekening": "-",
             "Saldo Terakhir": "-"}]),
            use_container_width=True, hide_index=True)
        st.markdown("<div style='text-align:right;font-size:1.05rem'>"
                    "<b>💰 TOTAL SALDO: -</b></div>", unsafe_allow_html=True)

        # -----------------------------------------------------------------
    # 1b) LABA/RUGI CASH (bulan berjalan & tahunan)
    #     Laba/Rugi bulan  = penjualan emas bulan ini − pengeluaran bulan ini
    #     Laba/Rugi tahun  = penjualan emas (1 Jan s.d. kini) −
    #                        pengeluaran (1 Jan s.d. kini)
    #     Pengeluaran = Beli Emas + Pengeluaran Operasional
    #                   (transfer antar rekening tidak dihitung)
    # -----------------------------------------------------------------
    ym_now = now.year * 12 + (now.month - 1)

    kel_bulan = {}          # pengeluaran per bulan (semua tahun)
    if ada_trx:
        kel_all = df_trx[df_trx["Jenis"].isin(JENIS_KELUAR)].copy()
        if not kel_all.empty:
            kel_all["_ym"] = (kel_all["_t"].dt.year * 12
                              + kel_all["_t"].dt.month - 1)
            kel_bulan = {int(k): float(v) for k, v in
                         kel_all.groupby("_ym")["Nominal"].sum().items()}

    if ada_trx:
        # bulan berjalan
        jual_now = penj_bulan.get(ym_now, 0.0)
        kel_now = kel_bulan.get(ym_now, 0.0)
        laba_bulan = jual_now - kel_now
        # tahun berjalan (1 Jan s.d. sekarang)
        d_tahun = df_trx[df_trx["_t"].dt.year == now.year]
        jual_thn = float(d_tahun.loc[d_tahun["Jenis"].eq("Jual Emas"),
                                     "Nominal"].sum())
        kel_thn = float(d_tahun.loc[d_tahun["Jenis"].isin(JENIS_KELUAR),
                                    "Nominal"].sum())
        laba_tahun = jual_thn - kel_thn
    else:
        jual_now = kel_now = laba_bulan = None
        jual_thn = kel_thn = laba_tahun = None

    st.subheader("📊 Laba/Rugi Cash")
    lb1, lb2 = st.columns(2)
    with lb1:
        st.metric(f"💵 Laba/Rugi — {BULAN_ID[now.month - 1]} {now.year}",
                  tanda_rp(laba_bulan),
                  help="Total penjualan emas bulan berjalan dikurangi total "
                       "pengeluaran bulan berjalan. Pengeluaran = beli emas + "
                       "pengeluaran operasional (transfer antar rekening "
                       "tidak dihitung). Positif (+) = laba, negatif (−) = rugi.")
        st.caption("-" if jual_now is None else
                   f"Penjualan: {format_rupiah(jual_now)} · "
                   f"Pengeluaran: {format_rupiah(kel_now)}")
    with lb2:
        st.metric(f"📆 Laba/Rugi — Tahun {now.year} (s.d. sekarang)",
                  tanda_rp(laba_tahun),
                  help="Total penjualan emas sejak 1 Januari sampai hari ini "
                       "dikurangi total pengeluaran pada periode yang sama.")
        st.caption("-" if jual_thn is None else
                   f"Penjualan: {format_rupiah(jual_thn)} · "
                   f"Pengeluaran: {format_rupiah(kel_thn)}")

    # ----- tabel laba/rugi bulanan (tahun berjalan) -----
    st.markdown(f"**📋 Tabel Laba/Rugi Bulanan — Tahun {now.year}**")
    baris_pl = []
    for b in bulan_list:
        if b == now.month:
            label = f"{BULAN_ID[b - 1]} (berjalan)"
        elif b > now.month:
            label = f"{BULAN_ID[b - 1]} (belum ada data)"
        else:
            label = BULAN_ID[b - 1]
        if (not ada_trx) or (b > now.month):
            baris_pl.append({"Bulan": label,
                             "Penjualan (Rp)": "-",
                             "Pengeluaran (Rp)": "-",
                             "Laba/Rugi (Rp)": "-"})
            continue
        ym = now.year * 12 + (b - 1)
        j_b = penj_bulan.get(ym, 0.0)
        k_b = kel_bulan.get(ym, 0.0)
        baris_pl.append({
            "Bulan": label,
            "Penjualan (Rp)": format_rupiah(j_b),
            "Pengeluaran (Rp)": format_rupiah(k_b),
            "Laba/Rugi (Rp)": tanda_rp(j_b - k_b)})
    st.dataframe(pd.DataFrame(baris_pl), use_container_width=True,
                 hide_index=True)
    st.caption("Laba/Rugi = penjualan emas − pengeluaran (beli emas + "
               "pengeluaran operasional; transfer antar rekening tidak "
               "dihitung). Bulan setelah bulan berjalan ditampilkan (-) "
               "karena belum ada data.")

    st.markdown("---")

    # -----------------------------------------------------------------
    # 2) STOK EMAS — bulan berjalan + sisa stok total (gram)
    # -----------------------------------------------------------------
    st.subheader(f"🪙 Stok Emas — {BULAN_ID[now.month - 1]} {now.year}")
    if ada_trx:
        emas = data_emas(df_trx)
        stok_total = float(emas["Sisa_Stok"].iloc[-1]) if not emas.empty else None
        msk_bulan = ((df_trx["_t"].dt.year == now.year) &
                     (df_trx["_t"].dt.month == now.month))
        beli_bulan = float(df_trx.loc[msk_bulan & df_trx["Jenis"].eq("Beli Emas"),
                                      "Berat_Gram"].sum())
        jual_bulan = float(df_trx.loc[msk_bulan & df_trx["Jenis"].eq("Jual Emas"),
                                      "Berat_Gram"].sum())
        n_jual = int((msk_bulan & df_trx["Jenis"].eq("Jual Emas")).sum())
    else:
        stok_total = beli_bulan = jual_bulan = None
        n_jual = None

    m = st.columns(3)
    m[0].metric("Emas Dibeli Bulan Ini", gr(beli_bulan))
    m[1].metric("Emas Dijual Bulan Ini", gr(jual_bulan))
    m[2].metric("Total Stok Emas Tersisa", gr(stok_total))

    st.markdown("---")

    # -----------------------------------------------------------------
    # 3) KUMPULAN GRAFIK — tahun berjalan
    #    OMZET = akumulasi PENJUALAN EMAS saja (Jual Emas)
    # -----------------------------------------------------------------
    st.subheader(f"📈 Kumpulan Grafik — Tahun {now.year}")
    if ada_tahun:
        st.caption(f"Perbandingan bulanan dimulai dari {BULAN_ID[bulan_awal - 1]} "
                   f"(bulan transaksi pertama tahun ini) sampai Desember "
                   f"{now.year}. Omzet = akumulasi penjualan emas.")
    else:
        st.caption(f"Belum ada transaksi tahun {now.year} — grafik menampilkan "
                   f"template bulan {BULAN_ID[bulan_awal - 1]} s.d. Desember "
                   f"{now.year}. Omzet = akumulasi penjualan emas.")

    def sum_per_bulan(dframe):
        if dframe is None or dframe.empty:
            return {}
        g = dframe.groupby(dframe["_t"].dt.month)["Nominal"].sum()
        return {int(k): float(v) for k, v in g.items()}

    if ada_tahun:
        omzet_map = sum_per_bulan(df_tahun[df_tahun["Jenis"].eq("Jual Emas")])
        keluar_map = sum_per_bulan(df_tahun[df_tahun["Jenis"].isin(JENIS_KELUAR)])
    else:
        omzet_map, keluar_map = {}, {}
    ser_omzet = [omzet_map.get(b, 0.0) for b in bulan_list]
    ser_keluar = [keluar_map.get(b, 0.0) for b in bulan_list]
    ser_laba = ([omzet_map.get(b, 0.0) - keluar_map.get(b, 0.0)
                 if ada_tahun and b <= now.month else None
                 for b in bulan_list])

    def v_bulan_ini(ser):
        if not ada_trx:
            return None
        i = now.month - bulan_awal
        return float(ser[i]) if 0 <= i < len(ser) else 0.0

    def v_bulan_lalu(ser):
        if not ada_trx:
            return None
        i = now.month - 1 - bulan_awal
        return float(ser[i]) if 0 <= i < len(ser) else None

    def delta_teks(v_ini, v_lalu):
        if v_ini is None or v_lalu is None or v_lalu == 0:
            return None
        selisih = v_ini - v_lalu
        return ("+" if selisih >= 0 else "-") + format_rupiah(abs(selisih))

    tab1, tab2, tab3, tab4 = st.tabs(
        ["📈 Omzet Penjualan Emas",
         "💸 Pengeluaran",
         "🥧 Porsi Pengeluaran (Bulan Ini)",
         "⚖️ Laba/Rugi Bulanan"])

    with tab1:
        v_ini, v_lalu = v_bulan_ini(ser_omzet), v_bulan_lalu(ser_omzet)
        t1, t2, t3, t4 = st.columns(4)
        t1.metric("Omzet Penjualan Bulan Ini", rp(v_ini),
                  delta=delta_teks(v_ini, v_lalu))
        t2.metric("Transaksi Penjualan Bulan Ini",
                  "-" if n_jual is None else f"{n_jual} transaksi",
                  help="Jumlah transaksi 'Jual Emas' pada bulan berjalan. "
                       "Nominal seluruhnya terakumulasi di Omzet.")
        t3.metric("Total Tahun Berjalan",
                  rp(sum(ser_omzet) if ada_tahun else None))
        if ser_omzet and max(ser_omzet) > 0:
            i_maks = ser_omzet.index(max(ser_omzet))
            with t4:
                st.metric("Bulan Tertinggi", label_bulan[i_maks])
                st.caption(format_rupiah(max(ser_omzet)))
        else:
            t4.metric("Bulan Tertinggi", "-")
        grafik_batang_bulan(label_bulan, ser_omzet,
                            "Omzet Penjualan Emas (Rp)", "#f0b429")
        with st.expander("🔢 Lihat angka per bulan"):
            nilai = [format_rupiah(v) if ada_tahun else "-" for v in ser_omzet]
            st.dataframe(pd.DataFrame({"Bulan": label_bulan,
                                       "Omzet Penjualan (Rp)": nilai}),
                         use_container_width=True, hide_index=True)

    with tab2:
        v_ini, v_lalu = v_bulan_ini(ser_keluar), v_bulan_lalu(ser_keluar)
        t1, t2, t3 = st.columns(3)
        t1.metric("Pengeluaran Bulan Ini", rp(v_ini),
                  delta=delta_teks(v_ini, v_lalu), delta_color="inverse")
        t2.metric("Total Tahun Berjalan",
                  rp(sum(ser_keluar) if ada_tahun else None))
        if ser_keluar and max(ser_keluar) > 0:
            i_maks = ser_keluar.index(max(ser_keluar))
            with t3:
                st.metric("Bulan Tertinggi", label_bulan[i_maks])
                st.caption(format_rupiah(max(ser_keluar)))
        else:
            t3.metric("Bulan Tertinggi", "-")
        grafik_batang_bulan(label_bulan, ser_keluar,
                            "Pengeluaran (Rp)", "#d62728")
        with st.expander("🔢 Lihat angka per bulan"):
            nilai = [format_rupiah(v) if ada_tahun else "-" for v in ser_keluar]
            st.dataframe(pd.DataFrame({"Bulan": label_bulan,
                                       "Pengeluaran (Rp)": nilai}),
                         use_container_width=True, hide_index=True)

    with tab3:
        st.caption(f"Data pengeluaran bulan berjalan: "
                   f"{BULAN_ID[now.month - 1]} {now.year}.")
        if ada_trx:
            d_p = df_trx[(df_trx["_t"].dt.year == now.year) &
                         (df_trx["_t"].dt.month == now.month)]
            d_p = d_p[d_p["Jenis"].isin(JENIS_KELUAR)]
        else:
            d_p = pd.DataFrame()
        if d_p.empty:
            grup = pd.Series(dtype="float64")
        else:
            grup = d_p.groupby("Jenis")["Nominal"].sum().sort_values(ascending=False)
        if grup.empty or float(grup.sum()) <= 0:
            st.info(f"Belum ada data pengeluaran pada {BULAN_ID[now.month - 1]} "
                    f"{now.year}. Pie chart & tabel rincian akan tampil otomatis "
                    f"di sini setelah ada transaksi pengeluaran bulan ini.")
            st.dataframe(pd.DataFrame([
                {"Kategori Pengeluaran": "-", "Nilai (Rp)": "-",
                 "Porsi (%)": "-"}]),
                use_container_width=True, hide_index=True)
            st.markdown("**💸 Total Pengeluaran Bulan Ini: -**")
        else:
            total = float(grup.sum())
            data = pd.DataFrame({"Kategori": [str(k) for k in grup.index],
                                 "Nilai": [float(v) for v in grup.values]})
            data["Persen"] = data["Nilai"] / total * 100.0
            data["Label"] = [f"{format_rupiah(n)} ({p:.1f}%)"
                             for n, p in zip(data["Nilai"], data["Persen"])]
            import altair as alt
            pie = alt.Chart(data).mark_arc(innerRadius=55, outerRadius=120).encode(
                theta=alt.Theta("Nilai:Q", stack=True),
                color=alt.Color("Kategori:N", legend=alt.Legend(title="Kategori")),
                tooltip=["Kategori:N", "Label:N"])
            st.altair_chart(pie, use_container_width=True)
            st.dataframe(pd.DataFrame({
                "Kategori Pengeluaran": data["Kategori"],
                "Nilai (Rp)": data["Nilai"].map(format_rupiah),
                "Porsi (%)": data["Persen"].map(lambda p: f"{p:.1f}%")}),
                use_container_width=True, hide_index=True)
            st.markdown(f"**💸 Total Pengeluaran Bulan Ini: "
                        f"{format_rupiah(total)}**")
    with tab4:
        t1, t2 = st.columns(2)
        t1.metric("Laba/Rugi Bulan Ini",
                  tanda_rp(None if not ada_trx else
                           (omzet_map.get(now.month, 0.0)
                            - keluar_map.get(now.month, 0.0))),
                  help="Penjualan emas − pengeluaran pada bulan berjalan. "
                       "Hijau = laba, merah = rugi.")
        t2.metric("Laba/Rugi Tahun Berjalan",
                  tanda_rp(None if not ada_tahun else
                           (sum(ser_omzet) - sum(ser_keluar))),
                  help="Akumulasi laba/rugi sejak bulan pertama tahun ini "
                       "sampai bulan berjalan.")
        grafik_batang_laba(label_bulan, ser_laba)
        with st.expander("🔢 Lihat angka per bulan"):
            nilai = ["-" if v is None else tanda_rp(v) for v in ser_laba]
            st.dataframe(pd.DataFrame({"Bulan": label_bulan,
                                       "Laba/Rugi (Rp)": nilai}),
                         use_container_width=True, hide_index=True)                        

    st.markdown("---")

    # -----------------------------------------------------------------
    # 4) TRACKING — aktivitas admin terakhir + peta lokasi
    # -----------------------------------------------------------------
    st.subheader("📍 Tracking Admin — Aktivitas Terakhir")
    k1, k2 = st.columns([1.1, 1.4])
    if ada_trx:
        terakhir = df_trx.iloc[-1]
        lat = float(terakhir["Latitude"] or 0)
        lon = float(terakhir["Longitude"] or 0)
        with k1:
            st.markdown(f"👤 **{terakhir['Admin'] or '-'}**")
            st.caption(f"🕒 {terakhir['Tanggal'] or '-'} "
                       f"({terakhir['Hari'] or '-'}) · {terakhir['Jam'] or '-'} WIB")
            st.markdown(f"🏢 {terakhir['Lokasi_Nama'] or '-'}")
            st.caption(f"📝 {terakhir['Aktivitas'] or '-'}")
            if lat != 0 or lon != 0:
                st.markdown(f"📍 {lat:.5f}, {lon:.5f} · "
                            f"[🗺️ Buka di Google Maps]"
                            f"(https://www.google.com/maps?q={lat},{lon})")
        with k2:
            if lat != 0 or lon != 0:
                try:
                    import streamlit.components.v1 as components
                    components.html(
                        f'<iframe src="https://maps.google.com/maps?q={lat},{lon}'
                        f'&z=15&output=embed" width="100%" height="270" '
                        f'style="border:0;border-radius:10px"></iframe>',
                        height=280)
                except Exception:
                    st.caption("Peta tidak dapat ditampilkan.")
            else:
                st.info("Koordinat tidak tersedia pada aktivitas terakhir. "
                        "Pastikan mengisi koordinat GPS saat input data agar "
                        "peta muncul.")
    else:
        with k1:
            st.markdown("👤 **-**")
            st.caption("🕒 - (-) · - WIB")
            st.markdown("🏢 -")
            st.caption("📝 -")
        with k2:
            st.info("Peta lokasi admin akan tampil di sini setelah ada transaksi "
                    "yang disimpan dengan koordinat GPS.")
    with st.expander("🕒 5 Aktivitas Admin Terakhir"):
        if ada_trx:
            lima = df_trx.tail(5).iloc[::-1]
            st.dataframe(lima[["Tanggal", "Hari", "Jam", "Admin", "Lokasi_Nama",
                               "Aktivitas", "Jenis"]],
                         use_container_width=True, hide_index=True)
        else:
            st.caption("Belum ada aktivitas admin tercatat — daftar akan tampil "
                       "otomatis setelah transaksi pertama disimpan.")


# =====================================================================
#  HALAMAN: KEUANGAN
# =====================================================================
def halaman_keuangan():
    st.header("💰 Keuangan — Buku Kas")
    df_trx = ambil_transaksi()
    df_rek = ambil_rekening()
    if df_trx.empty:
        st.info("📭 Belum ada transaksi.")
        return
    st.caption("Debit = uang MASUK · Kredit = uang KELUAR")

    saldo = hitung_saldo(df_trx, df_rek)
    st.subheader("💵 Saldo Saat Ini")
    if saldo:
        cols = st.columns(4)
        for i, (nama, s) in enumerate(saldo.items()):
            cols[i % 4].metric(nama, format_rupiah(s))
        st.metric("TOTAL KAS", format_rupiah(sum(saldo.values())))

    st.subheader("📘 Buku Kas")
    c1, c2, c3 = st.columns(3)
    with c1:
        pilih = st.selectbox("Rekening", ["Semua Rekening"] + list(saldo.keys()))
    with c2:
        dari = st.date_input("Dari tanggal",
                             value=sekarang_wib().date() - timedelta(days=30))
    with c3:
        sampai = st.date_input("Sampai tanggal", value=sekarang_wib().date())

    if pilih == "Semua Rekening":
        buku = bangun_buku(df_trx)
        buku["_t"] = pd.to_datetime(buku["Tanggal"], errors="coerce")
        buku = buku[(buku["_t"].dt.date >= dari) & (buku["_t"].dt.date <= sampai)]
        if buku.empty:
            st.warning("Tidak ada transaksi pada rentang tanggal ini.")
            return
        t1, t2, t3 = st.columns(3)
        t1.metric("Total Debit (Masuk)", format_rupiah(buku["Debit (Masuk)"].sum()))
        t2.metric("Total Kredit (Keluar)", format_rupiah(buku["Kredit (Keluar)"].sum()))
        t3.metric("Arus Kas Bersih",
                  format_rupiah(buku["Debit (Masuk)"].sum() - buku["Kredit (Keluar)"].sum()))
        tampil = buku[["Tanggal", "Hari", "Jam", "Jenis", "Rekening",
                       "Debit (Masuk)", "Kredit (Keluar)", "Keterangan",
                       "Admin"]].copy()
        for c in ("Debit (Masuk)", "Kredit (Keluar)"):
            tampil[c] = tampil[c].map(lambda v: format_rupiah(v) if v else "")
        st.dataframe(tampil, use_container_width=True, hide_index=True)
        st.download_button("⬇️ Unduh CSV", buku.to_csv(index=False).encode("utf-8"),
                           "keuangan.csv", "text/csv")
    else:
        mut, saldo_awal = mutasi_rekening(df_trx, df_rek, pilih)
        mut["_t"] = pd.to_datetime(mut["Tanggal"], errors="coerce")
        mut_f = mut[(mut["_t"].dt.date >= dari) & (mut["_t"].dt.date <= sampai)]
        saldo_akhir = float(mut["Saldo"].iloc[-1]) if not mut.empty else saldo_awal
        t1, t2, t3, t4 = st.columns(4)
        t1.metric("Saldo Awal", format_rupiah(saldo_awal))
        t2.metric("Total Masuk", format_rupiah(mut_f["Masuk"].sum()))
        t3.metric("Total Keluar", format_rupiah(mut_f["Keluar"].sum()))
        t4.metric("Saldo Akhir (terkini)", format_rupiah(saldo_akhir))
        if mut_f.empty:
            st.warning("Tidak ada mutasi pada rentang tanggal ini.")
            return
        tampil = mut_f[["Tanggal", "Hari", "Jam", "Jenis", "Aktivitas",
                        "Lokasi_Nama", "Masuk", "Keluar", "Saldo"]].copy()
        for c in ("Masuk", "Keluar", "Saldo"):
            tampil[c] = tampil[c].map(lambda v: format_rupiah(v) if v else "")
        st.dataframe(tampil, use_container_width=True, hide_index=True)
        st.download_button("⬇️ Unduh CSV", mut_f.to_csv(index=False).encode("utf-8"),
                           f"mutasi_{pilih}.csv", "text/csv")


# =====================================================================
#  HALAMAN: REKENING
# =====================================================================
def halaman_rekening():
    st.header("🏦 Rekening")
    df_trx = ambil_transaksi()
    df_rek = ambil_rekening()
    if df_rek.empty:
        st.info("Belum ada rekening. Admin dapat menambahkannya di menu "
                "📝 Input Data → Kelola Rekening.")
        return
    if df_trx.empty:
        saldo = {str(r["Nama_Rekening"]).strip(): float(r["Saldo_Awal"] or 0)
                 for _, r in df_rek.iterrows()}
    else:
        saldo = hitung_saldo(df_trx, df_rek)

    st.subheader("💵 Saldo Saat Ini")
    cols = st.columns(3)
    for i, (nama, s) in enumerate(saldo.items()):
        cols[i % 3].metric(nama, format_rupiah(s))

    tampil = pd.DataFrame({
        "Nama Rekening": df_rek["Nama_Rekening"],
        "Jenis": df_rek["Jenis"],
        "No. Rekening": df_rek["No_Rekening"],
        "Atas Nama": df_rek["Atas_Nama"],
        "Saldo Awal": df_rek["Saldo_Awal"].map(format_rupiah),
        "Saldo Sekarang": df_rek["Nama_Rekening"].map(
            lambda n: format_rupiah(saldo.get(str(n).strip(), 0))),
        "Keterangan": df_rek["Keterangan"],
    })
    st.dataframe(tampil, use_container_width=True, hide_index=True)

    st.subheader("📘 Mutasi Rekening")
    pilih = st.selectbox("Pilih rekening", list(saldo.keys()))
    if pilih and not df_trx.empty:
        mut, saldo_awal = mutasi_rekening(df_trx, df_rek, pilih)
        if mut.empty:
            st.caption("Belum ada mutasi untuk rekening ini.")
        else:
            tampil = mut.tail(25)[["Tanggal", "Hari", "Jam", "Jenis", "Aktivitas",
                                    "Lokasi_Nama", "Masuk", "Keluar", "Saldo"]].copy()
            for c in ("Masuk", "Keluar", "Saldo"):
                tampil[c] = tampil[c].map(lambda v: format_rupiah(v) if v else "")
            st.caption(f"25 mutasi terakhir — saldo awal {format_rupiah(saldo_awal)}")
            st.dataframe(tampil, use_container_width=True, hide_index=True)


# =====================================================================
#  HALAMAN: STOK EMAS
# =====================================================================
def halaman_stok():
    st.header("🪙 Stok Emas")
    df_trx = ambil_transaksi()
    if df_trx.empty:
        st.info("📭 Belum ada transaksi.")
        return
    emas = data_emas(df_trx)
    if emas.empty:
        st.info("Belum ada transaksi Beli/Jual Emas.")
        return

    stok = float(emas["Sisa_Stok"].iloc[-1])
    beli = emas[emas["Jenis"] == "Beli Emas"]
    jual = emas[emas["Jenis"] == "Jual Emas"]
    tot_beli_g = float(beli["Berat_Gram"].sum())
    harga_rata = float(beli["Nominal"].sum()) / tot_beli_g if tot_beli_g > 0 else 0.0

    m = st.columns(4)
    m[0].metric("Stok Saat Ini", format_gram(stok))
    with m[1]:
        st.metric("Total Pembelian", format_gram(tot_beli_g))
        st.caption(format_rupiah(beli["Nominal"].sum()))
    with m[2]:
        st.metric("Total Penjualan", format_gram(jual["Berat_Gram"].sum()))
        st.caption(format_rupiah(jual["Nominal"].sum()))
    m[3].metric("Estimasi Nilai Stok", format_rupiah(stok * harga_rata))
    st.caption(f"Estimasi memakai harga rata-rata pembelian "
               f"{format_rupiah(harga_rata)}/gram.")

    st.subheader("📈 Sisa Stok Berjalan")
    dfe = emas.copy()
    dfe["_t"] = pd.to_datetime(dfe["Tanggal"], errors="coerce")
    st.line_chart(dfe.groupby(dfe["_t"].dt.date)["Arah_Gram"].sum().cumsum())

    st.subheader("📘 Riwayat Transaksi Emas (jual / beli / lokasi / sisa stok)")
    tampil = emas.copy()
    tampil["Berat_Gram"] = tampil["Berat_Gram"].map(format_gram)
    tampil["Harga_Per_Gram"] = tampil["Harga_Per_Gram"].map(format_rupiah)
    tampil["Nominal"] = tampil["Nominal"].map(format_rupiah)
    tampil["Sisa_Stok"] = tampil["Sisa_Stok"].map(format_gram)
    tampil = tampil.rename(columns={"Berat_Gram": "Berat",
                                    "Harga_Per_Gram": "Harga/gram",
                                    "Sisa_Stok": "Sisa Stok",
                                    "Lokasi_Nama": "Lokasi"})
    st.dataframe(tampil[["Tanggal", "Hari", "Jam", "Jenis", "Berat", "Harga/gram",
                         "Nominal", "Lokasi", "Sisa Stok", "Admin"]],
                 use_container_width=True, hide_index=True)
    st.download_button("⬇️ Unduh CSV", emas.to_csv(index=False).encode("utf-8"),
                       "stok_emas.csv", "text/csv")


# =====================================================================
#  HALAMAN: TRACKING ADMIN
# =====================================================================
def halaman_tracking():
    st.header("📍 Tracking Admin")
    df = ambil_transaksi()
    if df.empty:
        st.info("📭 Belum ada aktivitas admin yang tercatat.")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        f_admin = st.selectbox("Admin",
                               ["Semua"] + sorted(df["Admin"].unique().tolist()))
    with c2:
        dari = st.date_input("Dari tanggal",
                             value=sekarang_wib().date() - timedelta(days=7))
    with c3:
        jumlah = st.selectbox("Tampilkan (terbaru)", [10, 20, 50, 100], index=1)

    d = df.copy()
    d["_t"] = pd.to_datetime(d["Tanggal"], errors="coerce")
    msk = d["_t"].dt.date >= dari
    if f_admin != "Semua":
        msk &= d["Admin"].eq(f_admin)
    d = d[msk].tail(jumlah).iloc[::-1]
    st.caption(f"Menampilkan {len(d)} aktivitas (terbaru di atas).")

    kolom_csv = [c for c in d.columns if c != "_baris"]
    st.download_button("⬇️ Unduh CSV (tanpa foto)",
                       d[kolom_csv].to_csv(index=False).encode("utf-8"),
                       "tracking_admin.csv", "text/csv")

    # Foto diunduh HANYA untuk baris yang ditampilkan
    foto_map = {}
    if (not d.empty) and ("_baris" in d.columns):
        baris_tampil = [int(x) for x in d["_baris"].tolist()]
        awal, akhir = min(baris_tampil), max(baris_tampil)
        if akhir >= awal:
            try:
                foto_map = ambil_foto_terbatas(awal, akhir)
            except Exception:
                foto_map = {}

    for _, r in d.iterrows():
        with st.container(border=True):
            k1, k2, k3 = st.columns([1.1, 1.6, 0.9])
            with k1:
                st.markdown(f"**{r['Tanggal']}** ({r['Hari']})")
                st.caption(f"⏰ {r['Jam']} WIB · 👤 {r['Admin']}")
                st.markdown(f"🏷️ {r['Jenis']}")
                st.caption(f"ID: {r['ID']}")
            with k2:
                st.markdown(f"🏢 **{r['Lokasi_Nama'] or '-'}**")
                if r["Latitude"] or r["Longitude"]:
                    st.caption(f"📍 {r['Latitude']:.5f}, {r['Longitude']:.5f}")
                    st.markdown(
                        f"[🗺️ Buka di Google Maps]"
                        f"(https://www.google.com/maps?q={r['Latitude']},{r['Longitude']})")
                st.caption(f"📝 {r['Aktivitas'] or '-'}")
            with k3:
                foto = ""
                if "_baris" in r.index:
                    foto = foto_map.get(int(r["_baris"]), "")
                if foto:
                    try:
                        img = Image.open(io.BytesIO(base64.b64decode(foto)))
                        st.image(img, use_container_width=True)
                    except Exception:
                        st.caption("foto tidak dapat dibaca")
                else:
                    st.caption("tanpa foto")


# =====================================================================
#  HALAMAN: INPUT DATA (KHUSUS ADMIN — SINGLE ENTRY FORM)
# =====================================================================
# =====================================================================
#  BANTU: INPUT VIA EXCEL
# =====================================================================
KOLOM_EXCEL = ["Tanggal", "Hari", "Jam", "Jenis", "Berat (gram)",
               "Harga per Gram (Rp)", "Nominal (Rp)", "Rekening",
               "Rekening Tujuan", "Nama Lokasi", "Latitude", "Longitude",
               "Aktivitas"]

SINONIM_EXCEL = {
    "tanggal": "Tanggal", "tgl": "Tanggal", "tanggaltransaksi": "Tanggal",
    "hari": "Hari",
    "jam": "Jam", "waktu": "Jam",
    "jenis": "Jenis", "jenistransaksi": "Jenis", "transaksi": "Jenis",
    "berat": "Berat (gram)", "beratgram": "Berat (gram)",
    "beratemas": "Berat (gram)", "gram": "Berat (gram)",
    "harga": "Harga per Gram (Rp)", "hargapergram": "Harga per Gram (Rp)",
    "hargapergramrp": "Harga per Gram (Rp)", "hargaemas": "Harga per Gram (Rp)",
    "hargagram": "Harga per Gram (Rp)",
    "nominal": "Nominal (Rp)", "nominalrp": "Nominal (Rp)",
    "total": "Nominal (Rp)", "totalnominal": "Nominal (Rp)",
    "jumlah": "Nominal (Rp)", "nilai": "Nominal (Rp)",
    "rekening": "Rekening", "rekeningutama": "Rekening", "rek": "Rekening",
    "darirekening": "Rekening", "rekeningasal": "Rekening",
    "rekeningtujuan": "Rekening Tujuan", "kerekening": "Rekening Tujuan",
    "rektujuan": "Rekening Tujuan",
    "namalokasi": "Nama Lokasi", "lokasi": "Nama Lokasi",
    "tempat": "Nama Lokasi", "namatempat": "Nama Lokasi",
    "lokasitransaksi": "Nama Lokasi",
    "latitude": "Latitude", "lat": "Latitude",
    "longitude": "Longitude", "lon": "Longitude", "long": "Longitude",
    "lng": "Longitude",
    "aktivitas": "Aktivitas", "keterangan": "Aktivitas", "uraian": "Aktivitas",
    "catatan": "Aktivitas", "deskripsi": "Aktivitas",
}


def _norm_header(h):
    """Seragamkan judul kolom: kecilkan huruf & buang spasi/tanda baca."""
    return re.sub(r"[^a-z0-9]", "", str(h).lower())


def parse_angka_excel(nilai, ribuan_dulu=True):
    """Ubah isi sel menjadi angka. Menerima 1500000, '1500000',
    '1.500.000', '1,5', 'Rp 1.500.000', dsb.
    ribuan_dulu=True untuk uang ('2.000' = 2000),
    ribuan_dulu=False untuk berat ('2.000' = 2,0 gram)."""
    if nilai is None or isinstance(nilai, bool):
        return None
    if isinstance(nilai, (int, float)):
        f = float(nilai)
        return None if pd.isna(f) else f
    s = str(nilai).strip().lower().replace("rp", "").replace(" ", "")
    s = s.replace("gram", "").replace("gr", "")
    if s in ("", "-", "none", "nan", "nat"):
        return None
    neg = s.startswith("-") or s.startswith("(")
    s = s.strip("()").lstrip("-")
    if not re.search(r"\d", s):
        return None
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")   # 1.234.567,89
        else:
            s = s.replace(",", "")                     # 1,234,567.89
    elif "," in s:
        bagian = s.split(",")
        if len(bagian) == 2 and len(bagian[1]) != 3:
            s = s.replace(",", ".")                    # 2,5 -> 2.5
        else:
            s = s.replace(",", "")                     # 1,500,000
    elif "." in s:
        bagian = s.split(".")
        if len(bagian) > 2:
            s = s.replace(".", "")                     # 1.500.000
        elif len(bagian[1]) == 3 and bagian[0] != "0":
            if ribuan_dulu:                            # 2.000 -> uang: 2000
                s = s.replace(".", "")                 # berat: tetap 2.0
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def parse_tanggal_excel(nilai):
    """Terima 2026-01-05, 05/01/2026, 05-01-2026, dsb (tgl-bulan-tahun)."""
    if nilai is None:
        return None
    if isinstance(nilai, (pd.Timestamp, dt.datetime, dt.date)):
        return pd.Timestamp(nilai).date()
    s = str(nilai).strip()
    if not s or s.lower() in ("", "nan", "none", "nat", "-"):
        return None
    s = s.split(" ")[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y",
                "%d-%m-%y", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    t = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.isna(t):
        return None
    return t.date()


def parse_jam_excel(nilai):
    """Terima '13:45', sel waktu Excel, atau ambil dari sel tanggal."""
    if nilai is None:
        return None
    if isinstance(nilai, (dt.time, pd.Timestamp, dt.datetime)):
        return pd.Timestamp(nilai).strftime("%H:%M") if not isinstance(nilai, dt.time) \
            else nilai.strftime("%H:%M")
    s = str(nilai).strip()
    if not s or s.lower() in ("", "nan", "none", "nat", "-"):
        return None
    m = re.search(r"(\d{1,2})[:.](\d{2})", s)
    if m:
        j, mnt = int(m.group(1)), int(m.group(2))
        if 0 <= j <= 23 and 0 <= mnt <= 59:
            return f"{j:02d}:{mnt:02d}"
    return None


def parse_jenis_excel(teks):
    """Kenali jenis transaksi walau ditulis bebas (mis. 'penjualan')."""
    if teks is None:
        return None
    s = re.sub(r"[^a-z]", "", str(teks).lower())
    if not s:
        return None
    if "transfer" in s:
        return "Transfer Antar Rekening"
    if "beli" in s:
        return "Beli Emas"
    if "jual" in s:
        return "Jual Emas"
    if "operasional" in s or "pengeluaran" in s or "biaya" in s:
        return "Pengeluaran Operasional"
    if "pemasukan" in s or "masuk" in s or "lain" in s:
        return "Pemasukan Lain"
    return None


def proses_file_excel(df_raw, daftar_rek):
    """Validasi isi file Excel -> (baris_siap_simpan, daftar_error,
    df_pratinjau). Prinsip: bila ada 1 baris salah, TIDAK ada yang disimpan."""
    if df_raw is None or df_raw.empty:
        return [], ["File kosong — tidak ada baris data."], pd.DataFrame()

    petamax = {}
    for kolom_asli in df_raw.columns:
        k = _norm_header(kolom_asli)
        if not k:
            continue
        target = SINONIM_EXCEL.get(k)
        if target and target not in petamax:
            petamax[target] = kolom_asli

    wajib = ["Tanggal", "Jenis", "Rekening", "Nama Lokasi"]
    kurang = [k for k in wajib if k not in petamax]
    if kurang:
        return [], ["Kolom wajib tidak ditemukan di file: "
                    f"{', '.join(kurang)}. Unduh & gunakan template yang "
                    "disediakan agar judul kolom sesuai."], pd.DataFrame()

    def ambil(row, k):
        if k not in petamax:
            return None
        v = row[petamax[k]]
        if v is None:
            return None
        if isinstance(v, float) and pd.isna(v):
            return None
        if isinstance(v, str) and v.strip().lower() in ("", "nan", "none",
                                                        "nat", "-"):
            return None
        return v

    rek_map = {str(r).strip().lower(): str(r).strip() for r in daftar_rek}
    baris_siap, pratinjau, errors = [], [], []

    for nomor, (_, row) in enumerate(df_raw.iterrows(), start=2):
        if all(str(v).strip().lower() in ("", "nan", "none", "nat", "-")
               for v in row.values):
            continue  # baris kosong dilewati

        msl = []
        tanggal = parse_tanggal_excel(ambil(row, "Tanggal"))
        if tanggal is None:
            msl.append("Tanggal kosong/tidak valid (contoh benar: 05/01/2026 "
                       "atau 2026-01-05)")
        jenis = parse_jenis_excel(ambil(row, "Jenis"))
        if jenis is None:
            msl.append("Jenis tidak dikenali. Pilihan: Beli Emas / Jual Emas / "
                       "Pemasukan Lain / Pengeluaran Operasional / Transfer "
                       "Antar Rekening")
        berat = parse_angka_excel(ambil(row, "Berat (gram)"), ribuan_dulu=False)
        berat = berat if berat else 0.0
        harga = parse_angka_excel(ambil(row, "Harga per Gram (Rp)"),
                                  ribuan_dulu=True)
        harga = harga if harga else 0.0
        nominal = parse_angka_excel(ambil(row, "Nominal (Rp)"),
                                    ribuan_dulu=True)

        rek_asli = ambil(row, "Rekening")
        rek = (rek_map.get(str(rek_asli).strip().lower())
               if rek_asli is not None else None)
        if rek is None:
            msl.append(f"Rekening '{rek_asli if rek_asli is not None else ''}' "
                       f"tidak terdaftar. Rekening tersedia: "
                       f"{', '.join(daftar_rek) if daftar_rek else '(belum ada rekening)'}")
        rek_t_asli = ambil(row, "Rekening Tujuan")
        rek_t = (rek_map.get(str(rek_t_asli).strip().lower())
                 if rek_t_asli is not None else None)

        lokasi = str(ambil(row, "Nama Lokasi") or "").strip()
        lat = parse_angka_excel(ambil(row, "Latitude"), ribuan_dulu=False) or 0.0
        lon = parse_angka_excel(ambil(row, "Longitude"), ribuan_dulu=False) or 0.0
        aktivitas = str(ambil(row, "Aktivitas") or "").strip()

        if jenis in ("Beli Emas", "Jual Emas"):
            if berat <= 0:
                msl.append("Berat (gram) wajib lebih dari 0 untuk transaksi emas")
            if harga <= 0:
                msl.append("Harga per Gram wajib lebih dari 0 untuk transaksi emas")
            if nominal is None:
                nominal = berat * harga
        if nominal is None or nominal <= 0:
            msl.append("Nominal (Rp) wajib lebih dari 0 — atau kosongkan Nominal "
                       "untuk emas (otomatis dihitung Berat × Harga)")
        if jenis == "Transfer Antar Rekening":
            if rek_t is None:
                msl.append("Rekening Tujuan wajib diisi & terdaftar untuk jenis "
                           "Transfer")
            elif rek is not None and rek_t == rek:
                msl.append("Rekening Tujuan tidak boleh sama dengan Rekening asal")
        if not lokasi:
            msl.append("Nama Lokasi wajib diisi (isi tanda - jika tidak ada)")

        if msl:
            errors.append({"Baris Excel": nomor, "Masalah": "; ".join(msl)})
            continue

        hari_isi = ambil(row, "Hari")
        hari_final = (str(hari_isi).strip() if hari_isi is not None else "") \
            or NAMA_HARI[tanggal.weekday()]
        jam_final = parse_jam_excel(ambil(row, "Jam"))
        if jam_final is None:
            jam_final = parse_jam_excel(ambil(row, "Tanggal"))
        if jam_final is None:
            jam_final = "12:00"
        arah = ("Debit" if jenis in JENIS_MASUK else
                ("Kredit" if jenis in JENIS_KELUAR else "Transfer"))
        aktivitas_final = aktivitas or f"{jenis} — {lokasi}"
        id_trx = (f"TRX-{tanggal:%Y%m%d}-{sekarang_wib():%H%M%S}-"
                  f"{uuid.uuid4().hex[:4].upper()}")
        baris_siap.append([
            id_trx,
            tanggal.strftime("%Y-%m-%d"),
            hari_final,
            jam_final,
            jenis,
            float(berat),
            int(harga),
            int(round(float(nominal))),
            rek,
            (rek_t or "") if jenis == "Transfer Antar Rekening" else "",
            arah,
            lokasi,
            float(lat),
            float(lon),
            aktivitas_final,
            st.session_state.get("username", ""),
            sekarang_wib().strftime("%Y-%m-%d %H:%M:%S"),
            ""])
        pratinjau.append({
            "Tanggal": tanggal.strftime("%Y-%m-%d"),
            "Hari": hari_final, "Jam": jam_final, "Jenis": jenis,
            "Berat": format_gram(berat) if berat else "-",
            "Harga/gram": format_rupiah(harga) if harga else "-",
            "Nominal": format_rupiah(nominal),
            "Rekening": rek,
            "Tujuan": (rek_t or "") or "-",
            "Lokasi": lokasi,
            "Koordinat": (f"{lat:.5f}, {lon:.5f}" if (lat or lon) else "-"),
            "Aktivitas": aktivitas_final,
        })

    return baris_siap, errors, pd.DataFrame(pratinjau)


# =====================================================================
#  HALAMAN: INPUT DATA (KHUSUS ADMIN — MODE MANUAL & MODE EXCEL)
# =====================================================================
def halaman_input():
    st.header("📝 Input Data — Satu Form untuk Semua Menu")
    st.caption("✅ Sekali SIMPAN → Dashboard, Keuangan, Rekening, Stok Emas, dan "
               "Tracking Admin otomatis ter-update.")

    df_rek = ambil_rekening()
    daftar_rek = [x for x in df_rek["Nama_Rekening"].tolist() if x]

    mode = st.radio("Pilih cara input data:",
                    ["✍️ Input Manual", "📥 Input Excel", "🧪 Data Dummy"],
                    horizontal=True)
    if mode == "🧪 Data Dummy":
        halaman_input_dummy(df_rek)
    elif mode == "📥 Input Excel":
        halaman_input_excel(daftar_rek)
    else:
        form_input_manual(daftar_rek)

    # ---------- PENGATURAN DATA MASTER ----------
    st.markdown("---")
    st.subheader("⚙️ Pengaturan Data Master (khusus admin)")
    pesan_rek = st.session_state.pop("pesan_hapus_rek", None)
    if pesan_rek:
        st.success(pesan_rek)

    with st.expander("🏦 Kelola Rekening — tambah / hapus rekening & kas"):
        if not df_rek.empty:
            st.dataframe(df_rek[["Nama_Rekening", "Jenis", "No_Rekening",
                                 "Atas_Nama", "Saldo_Awal"]],
                         use_container_width=True, hide_index=True)
        with st.form("form_rek", border=True, clear_on_submit=True):
            st.markdown("**➕ Tambah Rekening Baru**")
            n1, n2 = st.columns(2)
            nama_baru = n1.text_input("Nama Rekening *",
                                      placeholder="mis. BCA, Kas Toko, GoPay")
            jenis_baru = n2.selectbox("Jenis",
                                      ["Kas Tunai", "Bank", "E-Wallet", "Lainnya"])
            n3, n4 = st.columns(2)
            no_baru = n3.text_input("No. Rekening (opsional)")
            pemilik = n4.text_input("Atas Nama (opsional)")
            saldo_awal = st.number_input("Saldo Awal (Rp)", min_value=0.0,
                                         step=1000.0, format="%.0f")
            ket_baru = st.text_input("Keterangan (opsional)")
            tambah = st.form_submit_button("➕ Tambah Rekening", type="primary")
        if tambah:
            if not nama_baru.strip():
                st.error("Nama rekening wajib diisi.")
            elif nama_baru.strip() in daftar_rek:
                st.error("Nama rekening sudah ada.")
            else:
                try:
                    tambah_baris("Rekening", [nama_baru.strip(), jenis_baru,
                                              no_baru.strip(), pemilik.strip(),
                                              float(saldo_awal),
                                              ket_baru.strip()])
                    st.success(f"Rekening '{nama_baru}' ditambahkan.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal: {e}")

        st.markdown("---")
        st.markdown("**🗑️ Hapus Rekening**")
        if not daftar_rek:
            st.caption("Belum ada rekening yang terdaftar.")
        else:
            # jaring pengaman: reset pilihan bila rekening terpilih sudah
            # tidak ada lagi (misal baru saja dihapus)
            if st.session_state.get("hapus_rek_pilih") not in daftar_rek:
                st.session_state.pop("hapus_rek_pilih", None)
            target_hapus = st.selectbox("Pilih rekening yang akan dihapus",
                                        daftar_rek, key="hapus_rek_pilih")
            try:
                df_trx_rek = ambil_transaksi()
            except Exception:
                df_trx_rek = pd.DataFrame()
            if df_trx_rek.empty:
                n_pakai = 0
            else:
                t_low = target_hapus.strip().lower()
                n_pakai = int(
                    (df_trx_rek["Rekening_Utama"].astype(str).str.strip()
                     .str.lower() == t_low).sum()
                    + (df_trx_rek["Rekening_Tujuan"].astype(str).str.strip()
                       .str.lower() == t_low).sum())
            info_rek = df_rek[df_rek["Nama_Rekening"].astype(str).str.strip()
                              .str.lower() == target_hapus.strip().lower()]
            saldo_awal_rek = (float(info_rek.iloc[0]["Saldo_Awal"])
                              if not info_rek.empty else 0.0)
            st.caption(f"Saldo awal tercatat: {format_rupiah(saldo_awal_rek)} · "
                       f"Digunakan oleh {n_pakai} transaksi.")

            mode_hapus, pengganti = None, None
            if n_pakai > 0:
                mode_hapus = st.radio(
                    "Rekening ini masih dipakai transaksi. Pilih tindakan:",
                    ["Riwayat transaksi TETAP tersimpan (nama lama tetap "
                     "tercatat di laporan)",
                     "Pindahkan seluruh transaksinya ke rekening lain"])
                if mode_hapus.startswith("Pindahkan"):
                    lain = [r for r in daftar_rek if r != target_hapus]
                    if not lain:
                        st.warning("Tidak ada rekening lain sebagai tujuan "
                                   "pindah — tambahkan rekening baru dulu, atau "
                                   "pilih opsi riwayat tetap tersimpan.")
                    else:
                        if st.session_state.get("hapus_rek_ganti") not in lain:
                            st.session_state.pop("hapus_rek_ganti", None)
                        pengganti = st.selectbox("Pindahkan transaksinya ke:",
                                                 lain, key="hapus_rek_ganti")
            if len(daftar_rek) == 1:
                st.warning("⚠️ Ini satu-satunya rekening — setelah dihapus, "
                           "tambahkan rekening baru terlebih dahulu sebelum "
                           "bisa input transaksi lagi.")
            yakin_rek = st.checkbox("Saya yakin — hapus rekening ini permanen")
            if st.button("🗑️ Hapus Rekening Sekarang", disabled=not yakin_rek):
                pindah_mode = (n_pakai > 0 and mode_hapus is not None
                               and mode_hapus.startswith("Pindahkan"))
                if pindah_mode and not pengganti:
                    st.error("Pilih rekening tujuan pindahan terlebih dahulu "
                             "(atau pilih opsi riwayat tetap tersimpan).")
                else:
                    with st.spinner("Memproses penghapusan..."):
                        try:
                            sh = buat_koneksi()
                            # opsional: pindahkan transaksi ke rekening lain
                            if n_pakai > 0 and pengganti:
                                resp = sh.values_get("Transaksi!I2:J")
                                updates = []
                                t_low = target_hapus.strip().lower()
                                for i, r in enumerate(resp.get("values", [])):
                                    nomor = i + 2
                                    u = (r[0] if len(r) > 0 else "").strip()
                                    t = (r[1] if len(r) > 1 else "").strip()
                                    if u.lower() == t_low:
                                        updates.append({"range": f"I{nomor}",
                                                        "values": [[pengganti]]})
                                    if t.lower() == t_low:
                                        updates.append({"range": f"J{nomor}",
                                                        "values": [[pengganti]]})
                                if updates:
                                    sh.worksheet("Transaksi").batch_update(updates)
                            # hapus baris rekening dari sheet Rekening
                            wr = sh.worksheet("Rekening")
                            semua = wr.get_all_values()
                            baris_hapus = None
                            for i, r in enumerate(semua):
                                if i > 0 and str(r[0]).strip().lower() == \
                                        target_hapus.strip().lower():
                                    baris_hapus = i + 1
                                    break
                            if baris_hapus is None:
                                st.error("Rekening tidak ditemukan di sheet — "
                                         "mungkin sudah dihapus sebelumnya. "
                                         "Muat ulang halaman.")
                            else:
                                wr.delete_rows(baris_hapus)
                                st.cache_data.clear()
                                pesan = (f"✅ Rekening **{target_hapus}** "
                                         "berhasil dihapus.")
                                if n_pakai > 0 and pengganti:
                                    pesan += (f" {n_pakai} transaksinya "
                                              f"dipindahkan ke **{pengganti}**.")
                                elif n_pakai > 0:
                                    pesan += (" Riwayat transaksinya tetap "
                                              "tersimpan (nama lama tetap "
                                              "tercatat di laporan).")
                                st.session_state["pesan_hapus_rek"] = pesan
                                st.rerun()
                        except Exception as e:
                            st.error(f"Gagal menghapus rekening: {e}")

        st.caption("✏️ Mengubah detail rekening (nama/no. rekening/saldo awal): "
                   "edit langsung barisnya di Google Sheets, tab 'Rekening'. "
                   "Menambah & menghapus rekening cukup lewat aplikasi ini.")

    with st.expander("🗑️ Hapus Transaksi (jika salah input)"):
        if st.session_state.pop("pesan_hapus_trx", False):
            st.success("✅ Transaksi dihapus. Semua laporan (kas, stok, saldo) "
                       "otomatis terkoreksi.")
        if st.button("🔄 Muat / Segarkan Daftar Transaksi"):
            try:
                sh = buat_koneksi()
                resp = sh.values_get("Transaksi!A2:H")
                st.session_state["hapus_trx_list"] = resp.get("values", [])
            except Exception as e:
                st.error(f"Gagal membaca data transaksi: {e}")
        baris_data = st.session_state.get("hapus_trx_list")
        if baris_data is None:
            st.caption("Daftar sengaja TIDAK dimuat otomatis agar halaman ini "
                       "ringan & tanpa loading saat Anda mengisi formulir. Klik "
                       "tombol 🔄 di atas untuk menampilkan daftar transaksi "
                       "yang bisa dihapus.")
        else:
            opsi, id_map = [], {}
            for b in baris_data:
                if not b or not str(b[0]).strip():
                    continue
                idt = str(b[0]).strip()
                label = (f"{idt} — {b[1] if len(b) > 1 else ''} "
                         f"{b[3] if len(b) > 3 else ''} — "
                         f"{b[4] if len(b) > 4 else ''} — "
                         f"{format_rupiah(b[7] if len(b) > 7 else 0)}")
                opsi.append(label)
                id_map[label] = idt
            if not opsi:
                st.caption("Belum ada transaksi.")
            else:
                pilihan_hapus = st.selectbox(
                    "Pilih transaksi yang salah (terbaru di atas)",
                    list(reversed(opsi)))
                yakin = st.checkbox("Saya yakin — hapus permanen transaksi ini")
                if st.button("🗑️ Hapus Sekarang", disabled=not yakin):
                    try:
                        sh = buat_koneksi()
                        wtr = sh.worksheet("Transaksi")
                        cari = None
                        try:
                            cari = wtr.find(id_map[pilihan_hapus])
                        except Exception:
                            cari = None
                        if cari is None:
                            st.error("Transaksi tidak ditemukan di sheet — "
                                     "mungkin sudah dihapus. Klik 🔄 Muat untuk "
                                     "menyegarkan daftar.")
                        else:
                            wtr.delete_rows(cari.row)
                            st.cache_data.clear()
                            resp = sh.values_get("Transaksi!A2:H")
                            st.session_state["hapus_trx_list"] = resp.get(
                                "values", [])
                            st.session_state["pesan_hapus_trx"] = True
                            st.rerun()
                    except Exception as e:
                        st.error(f"Gagal menghapus: {e}")

    st.caption("👤 Kelola pengguna (profil, ganti password, tambah user) "
               "sekarang ada di menu **⚙️ Pengaturan**.")

        # ---------- HAPUS SEMUA DATA (RESET DATABASE) ----------
    with st.expander("🧹 Hapus Semua Data — Reset Database (HATI-HATI)"):
        st.caption("Mengosongkan SELURUH transaksi (termasuk foto bukti) dari "
                   "database. Dipakai untuk membersihkan data dummy/latihan "
                   "sebelum memasukkan data asli. Akun login TIDAK dihapus.")
        st.warning("⚠️ Tindakan ini TIDAK dapat dibatalkan. Pastikan Anda "
                   "sudah mengekspor/mencadangkan data penting (menu "
                   "Keuangan/Stok Emas punya tombol Unduh CSV).")

        try:
            df_cek = ambil_transaksi()
            n_data = 0 if df_cek.empty else len(df_cek)
        except Exception:
            n_data = None

        if n_data == 0:
            st.success("✅ Database transaksi sudah kosong — tidak ada yang "
                       "perlu dihapus. Anda bisa langsung menginput data asli.")
        else:
            st.markdown(f"**Data yang akan dihapus:** {n_data} transaksi "
                        "(seluruh baris pada tab 'Transaksi' di Google "
                        "Sheets, termasuk semua foto bukti).")
            opsi_reset = st.radio(
                "Rekening & saldo awal ikut dibersihkan?",
                ["Tidak — pertahankan daftar rekening (disarankan)",
                 "Ya — kosongkan juga daftar rekening"],
                help="Jika dipilih 'Ya': seluruh baris tab 'Rekening' "
                     "dihapus. Anda mendaftarkan ulang rekening setelah "
                     "reset (bisa lewat Kelola Rekening di atas).")
            ketik = st.text_input(
                "Ketik persis kalimat berikut untuk mengonfirmasi: "
                "**HAPUS SEMUA DATA**",
                placeholder="HAPUS SEMUA DATA")
            if st.button("🧹 HAPUS SEKARANG — SEMUA TRANSAKSI",
                         disabled=(ketik.strip() != "HAPUS SEMUA DATA"),
                         type="primary"):
                with st.spinner("Menghapus seluruh data transaksi..."):
                    try:
                        sh = buat_koneksi()
                        # 1) kosongkan seluruh baris transaksi (kolom A s.d. R)
                        sh.values_clear("Transaksi!A2:R100000")
                        # 2) opsional: kosongkan daftar rekening
                        if opsi_reset.startswith("Ya"):
                            sh.values_clear("Rekening!A2:F1000")
                        st.cache_data.clear()
                        st.session_state.pop("hapus_trx_list", None)
                        st.session_state.pop("dummy_df", None)
                        st.session_state.pop("dummy_konfig", None)
                        st.session_state.pop("dummy_ada_trx", None)
                        pesan = "✅ Semua data transaksi berhasil dihapus."
                        if opsi_reset.startswith("Ya"):
                            pesan += (" Daftar rekening juga dikosongkan — "
                                      "daftarkan ulang rekening Anda di "
                                      "'Kelola Rekening' sebelum input data "
                                      "baru.")
                        pesan += (" Database siap diisi data asli. Semua "
                                  "menu kembali menampilkan template kosong.")
                        st.success(pesan)
                    except Exception as e:
                        st.error(f"Gagal menghapus data: {e}")


def form_input_manual(daftar_rek):
    if st.session_state.get("pesan_sukses"):
        st.success(f"✅ Transaksi berhasil disimpan — ID: "
                   f"**{st.session_state['pesan_sukses']}**")
        del st.session_state["pesan_sukses"]

    if "form_id" not in st.session_state:
        st.session_state["form_id"] = 0
    fid = st.session_state["form_id"]

    # ---------- 1. DATA OTOMATIS ----------
    st.subheader("1️⃣ Data Waktu & Admin (otomatis)")
    c1, c2, c3 = st.columns(3)
    with c1:
        tanggal = st.date_input("📅 Tanggal", value=sekarang_wib().date(),
                                key=f"tgl_{fid}")
    with c2:
        jam = st.time_input("⏰ Jam", value=sekarang_wib().time(), key=f"jam_{fid}")
    with c3:
        st.metric("📆 Hari", NAMA_HARI[tanggal.weekday()])
    st.caption(f"👤 Admin PIC: **{st.session_state.get('nama', '')}** "
               f"({st.session_state.get('username', '')}) — terisi otomatis dari login.")

    # ---------- 2. LOKASI / GPS ----------
    st.subheader("2️⃣ Lokasi / GPS")
    paste = st.text_input(
        "📍 Tempel koordinat dari Google Maps",
        placeholder="contoh: -6.2146, 106.8451",
        key=f"paste_{fid}",
        help="Cara: buka Google Maps → tahan-tekan titik lokasi → koordinat muncul "
             "→ salin (copy) → tempel (paste) di sini. Lat & Long di bawah terisi otomatis.")
    if paste and st.session_state.get(f"paste_ok_{fid}") != paste:
        p = parse_koordinat(paste)
        if p:
            st.session_state[f"lat_{fid}"] = p[0]
            st.session_state[f"lon_{fid}"] = p[1]
            st.session_state[f"paste_ok_{fid}"] = paste

    with st.expander("🔎 Deteksi GPS otomatis (opsional)"):
        hasil = deteksi_gps()
        if hasil:
            if st.session_state.get(f"gps_ok_{fid}") != hasil:
                st.session_state[f"lat_{fid}"] = hasil[0]
                st.session_state[f"lon_{fid}"] = hasil[1]
                st.session_state[f"gps_ok_{fid}"] = hasil
            st.success(f"Terdeteksi: {hasil[0]:.6f}, {hasil[1]:.6f}")
        else:
            st.caption("Tidak terdeteksi / izin ditolak. Gunakan kolom Tempel "
                       "Koordinat di atas (paling andal di HP).")

    if f"lat_{fid}" not in st.session_state:
        st.session_state[f"lat_{fid}"] = -6.200000
    if f"lon_{fid}" not in st.session_state:
        st.session_state[f"lon_{fid}"] = 106.816666
    c1, c2 = st.columns(2)
    with c1:
        lat = st.number_input("Latitude", step=0.000001, format="%.6f",
                              key=f"lat_{fid}")
    with c2:
        lon = st.number_input("Longitude", step=0.000001, format="%.6f",
                              key=f"lon_{fid}")
    st.markdown(f"🗺️ [Lihat titik ini di Google Maps]"
                f"(https://www.google.com/maps?q={lat},{lon})")
    lokasi_nama = st.text_input("🏢 Nama Lokasi / Tempat *",
                                placeholder="mis. Toko Emas Jaya, Pasar Anyar",
                                key=f"lok_{fid}")

    # ---------- 3. TRANSAKSI ----------
    st.subheader("3️⃣ Data Transaksi & Keuangan")
    jenis = st.selectbox("🏷️ Jenis Transaksi", JENIS_SEMUA, key=f"jenis_{fid}",
                         help="Beli/Jual Emas → masuk Stok Emas & Keuangan. "
                              "Pemasukan/Pengeluaran → Keuangan saja. "
                              "Transfer → pindah saldo antar rekening.")
    berat, harga, nominal = 0.0, 0, 0.0
    if jenis in ("Beli Emas", "Jual Emas"):
        c1, c2 = st.columns(2)
        with c1:
            berat = st.number_input("⚖️ Berat emas (gram)", min_value=0.0, step=0.01,
                                    format="%.2f", key=f"berat_{fid}")
        with c2:
            harga = st.number_input("💰 Harga per gram (Rp)", min_value=0, step=1000,
                                    format="%.0f", key=f"harga_{fid}")
        nominal = float(berat) * float(harga)
        if nominal:
            st.metric("💵 Total Nominal (otomatis)", format_rupiah(nominal))
        if st.checkbox("✏️ Nominal berbeda dari hitungan? (mis. ada biaya tambahan)",
                       key=f"ubah_{fid}"):
            nominal = st.number_input("💵 Total Nominal (Rp)", min_value=0.0,
                                      step=1000.0, format="%.0f", key=f"nom_{fid}")
    else:
        nominal = st.number_input("💵 Nominal (Rp)", min_value=0.0, step=1000.0,
                                  format="%.0f", key=f"nom_{fid}")

    if not daftar_rek:
        st.warning("⚠️ Belum ada rekening. Tambahkan dulu di 'Kelola Rekening' "
                   "di bagian bawah halaman ini.")
    rek_utama, rek_tujuan = "", ""
    if jenis == "Transfer Antar Rekening":
        c1, c2 = st.columns(2)
        with c1:
            rek_utama = st.selectbox("📤 Dari rekening (asal)", daftar_rek,
                                     key=f"reka_{fid}")
        with c2:
            rek_tujuan = st.selectbox("📥 Ke rekening (tujuan)", daftar_rek,
                                      key=f"rekb_{fid}")
    else:
        rek_utama = st.selectbox("🏦 Rekening / Kas yang terpengaruh", daftar_rek,
                                 key=f"rek_{fid}")

    # ---------- 4. TRACKING & BUKTI ----------
    st.subheader("4️⃣ Data Tracking & Bukti")
    aktivitas = st.text_area("📝 Aktivitas / Uraian",
                             placeholder="mis. COD pembelian emas 5 gram dari pelanggan",
                             key=f"akt_{fid}")
    foto_b64 = ""
    cf1, cf2 = st.columns(2)
    with cf1:
        foto_kamera = st.camera_input("📸 Ambil foto (kamera HP)")
    with cf2:
        foto_file = st.file_uploader("🖼️ atau unggah gambar", type=["jpg", "jpeg", "png"])
    berkas = foto_kamera or foto_file
    if berkas is not None:
        foto_b64 = kompres_foto(berkas.getvalue())
        if foto_b64:
            st.image(io.BytesIO(base64.b64decode(foto_b64)), width=240,
                     caption="Pratinjau foto (disimpan versi kecil)")

    # ---------- 5. RINGKASAN & SIMPAN ----------
    st.subheader("5️⃣ Ringkasan & Simpan")
    with st.container(border=True):
        r1, r2 = st.columns(2)
        with r1:
            st.markdown(f"""
- **Tanggal:** {tanggal.strftime('%d-%m-%Y')} ({NAMA_HARI[tanggal.weekday()]}) — {jam.strftime('%H:%M')}
- **Jenis:** {jenis}
- **Nominal:** {format_rupiah(nominal)}
- **Berat emas:** {format_gram(berat) if berat else '-'}
""")
        with r2:
            st.markdown(f"""
- **Rekening:** {rek_utama or '-'}{(' → ' + rek_tujuan) if rek_tujuan else ''}
- **Lokasi:** {lokasi_nama or '-'} ({lat:.5f}, {lon:.5f})
- **Aktivitas:** {aktivitas or '(otomatis dari jenis transaksi)'}
- **Foto:** {'ada' if foto_b64 else 'tidak ada'}
""")

    if st.button("💾 SIMPAN TRANSAKSI", type="primary", use_container_width=True,
                 key=f"simpan_{fid}"):
        masalah = []
        if not daftar_rek or not rek_utama:
            masalah.append("Rekening belum dipilih / belum ada (tambahkan di Kelola Rekening).")
        if nominal <= 0:
            masalah.append("Nominal harus lebih dari 0.")
        if jenis in ("Beli Emas", "Jual Emas"):
            if berat <= 0:
                masalah.append("Berat emas harus diisi.")
            if harga <= 0:
                masalah.append("Harga per gram harus diisi.")
        if jenis == "Transfer Antar Rekening" and rek_utama == rek_tujuan:
            masalah.append("Rekening asal dan tujuan tidak boleh sama.")
        if not str(lokasi_nama).strip():
            masalah.append("Nama lokasi wajib diisi.")

        if masalah:
            st.error("❗ Perbaiki dulu:\n- " + "\n- ".join(masalah))
        else:
            id_trx = (f"TRX-{tanggal:%Y%m%d}-{sekarang_wib():%H%M%S}-"
                      f"{uuid.uuid4().hex[:4].upper()}")
            arah = ("Debit" if jenis in JENIS_MASUK
                    else ("Kredit" if jenis in JENIS_KELUAR else "Transfer"))
            aktivitas_final = str(aktivitas).strip() or f"{jenis} — {lokasi_nama}"
            baris = [id_trx,
                     tanggal.strftime("%Y-%m-%d"),
                     NAMA_HARI[tanggal.weekday()],
                     jam.strftime("%H:%M"),
                     jenis,
                     float(berat),
                     int(harga),
                     int(round(float(nominal))),
                     rek_utama or "",
                     (rek_tujuan or "") if jenis == "Transfer Antar Rekening" else "",
                     arah,
                     str(lokasi_nama).strip(),
                     float(lat),
                     float(lon),
                     aktivitas_final,
                     st.session_state.get("username", ""),
                     sekarang_wib().strftime("%Y-%m-%d %H:%M:%S"),
                     foto_b64]
            try:
                tambah_baris("Transaksi", baris)
                st.session_state["pesan_sukses"] = id_trx
                st.session_state["form_id"] = fid + 1
                st.rerun()
            except Exception as e:
                st.error(f"Gagal menyimpan ke Google Sheets: {e}")


def halaman_input_excel(daftar_rek):
    if st.session_state.get("pesan_sukses_excel"):
        st.success(f"✅ {st.session_state['pesan_sukses_excel']} transaksi dari "
                   "Excel berhasil disimpan — semua menu sudah ter-update.")
        del st.session_state["pesan_sukses_excel"]

    st.subheader("📥 Input via Excel")
    st.caption("Unggah file Excel (.xlsx) / CSV berisi banyak transaksi sekaligus. "
               "Foto bukti tidak bisa diikutkan lewat Excel (tambahkan manual "
               "bila perlu). Rekening di dalam file yang belum terdaftar akan "
               "dideteksi otomatis dan bisa didaftarkan 1 klik.")
    if not daftar_rek:
        st.warning("⚠️ Belum ada rekening terdaftar — tambahkan dulu minimal 1 "
                   "rekening di 'Kelola Rekening' (bagian bawah halaman ini).")

    # ---------- LIHAT FORMAT & UNDUH TEMPLATE ----------
    with st.expander("📋 Lihat Format Tabel yang Harus Diisi (WAJIB dibaca dulu)"):
        contoh = pd.DataFrame([
            {"Tanggal": "2026-01-05", "Hari": "", "Jam": "10:30",
             "Jenis": "Beli Emas", "Berat (gram)": 5,
             "Harga per Gram (Rp)": 1350000, "Nominal (Rp)": "",
             "Rekening": "BCA", "Rekening Tujuan": "",
             "Nama Lokasi": "Toko Emas Jaya", "Latitude": -6.2146,
             "Longitude": 106.8451, "Aktivitas": "COD beli emas 5 gram"},
            {"Tanggal": "05/01/2026", "Hari": "", "Jam": "14:00",
             "Jenis": "Jual Emas", "Berat (gram)": 2.5,
             "Harga per Gram (Rp)": 1380000, "Nominal (Rp)": "",
             "Rekening": "Kas Toko", "Rekening Tujuan": "",
             "Nama Lokasi": "Pasar Anyar", "Latitude": "", "Longitude": "",
             "Aktivitas": ""},
            {"Tanggal": "06/01/2026", "Hari": "Selasa", "Jam": "",
             "Jenis": "Pengeluaran Operasional", "Berat (gram)": "",
             "Harga per Gram (Rp)": "", "Nominal (Rp)": 250000,
             "Rekening": "Kas Toko", "Rekening Tujuan": "",
             "Nama Lokasi": "SPBU", "Latitude": "", "Longitude": "",
             "Aktivitas": "Isi bensin operasional"},
            {"Tanggal": "07/01/2026", "Hari": "", "Jam": "09:00",
             "Jenis": "Transfer Antar Rekening", "Berat (gram)": "",
             "Harga per Gram (Rp)": "", "Nominal (Rp)": 5000000,
             "Rekening": "Kas Toko", "Rekening Tujuan": "BCA",
             "Nama Lokasi": "Bank BCA", "Latitude": "", "Longitude": "",
             "Aktivitas": "Setor hasil penjualan ke bank"},
        ], columns=KOLOM_EXCEL)
        st.markdown("**Contoh format (4 baris contoh — hapus/ganti saat mengisi):**")
        st.dataframe(contoh, use_container_width=True, hide_index=True)

        st.markdown("**Aturan pengisian tiap kolom:**")
        aturan = pd.DataFrame({
            "Kolom": ["Tanggal", "Hari", "Jam", "Jenis", "Berat (gram)",
                      "Harga per Gram (Rp)", "Nominal (Rp)", "Rekening",
                      "Rekening Tujuan", "Nama Lokasi",
                      "Latitude / Longitude", "Aktivitas"],
            "Wajib?": ["✅ Wajib", "Opsional", "Opsional", "✅ Wajib",
                       "Wajib utk emas", "Wajib utk emas",
                       "✅ Wajib (otomatis utk emas)", "✅ Wajib",
                       "Wajib utk transfer", "✅ Wajib", "Opsional",
                       "Opsional"],
            "Keterangan": [
                "Format: 2026-01-05 atau 05/01/2026 (tanggal/bulan/tahun)",
                "Kosongkan — dihitung otomatis dari tanggal",
                "Format 24 jam, mis. 13:45. Kosong = diisi 12:00",
                "Beli Emas / Jual Emas / Pemasukan Lain / Pengeluaran Operasional / Transfer Antar Rekening",
                "Angka polos, mis. 2,5 — tanpa titik ribuan",
                "Angka polos, mis. 1350000",
                "Total Rp. Untuk emas boleh kosong → otomatis Berat × Harga",
                "Harus sama dengan rekening terdaftar — rekening baru di file "
                "dideteksi otomatis setelah unggah & bisa didaftarkan 1 klik",
                "Diisi hanya jika Jenis = Transfer Antar Rekening",
                "Nama tempat transaksi (isi tanda - jika tidak ada)",
                "Koordinat GPS, boleh kosong (mis. -6.2146 dan 106.8451)",
                "Uraian/catatan; kosong = dibuat otomatis dari Jenis & Lokasi"],
        })
        st.dataframe(aturan, use_container_width=True, hide_index=True)

        k1, k2 = st.columns(2)
        with k1:
            try:
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    contoh.to_excel(w, index=False, sheet_name="Transaksi")
                st.download_button(
                    "⬇️ Unduh Template (.xlsx)", buf.getvalue(),
                    "template_transaksi.xlsx",
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet")
            except Exception:
                st.caption("Template .xlsx gagal dibuat (openpyxl belum "
                           "terpasang?) — pakai tombol CSV di samping.")
        with k2:
            st.download_button("⬇️ Unduh Template (.csv)",
                               contoh.to_csv(index=False).encode("utf-8-sig"),
                               "template_transaksi.csv", "text/csv")

    # ---------- UNGGAH FILE ----------
    st.subheader("📤 Unggah File Excel")
    if st.session_state.pop("reset_uploader", False):
        st.session_state["excel_uploader"] = None
    berkas = st.file_uploader("Pilih file .xlsx atau .csv lalu isi sesuai format",
                              type=["xlsx", "csv"], key="excel_uploader")
    if berkas is not None:
        df_raw = None
        try:
            if berkas.name.lower().endswith(".csv"):
                df_raw = pd.read_csv(berkas, keep_default_na=False, dtype=str)
            else:
                df_raw = pd.read_excel(berkas)
        except Exception as e:
            st.error(f"File tidak dapat dibaca: {e} — pastikan file berformat "
                     ".xlsx atau .csv (jika file .xls lama, buka di Excel lalu "
                     "Save As → .xlsx).")
        if df_raw is not None:
            # ====== TAHAP 1: deteksi rekening yang belum terdaftar ======
            terdaftar_lwr = {str(r).strip().lower() for r in daftar_rek}
            petamax = {}
            for kolom_asli in df_raw.columns:
                k = _norm_header(kolom_asli)
                target = SINONIM_EXCEL.get(k)
                if target and target not in petamax:
                    petamax[target] = kolom_asli
            nama_rek_file = set()
            for target in ("Rekening", "Rekening Tujuan"):
                if target in petamax:
                    for v in df_raw[petamax[target]].dropna().astype(str):
                        v = v.strip()
                        if v and v.lower() not in ("", "nan", "none", "nat", "-"):
                            nama_rek_file.add(v)
            belum = sorted(n for n in nama_rek_file
                           if n.strip().lower() not in terdaftar_lwr)

            if belum:
                st.warning(f"⚠️ File memuat **{len(belum)} rekening yang belum "
                           f"terdaftar di database**: {', '.join(belum)}.\n\n"
                           "Transaksi yang memakai rekening tersebut akan "
                           "ditolak validasi. Daftarkan dulu di bawah — setelah "
                           "itu file divalidasi ulang otomatis.")
                st.markdown("**➕ Daftarkan rekening yang belum ada:**")
                kol = st.columns(3)
                saldo_baru = {}
                for i, n in enumerate(belum):
                    with kol[i % 3]:
                        saldo_baru[n] = st.number_input(
                            f"Saldo awal '{n}' (Rp)", min_value=0.0,
                            step=1_000_000.0, format="%.0f",
                            key=f"daftar_rek_baru_{i}")
                if st.button("➕ DAFTARKAN SEMUA REKENING DI ATAS",
                             type="primary"):
                    try:
                        sh = buat_koneksi()
                        wr = sh.worksheet("Rekening")
                        for n in belum:
                            nl = n.lower()
                            if any(b in nl for b in ("bca", "mandiri", "bri",
                                                     "bni", "bank", "btn",
                                                     "cimb", "danamon",
                                                     "citibank")):
                                jenis = "Bank"
                            elif any(e in nl for e in ("gopay", "ovo", "dana",
                                                       "shopee", "linkaja")):
                                jenis = "E-Wallet"
                            else:
                                jenis = "Kas Tunai"
                            wr.append_row([n, jenis, "", "",
                                           float(saldo_baru.get(n, 0.0)),
                                           "Rekening dari unggahan Excel"])
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal mendaftarkan rekening: {e}")
                st.caption("ℹ️ Saldo awal bisa diisi bebas dan dapat diubah "
                           "kapan saja di Google Sheets (tab Rekening). Setelah "
                           "didaftarkan, Anda TIDAK perlu memilih file lagi — "
                           "validasi berjalan sendiri.")
                st.markdown("---")
                return   # validasi penuh ditunda sampai rekening terdaftar

            # ====== TAHAP 2: validasi penuh ======
            baris_siap, errors, pratinjau = proses_file_excel(df_raw, daftar_rek)
            if errors:
                st.error(f"❌ Ada {len(errors)} baris bermasalah — data BELUM "
                         "disimpan. Perbaiki di Excel lalu unggah ulang:")
                st.dataframe(pd.DataFrame(errors), use_container_width=True,
                             hide_index=True)
                if baris_siap:
                    st.caption(f"({len(baris_siap)} baris lain sudah valid — "
                               "mereka baru disimpan SETELAH semua baris "
                               "diperbaiki.)")
            elif not baris_siap:
                st.warning("File terbaca, tetapi tidak berisi baris data.")
            else:
                st.success(f"✅ Semua {len(baris_siap)} baris valid & siap "
                           "disimpan. Periksa pratinjau, lalu klik tombol simpan.")
                with st.expander(f"👀 Pratinjau {len(baris_siap)} transaksi",
                                 expanded=True):
                    st.dataframe(pratinjau, use_container_width=True,
                                 hide_index=True)
                if st.button(f"💾 SIMPAN SEMUA ({len(baris_siap)} transaksi)",
                             type="primary", use_container_width=True):
                    try:
                        sh = buat_koneksi()
                        sh.worksheet("Transaksi").append_rows(baris_siap)
                        st.cache_data.clear()
                        st.session_state["pesan_sukses_excel"] = len(baris_siap)
                        st.session_state["reset_uploader"] = True
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal menyimpan ke Google Sheets: {e}")


# =====================================================================
#  HALAMAN: PENGATURAN (KHUSUS ADMIN)
# =====================================================================
def halaman_setting():
    st.header("⚙️ Pengaturan")
    username = st.session_state.get("username", "")
    peran = st.session_state.get("role", "")
    st.caption(f"Login sebagai: **{username}** · Jenis: "
               f"**{'Admin' if peran == 'admin' else 'Reguler'}**")

    df_u = ambil_users()
    aku = df_u[df_u["Username"].astype(str).str.strip().str.lower()
               == str(username).strip().lower()]
    if aku.empty:
        st.error("Data pengguna Anda tidak ditemukan di database.")
        return
    aku = aku.iloc[0]

    # ---------- 1. PROFIL SAYA ----------
    st.subheader("👤 Profil Saya")
    with st.form("form_profil", border=True):
        c1, c2 = st.columns(2)
        with c1:
            nama_baru = st.text_input("Nama", value=str(aku["Nama"] or ""))
        with c2:
            st.text_input("Username", value=str(username), disabled=True,
                          help="Username tidak dapat diubah (dipakai sebagai "
                               "identitas pada data transaksi).")
        email_baru = st.text_input("Email",
                                   value=str(aku["Email"] or ""),
                                   placeholder="nama@email.com")
        simpan_profil = st.form_submit_button("💾 Simpan Profil", type="primary")
    if simpan_profil:
        email = str(email_baru).strip()
        masalah = []
        if not str(nama_baru).strip():
            masalah.append("Nama tidak boleh kosong.")
        if email and ("@" not in email or "." not in email.split("@")[-1]):
            masalah.append("Format email tidak valid.")
        if masalah:
            st.error("❗ " + " | ".join(masalah))
        else:
            try:
                sh = buat_koneksi()
                wu = sh.worksheet("Users")
                baris = cari_baris_user(username)
                if baris:
                    wu.update_acell(f"C{baris}", str(nama_baru).strip())
                    wu.update_acell(f"E{baris}", email)
                    st.session_state["nama"] = str(nama_baru).strip()
                    st.cache_data.clear()
                    st.success("✅ Profil berhasil disimpan.")
                else:
                    st.error("Baris pengguna tidak ditemukan.")
            except Exception as e:
                st.error(f"Gagal menyimpan profil: {e}")

    # ---------- 2. GANTI PASSWORD SAYA ----------
    st.subheader("🔑 Ganti Password Saya")
    with st.form("form_ganti_pw", border=True, clear_on_submit=True):
        pw_lama = st.text_input("Password Lama", type="password")
        pw_baru = st.text_input("Password Baru", type="password")
        pw_konf = st.text_input("Konfirmasi Password Baru", type="password")
        ganti_pw = st.form_submit_button("🔁 Ganti Password", type="primary")
    if ganti_pw:
        masalah = []
        if hash_password(pw_lama) != str(aku["PasswordHash"]).strip():
            masalah.append("Password lama salah.")
        if len(pw_baru) < 4:
            masalah.append("Password baru minimal 4 karakter.")
        if pw_baru != pw_konf:
            masalah.append("Konfirmasi password tidak sama dengan password baru.")
        if pw_lama and pw_baru == pw_lama:
            masalah.append("Password baru harus berbeda dari password lama.")
        if masalah:
            st.error("❗ " + " | ".join(masalah))
        else:
            try:
                sh = buat_koneksi()
                wu = sh.worksheet("Users")
                baris = cari_baris_user(username)
                if baris:
                    wu.update_acell(f"B{baris}", hash_password(pw_baru))
                    st.cache_data.clear()
                    st.success("✅ Password berhasil diganti. Gunakan password "
                               "baru saat login berikutnya.")
                else:
                    st.error("Baris pengguna tidak ditemukan.")
            except Exception as e:
                st.error(f"Gagal mengganti password: {e}")

    # ---------- 3. TAMBAH USER BARU ----------
    st.subheader("➕ Tambah User Baru")
    with st.form("form_user_baru", border=True, clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            u_nama = st.text_input("Nama *", placeholder="Nama lengkap")
        with c2:
            u_user = st.text_input("Username *", placeholder="tanpa spasi")
        c3, c4 = st.columns(2)
        with c3:
            u_pw = st.text_input("New Password *", type="password")
        with c4:
            u_pw2 = st.text_input("Confirm New Password *", type="password")
        c5, c6 = st.columns(2)
        with c5:
            u_role = st.selectbox(
                "Jenis Username *", ["reguler", "admin"],
                help="admin = bisa Input Data & Pengaturan · "
                     "reguler = hanya melihat laporan (Dashboard, Keuangan, "
                     "Rekening, Stok Emas, Tracking Admin)")
        with c6:
            u_email = st.text_input("Email (opsional)",
                                    placeholder="nama@email.com")
        tambah_user = st.form_submit_button("➕ Tambah User", type="primary")
    if tambah_user:
        email = str(u_email).strip()
        masalah = []
        if not str(u_nama).strip():
            masalah.append("Nama wajib diisi.")
        if not str(u_user).strip():
            masalah.append("Username wajib diisi.")
        elif " " in str(u_user):
            masalah.append("Username tidak boleh mengandung spasi.")
        elif (str(u_user).strip().lower() in
              df_u["Username"].astype(str).str.strip().str.lower().tolist()):
            masalah.append("Username sudah dipakai.")
        if len(u_pw) < 4:
            masalah.append("Password minimal 4 karakter.")
        if u_pw != u_pw2:
            masalah.append("Confirm New Password tidak sama dengan New Password.")
        if email and ("@" not in email or "." not in email.split("@")[-1]):
            masalah.append("Format email tidak valid.")
        if masalah:
            st.error("❗ Perbaiki dulu:\n- " + "\n- ".join(masalah))
        else:
            try:
                tambah_baris("Users",
                             [str(u_user).strip().lower(),
                              hash_password(u_pw),
                              str(u_nama).strip(),
                              u_role,
                              email])
                st.success(f"✅ User **{u_user}** ({u_role}) berhasil ditambahkan. "
                           "Login dengan username & password tersebut.")
            except Exception as e:
                st.error(f"Gagal menambah user: {e}")

    # ---------- 4. DAFTAR USER ----------
    st.subheader("👥 Daftar Pengguna")
    df_u2 = ambil_users()
    if not df_u2.empty:
        tampil = df_u2[["Username", "Nama", "Role", "Email"]].copy()
        st.dataframe(tampil, use_container_width=True, hide_index=True)

    with st.expander("🔁 Reset Password User Lain (jika ada yang lupa password)"):
        with st.form("form_reset_pw", border=True, clear_on_submit=True):
            target = st.selectbox(
                "Pilih user",
                df_u2["Username"].tolist() if not df_u2.empty else ["-"])
            r_pw = st.text_input("Password Baru", type="password")
            r_pw2 = st.text_input("Konfirmasi Password Baru", type="password")
            reset = st.form_submit_button("🔁 Reset Password")
        if reset:
            masalah = []
            if not target or target == "-":
                masalah.append("Pilih user terlebih dahulu.")
            if len(r_pw) < 4:
                masalah.append("Password minimal 4 karakter.")
            if r_pw != r_pw2:
                masalah.append("Konfirmasi password tidak sama.")
            if masalah:
                st.error("❗ " + " | ".join(masalah))
            else:
                try:
                    sh = buat_koneksi()
                    wu = sh.worksheet("Users")
                    baris = cari_baris_user(target)
                    if baris:
                        wu.update_acell(f"B{baris}", hash_password(r_pw))
                        st.cache_data.clear()
                        st.success(f"✅ Password user **{target}** sudah di-reset.")
                    else:
                        st.error("User tidak ditemukan.")
                except Exception as e:
                    st.error(f"Gagal reset password: {e}")

    # ---------- 5. HAPUS USER ----------
    with st.expander("🗑️ Hapus User"):
        username_saya = str(username).strip().lower()
        df_u3 = ambil_users()
        # Akun yang sedang login TIDAK boleh dihapus (supaya tidak terkunci)
        kandidat = df_u3[df_u3["Username"].astype(str).str.strip().str.lower()
                         != username_saya]
        if kandidat.empty:
            st.caption("Tidak ada user lain yang bisa dihapus. Akun Anda sendiri "
                       "tidak dapat dihapus dari dalam aplikasi — agar Anda tidak "
                       "terkunci dari sistem. Bila ingin akun Anda dihapus, "
                       "minta admin lain yang melakukannya.")
        else:
            jumlah_admin = int((df_u3["Role"].astype(str).str.strip().str.lower()
                                == "admin").sum())
            target_hapus = st.selectbox("Pilih user yang akan dihapus",
                                        kandidat["Username"].tolist())
            info = kandidat[kandidat["Username"] == target_hapus].iloc[0]
            st.markdown(f"""
- **Nama:** {info['Nama'] or '-'}
- **Username:** {info['Username']}
- **Jenis:** {info['Role']}
- **Email:** {info['Email'] or '-'}
""")
            # tampilkan jumlah riwayat transaksi milik user ini
            try:
                df_trx = ambil_transaksi()
                n_trx = int((df_trx["Admin"].astype(str).str.strip().str.lower()
                             == str(target_hapus).strip().lower()).sum())
            except Exception:
                n_trx = 0
            st.caption(f"📦 Riwayat transaksi tercatat atas nama user ini: "
                       f"**{n_trx} transaksi** — riwayat transaksi TIDAK ikut "
                       f"terhapus (tetap tersimpan sebagai data historis).")

            if (str(info["Role"]).strip().lower() == "admin"
                    and jumlah_admin <= 1):
                st.warning("⚠️ User ini adalah satu-satunya admin yang tersisa. "
                           "Tambahkan admin baru terlebih dahulu (bagian "
                           "➕ Tambah User Baru di atas) sebelum menghapus "
                           "user ini, agar aplikasi tidak kehilangan semua admin.")
            else:
                yakin = st.checkbox("Saya yakin — hapus user ini permanen")
                if st.button("🗑️ Hapus User Sekarang", disabled=not yakin):
                    try:
                        baris = cari_baris_user(target_hapus)
                        if baris is None:
                            st.error("User tidak ditemukan di Google Sheets "
                                     "(mungkin sudah dihapus sebelumnya). "
                                     "Muat ulang halaman.")
                        else:
                            sh = buat_koneksi()
                            wu = sh.worksheet("Users")
                            wu.delete_rows(baris)
                            st.cache_data.clear()
                            st.success(f"✅ User **{target_hapus}** berhasil "
                                       "dihapus. User tersebut tidak bisa login "
                                       "lagi; riwayat transaksinya tetap "
                                       "tersimpan.")
                            st.rerun()
                    except Exception as e:
                        st.error(f"Gagal menghapus user: {e}")

    st.caption("👤 Akun yang sedang login tidak bisa dihapus dari dalam aplikasi. "
               "User yang dihapus tidak dapat login lagi, tetapi riwayat "
               "transaksinya tetap tersimpan sebagai data historis.")


# =====================================================================
#  PROGRAM UTAMA
# =====================================================================
def _error_jaringan(e):
    """Mendeteksi apakah error disebabkan masalah koneksi internet."""
    pesan = f"{type(e).__name__}: {e}"
    tanda = ("getaddrinfo", "connectionerror", "max retries exceeded",
             "timed out", "connection reset", "connection aborted",
             "newconnectionerror", "proxyerror", "sslerror",
             "failed to resolve", "temporarily unavailable")
    return any(t in pesan.lower() for t in tanda)


def _tampilkan_error_jaringan():
    st.error("🌐 **Koneksi ke Google Sheets terputus.**\n\n"
             "Laptop Anda saat itu tidak dapat menghubungi server Google "
             "(internet putus sesaat, berpindah jaringan, VPN aktif, atau "
             "gangguan DNS). Tenang — **data Anda aman** di Google Sheets.\n\n"
             "**Langkah:**\n"
             "1. Pastikan internet aktif (buka google.com di browser)\n"
             "2. Matikan VPN/proxy bila sedang aktif\n"
             "3. Tunggu ±30 detik, lalu tekan tombol di bawah")
    if st.button("🔄 Coba Lagi Sekarang", type="primary"):
        st.cache_data.clear()
        st.rerun()


@st.cache_resource(show_spinner=False)
def setup_database():
    """Menjalankan persiapan database HANYA SEKALI per hidup aplikasi
    (bukan sekali per pengunjung) — pembukaan app jadi jauh lebih cepat,
    terutama di Streamlit Cloud."""
    pastikan_setup()
    return True

def main():
    if "GANTI" in SPREADSHEET_ID:
        st.error("⚠️ Buka file app.py, lalu ganti teks "
                 "GANTI_DENGAN_ID_SPREADSHEET_ANDA dengan ID Google Sheets "
                 "Anda (lihat petunjuk pemasangan).")
        st.stop()

    if not st.session_state.get("logged_in", False):
        if not st.session_state.get("setup_ok", False):
            try:
                with st.spinner("Menyiapkan database Google Sheets "
                               "(sekali saja)..."):
                    setup_database()
                st.session_state["setup_ok"] = True
            except Exception as e:
                if _error_jaringan(e):
                    _tampilkan_error_jaringan()
                    st.stop()
                st.error(f"⚠️ Gagal menyiapkan database: {e}")
                st.info("Periksa 3 hal: (1) SPREADSHEET_ID di app.py sudah "
                        "benar; (2) file service_account.json ada di folder "
                        "yang sama dengan app.py; (3) spreadsheet sudah "
                        "di-SHARE ke email service account sebagai Editor.")
                st.stop()
        halaman_login()
    else:
        try:
            menu_utama()
        except Exception as e:
            if _error_jaringan(e):
                _tampilkan_error_jaringan()
            else:
                raise e


main()