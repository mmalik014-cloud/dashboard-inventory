import io
import math
import os
from datetime import datetime
import openpyxl
import pandas as pd
import plotly.express as px
from scipy.stats import norm
import streamlit as st

# ============================================================
# 1. KONFIGURASI HALAMAN DASHBOARD & HEADER UTAMA
# ============================================================
st.set_page_config(
    page_title="Smart Inventory Systems",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <div style="padding-bottom: 10px;">
        <h2 style="font-size: 30px; font-weight: 700; margin-bottom: 4px;">
            📦 INTEGRATED INVENTORY CONTROL & REAL-TIME MONITORING SYSTEM
        </h2>
        <p style="font-size: 18px; color: #A0AAB8; font-weight: 500; margin-top: 0;">
            Alat Pendukung Keputusan Persediaan Probabilistik Berbasis Algoritma Hadley–Whitin & Otomasi ABC Analysis
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

FILE_RIWAYAT = "riwayat_transaksi.xlsx"

# ============================================================
# 2. FUNGSI FORMATTING & PARSING ANGKA AMAN
# ============================================================
def bersihkan_angka(nilai):
    """Memastikan angka desimal koma (Indonesia) atau float dibaca dengan benar."""
    if pd.isna(nilai) or nilai is None:
        return 0.0
    if isinstance(nilai, (int, float)):
        return float(nilai)
    
    val_str = str(nilai).strip()
    if "," in val_str and "." not in val_str:
        val_str = val_str.replace(",", ".")
    elif "," in val_str and "." in val_str:
        val_str = val_str.replace(".", "").replace(",", ".")
    
    try:
        return float(val_str)
    except:
        return 0.0

def format_indonesia_kg(nilai):
    if nilai == "-" or pd.isna(nilai) or nilai is None:
        return "-"
    try:
        val = float(nilai)
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " kg"
    except:
        return str(nilai)

def format_rupiah(nilai):
    try:
        val = float(nilai)
        return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(nilai)

def muat_riwayat():
    if os.path.exists(FILE_RIWAYAT):
        try:
            df_log = pd.read_excel(FILE_RIWAYAT)
            df_log["Waktu"] = pd.to_datetime(df_log["Waktu"])
            return df_log
        except:
            pass
    return pd.DataFrame(columns=["Waktu", "Tanggal", "Bulan", "Bahan Baku", "Aktivitas", "Jumlah (kg)"])

def catat_transaksi(bahan, aktivitas, jumlah):
    df_log = muat_riwayat()
    sekarang = datetime.now()
    data_baru = pd.DataFrame([{
        "Waktu": sekarang,
        "Tanggal": sekarang.strftime("%Y-%m-%d"),
        "Bulan": sekarang.strftime("%Y-%m"),
        "Bahan Baku": bahan,
        "Aktivitas": aktivitas,
        "Jumlah (kg)": jumlah,
    }])
    df_log = pd.concat([df_log, data_baru], ignore_index=True)
    df_log.to_excel(FILE_RIWAYAT, index=False)

def buat_template_excel():
    """Membuat binary Excel template kosongan beserta tabel keterangan."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "data verifikasi"

    # Header Kolom Utama
    headers = ["Jenis Barang", "D", "Sigma", "L", "A", "h", "Cu", "pi"]
    ws.append(headers)

    # Keterangan
    keterangan = [
        ["", "", "", "", "", "", "", "", "", "", "Keterangan", ""],
        ["", "", "", "", "", "", "", "", "", "", "D", "Demand/Permintaan"],
        ["", "", "", "", "", "", "", "", "", "", "Sigma", "Deviasi"],
        ["", "", "", "", "", "", "", "", "", "", "L", "Akar Leadtime"],
        ["", "", "", "", "", "", "", "", "", "", "A", "Biaya Per pesan"],
        ["", "", "", "", "", "", "", "", "", "", "h", "Biaya Simpan"],
        ["", "", "", "", "", "", "", "", "", "", "Cu", "Biaya Backorder/Kekurangan"],
        ["", "", "", "", "", "", "", "", "", "", "Pi", "Harga Bahan Baku"],
    ]

    for row in keterangan:
        ws.append(row)

    wb.save(output)
    return output.getvalue()

# Session State Storage
if "data_gudang" not in st.session_state:
    st.session_state["data_gudang"] = None
if "stok_realtime" not in st.session_state:
    st.session_state["stok_realtime"] = {}

# ============================================================
# 3. SIDEBAR UPLOAD & UNDUH TEMPLATE
# ============================================================
st.sidebar.header("⚙️ Langkah 1: Upload Master Data")

# Tombol Unduh Template Otomatis
excel_template = buat_template_excel()
st.sidebar.download_button(
    label="📥 Unduh Template Excel Otomatis",
    data=excel_template,
    file_name="Template_Data_Kosongan.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded_file = st.sidebar.file_uploader("Upload File Template Excel/CSV", type=["csv", "xlsx"])

max_iter = st.sidebar.slider("Batas Maksimum Iterasi Hadley-Whitin", min_value=10, max_value=100, value=50)

# ==============================================================================
# 4. PEMPROSESAN DATA & OTOMASI ANALISIS ABC + HADLEY-WHITIN
# ==============================================================================
if uploaded_file is not None and st.session_state["data_gudang"] is None:
    try:
        # Membaca data dari file template
        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file, delimiter=";")
        else:
            df_raw = pd.read_excel(uploaded_file)

        df_raw.columns = df_raw.columns.str.strip()

        # Membersihkan baris kosong & menentukan nama kolom bahan baku
        col_nama = "Jenis Barang" if "Jenis Barang" in df_raw.columns else "Bahan Baku"
        df = df_raw.dropna(subset=[col_nama]).copy()

        df["Bahan_Nama"] = df[col_nama].astype(str).str.strip()
        df["D_num"] = df["D"].apply(bersihkan_angka)
        df["Sigma_num"] = df["Sigma"].apply(bersihkan_angka)

        # OTOMATIS KONVERSI LEAD TIME (HARI KE TAHUN)
        df["L_hari"] = df["L"].apply(bersihkan_angka)
        df["L_num"] = df["L_hari"].apply(lambda val: val / 365.0 if val > 1.0 else val)

        df["A_num"] = df["A"].apply(bersihkan_angka)
        df["h_num"] = df["h"].apply(bersihkan_angka)
        
        # Penanganan biaya Cu & pi (harga)
        df["Cu_num"] = df["Cu"].apply(bersihkan_angka) if "Cu" in df_raw.columns else 0.0

        if "pi" in df_raw.columns:
            df["pi_num"] = df["pi"].apply(bersihkan_angka)
        else:
            df["pi_num"] = df["Cu_num"]

        # ----------------------------------------------------
        # OTOMASI KLASIFIKASI ABC OLEH SISTEM
        # ----------------------------------------------------
        # 1. Hitung Nilai Penyerapan Modal (Investasi per Barang)
        df["Nilai_Penyerapan"] = df["D_num"] * df["pi_num"]
        
        # 2. Buat objek urutan terpisah berdasarkan nilai penyerapan
        df_abc = df.sort_values(by="Nilai_Penyerapan", ascending=False).copy()
        total_penyerapan = df_abc["Nilai_Penyerapan"].sum()
        
        if total_penyerapan > 0:
            df_abc["Persen"] = (df_abc["Nilai_Penyerapan"] / total_penyerapan) * 100
            df_abc["Kumulatif"] = df_abc["Persen"].cumsum()
        else:
            df_abc["Kumulatif"] = 0.0

	# 3. Penentuan Otomatis Kategori ABC secara Spesifik (A, B, C)
        def tentukan_kategori_otomatis(kumulatif_persen):
            if kumulatif_persen <= 80.0:
                return "A"
            elif kumulatif_persen <= 95.0:
                return "B"
            else:
                return "C"

        # Map hasil kategori ke DataFrame utama
        kat_mapping = dict(zip(df_abc["Bahan_Nama"], df_abc["Kumulatif"].apply(tentukan_kategori_otomatis)))
        df["Kategori"] = df["Bahan_Nama"].map(kat_mapping)

        # ----------------------------------------------------
        # ITERASI HADLEY-WHITIN PER ITEM
        # ----------------------------------------------------
        hasil_perhitungan = []
        df_log_ada = muat_riwayat()

        for idx, row in df.iterrows():
            bahan = row["Bahan_Nama"]
            D = row["D_num"]
            sigma = row["Sigma_num"]
            L = row["L_num"]
            A = row["A_num"]
            h = row["h_num"]
            Cu = row["Cu_num"]
            kategori = row["Kategori"]

            metode = "Continuous Review (s,S)" if kategori == "A" else "Continuous Review (s,Q)"

            sigma_L = sigma * math.sqrt(L) if sigma > 0 and L > 0 else 1.0
            demand_L = D * L
            
            # Tebakan awal EOQ
            q_prev = math.sqrt((2 * A * D) / h) if h > 0 else 1.0
            r_prev = demand_L

            # Iterasi Konvergensi Hadley-Whitin
            for i in range(max_iter):
                alpha = (h * q_prev) / (Cu * D) if (Cu * D) > 0 else 0.05
                alpha = min(max(alpha, 1e-6), 1 - 1e-6)
                
                Z = norm.ppf(1 - alpha)
                N = sigma_L * (norm.pdf(Z) - Z * (1 - norm.cdf(Z)))
                
                q_new = math.sqrt((2 * D * (A + Cu * N)) / h) if h > 0 else 1.0
                r_new = demand_L + (Z * sigma_L)

                if abs(r_new - r_prev) < 0.0001:
                    q_prev, r_prev = q_new, r_new
                    break
                q_prev, r_prev = q_new, r_new

            SS_final = max(0.0, r_prev - demand_L)
            
            # S_max hanya berlaku untuk Kategori A (s,S)
            S_final = round(q_prev + r_prev, 2) if kategori == "A" else None

            # Kalkulasi Biaya Total (OT)
            OP = (A * D) / q_prev if q_prev > 0 else 0
            OS = h * ((q_prev / 2) + SS_final)
            
            alpha_f = min(max((h * q_prev) / (Cu * D if Cu * D > 0 else 1), 1e-6), 1 - 1e-6)
            Z_f = norm.ppf(1 - alpha_f)
            N_f = sigma_L * (norm.pdf(Z_f) - Z_f * (1 - norm.cdf(Z_f)))
            OK = Cu * (D / q_prev if q_prev > 0 else 0) * N_f

            OT_final = OP + OS + OK

            # Inisialisasi Stok Realtime
            stok_awal = round(q_prev + r_prev, 2) if kategori == "A" else round(q_prev + SS_final, 2)

            if not df_log_ada.empty:
                log_bahan = df_log_ada[df_log_ada["Bahan Baku"] == bahan]
                masuk = log_bahan[log_bahan["Aktivitas"] == "Barang Datang (Stok Masuk)"]["Jumlah (kg)"].sum()
                keluar = log_bahan[log_bahan["Aktivitas"] == "Diambil Produksi (Stok Keluar)"]["Jumlah (kg)"].sum()
                stok_awal = stok_awal + masuk - keluar

            st.session_state["stok_realtime"][bahan] = float(stok_awal)

            hasil_perhitungan.append({
                "Bahan Baku": bahan,
                "Kategori": kategori,
                "Metode": metode,
                "Q_opt (kg)": round(q_prev, 2),
                "s_opt (kg)": round(r_prev, 2),
                "S_max (kg)": S_final,
                "SS (kg)": round(SS_final, 2),
                "Total Cost (OT)": round(OT_final, 2),
            })

        st.session_state["data_gudang"] = pd.DataFrame(hasil_perhitungan)
    except Exception as e:
        st.error(f"Gagal memproses data template Excel: {e}")

# ============================================================
# 5. TAMPILAN DASHBOARD
# ============================================================
if st.session_state["data_gudang"] is not None:
    df_hasil = st.session_state["data_gudang"].copy()
    df_hasil["Stok Saat Ini (kg)"] = df_hasil["Bahan Baku"].map(st.session_state["stok_realtime"])

    def cek_status(row):
        if row["Stok Saat Ini (kg)"] <= row["s_opt (kg)"]:
            return "🚨 HARUS REORDER!"
        return "✅ Stok Aman"

    df_hasil["Status Gudang"] = df_hasil.apply(cek_status, axis=1)

    # PANEL PERINGATAN REORDER
    item_reorder = df_hasil[df_hasil["Status Gudang"] == "🚨 HARUS REORDER!"]
    if not item_reorder.empty:
        st.error("### ⚠️ PERINGATAN REORDER DI GUDANG!")
        for idx, row in item_reorder.iterrows():
            stok_sekarang = row["Stok Saat Ini (kg)"]
            if row["Kategori"] == "A":
                target_S = row["S_max (kg)"]
                jumlah_pesan = max(round(target_S - stok_sekarang, 2), row["Q_opt (kg)"])
                pesan = f"👉 **Rekomendasi (s,S):** Pesan **{format_indonesia_kg(jumlah_pesan)}** untuk mencapai Target Maksimum ($S_{{\\text{{max}}}}$ = **{format_indonesia_kg(target_S)}**)."
            else:
                pesan = f"👉 **Rekomendasi (s,Q):** Pesan lot tetap ($Q_{{\\text{{opt}}}}$) sebesar **{format_indonesia_kg(row['Q_opt (kg)'])}**."

            st.warning(f"**{row['Bahan Baku']}** Kritis! Stok **{format_indonesia_kg(stok_sekarang)}** ≤ Batas Reorder **{format_indonesia_kg(row['s_opt (kg)'])}**.\n\n{pesan}")
    else:
        st.success("### ✅ Semua stok persediaan dalam kondisi aman.")

    st.markdown("---")

    # PANEL TRANSAKSI MUTASI
    st.markdown("### 📥 📤 Panel Transaksi Gudang")
    c1, c2, c3 = st.columns(3)
    with c1:
        pilih_bahan = st.selectbox("Pilih Jenis Barang:", df_hasil["Bahan Baku"].tolist())
    with c2:
        jenis_transaksi = st.radio("Aktivitas:", ["Barang Datang (Stok Masuk)", "Diambil Produksi (Stok Keluar)"])
    with c3:
        jumlah_mutasi = st.number_input("Jumlah (kg):", min_value=0.0, step=100.0)
        submit_button = st.button("Simpan Transaksi 💾")

    if submit_button and jumlah_mutasi > 0:
        if jenis_transaksi == "Barang Datang (Stok Masuk)":
            catat_transaksi(pilih_bahan, jenis_transaksi, jumlah_mutasi)
            st.session_state["stok_realtime"][pilih_bahan] += jumlah_mutasi
            st.toast("Berhasil mencatat stok masuk!", icon="📥")
        else:
            catat_transaksi(pilih_bahan, jenis_transaksi, jumlah_mutasi)
            st.session_state["stok_realtime"][pilih_bahan] -= jumlah_mutasi
            st.toast("Berhasil mencatat stok keluar!", icon="📤")
        st.rerun()

    st.markdown("---")

    # TAB MONITORING & TABEL HASIL VERIFIKASI
    tab1, tab2, tab3 = st.tabs(["📊 Live Monitoring & Chart", "📜 Laporan Riwayat Keluar-Masuk", "⚙️ Reset Data"])

    with tab1:
        st.markdown("#### Hasil Perhitungan & Status Gudang Real-Time 📝")
        df_tampilan = df_hasil.copy()
        df_tampilan["Stok Saat Ini"] = df_tampilan["Stok Saat Ini (kg)"].apply(format_indonesia_kg)
        df_tampilan["s_opt (Batas Aman/r)"] = df_tampilan["s_opt (kg)"].apply(format_indonesia_kg)
        df_tampilan["Q_opt (Pemesanan)"] = df_tampilan["Q_opt (kg)"].apply(format_indonesia_kg)
        df_tampilan["S_max (Target Maksimum)"] = df_tampilan["S_max (kg)"].apply(lambda x: format_indonesia_kg(x) if pd.notna(x) else "-")
        df_tampilan["Safety Stock (SS)"] = df_tampilan["SS (kg)"].apply(format_indonesia_kg)
        df_tampilan["Total Cost (OT)"] = df_tampilan["Total Cost (OT)"].apply(format_rupiah)

        kolom_tampil = ["Bahan Baku", "Kategori", "Stok Saat Ini", "Status Gudang", "s_opt (Batas Aman/r)", "Q_opt (Pemesanan)", "S_max (Target Maksimum)", "Safety Stock (SS)", "Total Cost (OT)"]
        st.dataframe(df_tampilan[kolom_tampil], use_container_width=True)

        # Grafik Komparasi
        fig = px.bar(
            df_hasil,
            x="Bahan Baku",
            y=["Stok Saat Ini (kg)", "s_opt (kg)"],
            barmode="group",
            title="Komparasi Visual Stok Riil Terhadap Batas Reorder Point (s)",
            labels={"value": "Volume (kg)", "variable": "Parameter Gudang"},
            color_discrete_sequence=["#3399FF", "#FF3333"],
        )
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.markdown("### 📜 Log Laporan Mutasi Barang Gudang")
        df_log_tampil = muat_riwayat()
        if not df_log_tampil.empty:
            df_filtered = df_log_tampil.sort_values(by="Waktu", ascending=False)
            df_filtered["Jumlah (kg)"] = df_filtered["Jumlah (kg)"].apply(format_indonesia_kg)
            st.dataframe(df_filtered[["Waktu", "Bahan Baku", "Aktivitas", "Jumlah (kg)"]], use_container_width=True)

    with tab3:
        if st.button("Hapus & Reset Master Data Gudang 🗑️"):
            st.session_state["data_gudang"] = None
            st.session_state["stok_realtime"] = {}
            if os.path.exists(FILE_RIWAYAT):
                os.remove(FILE_RIWAYAT)
            st.rerun()

else:
    # LANDING PAGE SAAT BELUM ADA FILE YANG DIUPLOAD
    st.info(
        "💡 Silakan upload file Excel data Inventory Anda di panel sebelah kiri"
        " untuk Mulai Perhitungan & mengaktifkan sistem monitoring."
    )

    st.markdown("---")
    st.markdown("### 📋 Rincian Data Parameter yang Dibutuhkan System")
    st.markdown(
        "Agar sistem dapat menghitung Reorder Point ($s$), Ukuran Pemesanan ($Q$), dan"
        " Maximum Stock ($S$), file Excel Anda wajib memiliki kolom berikut:"
    )

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("""
            **1. Demand ($D$) - Total Permintaan Tahunan**
            * **Apa itu?** Total jumlah pemakaian barang dalam 1 tahun (dalam kg/unit).
            * **Cara dapatnya:** Jumlahkan total pemakaian dari bulan Januari sampai Desember.

            **2. Standar Deviasi / Sigma ($\sigma$) - Fluktuasi Pemakaian**
            * **Apa itu?** Angka yang menunjukkan seberapa fluktuatif pemakaian barang tiap bulannya.
            * **Cara dapatnya:** Gunakan rumus Excel `=STDEV.S(data_12_bulan)*SQRT(12)`.

            **3. Lead Time ($L$) - Waktu Tunggu Pengiriman**
            * **Apa itu?** Durasi waktu sejak barang dipesan sampai tiba di gudang.
            * **Cara dapatnya:** Tanyakan ke Supplier jika pesan sekarang berapa hari sampainya.
            """)

    with col_b:
        st.markdown("""
            **4. Biaya Pesan ($A$) - Ordering Cost**
            * **Apa itu?** Biaya yang keluar setiap kali Anda melakukan 1 kali pemesanan.
            * **Contoh:** Biaya administrasi, ongkos kirim, telepon, pemeriksaan barang.

            **5. Biaya Simpan ($h$) - Holding Cost**
            * **Apa itu?** Biaya untuk menyimpan 1 unit barang di gudang selama 1 TAHUN.
            * **Contoh:** Biaya sewa tempat, listrik gudang, perawatan, atau fasilitas lain.

            **6. Biaya Kekurangan ($C_u$) - Shortage Cost**
            * **Apa itu?** Kerugian atau estimasi biaya denda/kehilangan profit jika persediaan gudang HABIS saat produksi berjalan.
            """)

    st.markdown("---")
    st.success(
        "💡 **Sudah Dijelaskan Tapi Masih Bingung, hehehe 😂. Tenang, Udah Disediakan"
        " Template Kosongan Kok 🥳** Klik tombol **'📥 Unduh Template Excel Di slide samping aja"
        " Otomatis'** di menu sebelah kiri untuk mengunduh template Excel siap"
        " isi!"
    )
