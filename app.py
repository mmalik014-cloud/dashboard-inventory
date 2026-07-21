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
        <h2 style="font-size: 36px; font-weight: 700; margin-bottom: 4px;">
            📦 INTEGRATED INVENTORY CONTROL & REAL-TIME MONITORING SYSTEM
        </h2>
        <p style="font-size: 26px; color: #A0AAB8; font-weight: 500; margin-top: 0;">
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
        df["Nilai_Penyerapan"] = df["D_num"] * df["pi_num"]
        
        df_abc = df.sort_values(by="Nilai_Penyerapan", ascending=False).copy()
        total_penyerapan = df_abc["Nilai_Penyerapan"].sum()
        
        if total_penyerapan > 0:
            df_abc["Persen"] = (df_abc["Nilai_Penyerapan"] / total_penyerapan) * 100
            df_abc["Kumulatif"] = df_abc["Persen"].cumsum()
        else:
            df_abc["Kumulatif"] = 0.0

        def tentukan_kategori_otomatis(kumulatif_persen):
            if kumulatif_persen <= 80.0:
                return "A"
            elif kumulatif_persen <= 95.0:
                return "B"
            else:
                return "C"

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
            pi = row["pi_num"]
            kategori = row["Kategori"]

            metode = "Continuous Review (s,S)" if kategori == "A" else "Continuous Review (s,Q)"

            sigma_L = sigma * math.sqrt(L) if sigma > 0 and L > 0 else 1.0
            demand_L = D * L
            
            q_prev = math.sqrt((2 * A * D) / h) if h > 0 else 1.0
            r_prev = demand_L

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
            S_final = round(q_prev + r_prev, 2) if kategori == "A" else None

            OP = (A * D) / q_prev if q_prev > 0 else 0
            OS = h * ((q_prev / 2) + SS_final)
            
            alpha_f = min(max((h * q_prev) / (Cu * D if Cu * D > 0 else 1), 1e-6), 1 - 1e-6)
            Z_f = norm.ppf(1 - alpha_f)
            N_f = sigma_L * (norm.pdf(Z_f) - Z_f * (1 - norm.cdf(Z_f)))
            OK = Cu * (D / q_prev if q_prev > 0 else 0) * N_f

            OT_final = OP + OS + OK

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
                "pi_num": pi,
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

    # --------------------------------------------------------------------------
    # PAPAN PERHATIAN KRITIS & MENDEKATI ROP
    # --------------------------------------------------------------------------
    st.markdown("### 📢 PAPAN PERHATIAN STOK PERSEDIAAN")

    df_kritis_po = []

    for idx, row in df_hasil.iterrows():
        stok_aktual = row["Stok Saat Ini (kg)"]
        rop = row["s_opt (kg)"]
        selisih = stok_aktual - rop
        pct_rop = (stok_aktual / rop * 100) if rop > 0 else 0
        qty_rekomendasi = (row["S_max (kg)"] - stok_aktual) if row["Kategori"] == "A" else row["Q_opt (kg)"]
        label_metode = "Rekomendasi Berdasarkan (s,S)" if row["Kategori"] == "A" else "Rekomendasi Berdasarkan (s,Q)"

        # KONDISI 1: KRITIS (Di bawah ROP)
        if selisih <= 0:
            defisit = abs(selisih)
            pct_defisit = ((rop - stok_aktual) / rop * 100) if rop > 0 else 0

            with st.status(f"🚨 **{row['Bahan Baku']}** — Defisit {defisit:,.2f} kg ({pct_defisit:.1f}% di bawah ROP)", state="error", expanded=True):
                c_m1, c_m2, c_m3 = st.columns(3)
                with c_m1:
                    st.caption("STATUS STOK AKTUAL")
                    st.metric(
                        label="Stok Saat Ini", 
                        value=f"{stok_aktual:,.2f} kg", 
                        delta=f"-{defisit:,.2f} kg turun di bawah ROP", 
                        delta_color="inverse"
                    )
                with c_m2:
                    st.caption("POSISI PERSENTASE")
                    st.metric(
                        label="Rasio terhadap ROP", 
                        value=f"{pct_rop:.1f}%", 
                        delta=f"-{pct_defisit:.1f}% Kekurangan", 
                        delta_color="inverse"
                    )
                with c_m3:
                    st.caption("REKOMENDASI PEMESANAN")
                    st.metric(label="Jumlah Harus Dipesan", value=f"{qty_rekomendasi:,.2f} kg")
                    st.caption(f"👉 **{label_metode}**")

            df_kritis_po.append({
                "Nama Bahan Baku": row["Bahan Baku"],
                "Kategori": row["Kategori"],
                "Stok Aktual (kg)": stok_aktual,
                "Batas ROP (kg)": rop,
                "Defisit (kg)": defisit,
                "Rekomendasi Pesan (kg)": qty_rekomendasi,
                "Estimasi Biaya (Rp)": qty_rekomendasi * row["pi_num"]
            })

        # KONDISI 2: WARNING (Mendekati ROP: 100% - 120%)
        elif 100 <= pct_rop <= 120:
            with st.status(f"⚠️ **{row['Bahan Baku']}** — Mendekati Batas Aman (Sisa +{selisih:,.2f} kg lagi)", state="warning", expanded=False):
                c_m1, c_m2, c_m3 = st.columns(3)
                with c_m1:
                    st.caption("STATUS STOK AKTUAL")
                    st.write(f"**Stok:** {stok_aktual:,.2f} kg")
                    st.write(f"**Batas ROP:** {rop:,.2f} kg")
                with c_m2:
                    st.caption("POSISI PERSENTASE")
                    st.write(f"**Kapasitas:** {pct_rop:.1f}% dari ROP")
                    st.write(f"**Margin Aman:** +{selisih:,.2f} kg")
                with c_m3:
                    st.caption("STATUS AKSI")
                    st.info(f"👀 Siapkan Pemesanan ({label_metode})")

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

    st.markdown("""
<style>
/* Mengubah ukuran teks Tab Utama (Gambar 1) */
button[data-baseweb="tab"] p {
    font-size: 26px !important; /* Ubah angka 18px untuk memperbesar/memperkecil */
    font-weight: bold !important;
}
</style>
""", unsafe_allow_html=True)

    # TAB MONITORING, SURAT PO, LAPORAN, & RESET
    tab1, tab_po, tab2, tab3 = st.tabs([
        "📊 Live Monitoring & Chart", 
        "📜 Draft Surat Purchase Order (PO)", 
        "📜 Laporan Riwayat Keluar-Masuk", 
        "⚙️ Reset Data"
    ])

    with tab1:
        st.markdown("#### Hasil Perhitungan & Status Gudang Real-Time 📝")
        df_tampilan = df_hasil.copy()
        
        def status_gudang_text(r):
            return "🚨 HARUS REORDER!" if r["Stok Saat Ini (kg)"] <= r["s_opt (kg)"] else "✅ Stok Aman"
            
        df_tampilan["Status Gudang"] = df_tampilan.apply(status_gudang_text, axis=1)
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

    # --------------------------------------------------------------------------
    # GENERATOR SURAT PURCHASE ORDER (PO) OTOMATIS (2 LEMBAR: INTERNAL & EKSTERNAL)
    # --------------------------------------------------------------------------
    with tab_po:
        st.markdown("<h2 style='font-size: 28px; font-weight: bold;'>📜 Generator Dokumen Purchase Order (PO)</h2>", unsafe_allow_html=True)

        if df_kritis_po:
            df_po = pd.DataFrame(df_kritis_po)
            today_str = datetime.now().strftime("%d %B %Y")
            no_po_auto = f"PO/INV/{datetime.now().strftime('%Y%m%d')}/001"

            with st.expander("✏️ Edit & Sesuaikan Informasi Header Surat PO", expanded=False):
                c_p1, c_p2 = st.columns(2)
                with c_p1:
                    vendor_nama = st.text_input("Nama Supplier / Vendor", value="PT. Supplier Logistik Utama")
                    nomor_po = st.text_input("Nomor Surat PO", value=no_po_auto)
                with c_p2:
                    petugas_nama = st.text_input("Nama Petugas Gudang / Manager", value="Admin Logistics")
                    catatan_po = st.text_area("Catatan Pengiriman", value="Mohon dikirimkan maksimal 3 hari kerja setelah surat PO ini diterbitkan.")

            # PILIHAN LEMBAR DOKUMEN: INTERNAL VS EKSTERNAL
            st.markdown("""
<style>
div[data-testid='stHorizontalBlock'] button[data-baseweb='tab'] div p {
    font-size: 26px !important; 
    font-weight: bold !important;
}
</style>
""", unsafe_allow_html=True)
            sub_po_tab1, sub_po_tab2 = st.tabs(["📄 1. Surat PO Resmi untuk Supplier (Eksternal)", "🔒 2. Draft Catatan Internal Perusahaan (Lengkap)"])

            # LEMBAR 1: EKSTERNAL SUPPLIER (LANGSUNG SUSUN KE BAWAH)
            with sub_po_tab1:
                st.info("🔒 Setiap bahan baku diterbitkan dalam Surat PO terpisah secara vertikal.")
                
                for i, bahan in enumerate(df_po["Nama Bahan Baku"].unique()):
                    row = df_po[df_po["Nama Bahan Baku"] == bahan].iloc[0]
                    no_po = f"PO/{bahan[:3].upper()}/{datetime.now().strftime('%Y%m%d')}/{i+1:02d}"
                    df_item = pd.DataFrame([{"No": 1, "Nama Bahan Baku": bahan, "Jumlah Pesanan": format_indonesia_kg(row["Rekomendasi Pesan (kg)"])}])
                    
                    with st.container(border=True):
                        st.markdown(f"<h3 style='text-align:center;'>SURAT PEMESANAN BARANG (PURCHASE ORDER)</h3><p style='text-align:center;'>Item: <b>{bahan}</b> (Kategori {row['Kategori']})</p><hr>", unsafe_allow_html=True)
                        
                        c1, c2 = st.columns(2)
                        c1.write(f"**Kepada Yth:** PT. Supplier {bahan} Utama\n\n**Tanggal:** {today_str}")
                        c2.write(f"**No. PO:** `{no_po}`\n\n**Diterbitkan Oleh:** {petugas_nama}")
                        
                        st.dataframe(df_item, use_container_width=True, hide_index=True)
                        st.caption(f"📌 *Catatan:* {catatan_po}")
                        
                        st.write("<br>", unsafe_allow_html=True)
                        t1, t2 = st.columns(2)
                        t1.write(f"Hormat Kami,\n\n\n**( {petugas_nama} )**")
                        t2.write("Disetujui Oleh,\n\n\n**( .................... )**")

                    st.download_button(
                        f"🖨️ Download Surat PO ({bahan})",
                        df_item.to_csv(index=False).encode('utf-8'),
                        f"Surat_PO_{bahan}.csv",
                        "text/csv",
                        type="primary",
                        key=f"dl_{i}"
                    )
                    st.markdown("<hr style='border:1px dashed #555;'><br>", unsafe_allow_html=True)

            # LEMBAR 2: INTERNAL PERUSAHAAN
            with sub_po_tab2:
                st.warning("⚠️ **Dokumen Rahasia Internal:** Memuat analisis defisit persediaan, ROP, Kategori ABC, dan Estimasi Anggaran Biaya untuk pengawasan internal perusahaan.")
                
                with st.container(border=True):
                    st.markdown("""
                    <div style="text-align: center; border-bottom: 2px solid #555; padding-bottom: 10px; margin-bottom: 20px;">
                        <h2 style="margin:0; color: #FF4B4B;">DRAFT MEMO INTERNAL - ANALISIS & ANGGARAN PO</h2>
                        <p style="margin:0; color: #888;">Dokumen Lampiran Internal / Approval Procurement</p>
                    </div>
                    """, unsafe_allow_html=True)

                    col_h1, col_h2 = st.columns(2)
                    with col_h1:
                        st.write(f"**Target Supplier:** Multiple Suppliers (Per Kategori)")
                        st.write(f"**Tanggal Analisis:** {today_str}")
                    with col_h2:
                        st.write(f"**Ref No. PO:** `{nomor_po}`")
                        st.write(f"**Analis Logistics:** {petugas_nama}")

                    st.markdown("#### **Rincian Parameter Persediaan & Estimasi Anggaran:**")
                    
                    df_po_show = df_po.copy()
                    df_po_show["Stok Aktual (kg)"] = df_po_show["Stok Aktual (kg)"].apply(format_indonesia_kg)
                    df_po_show["Batas ROP (kg)"] = df_po_show["Batas ROP (kg)"].apply(format_indonesia_kg)
                    df_po_show["Rekomendasi Pesan (kg)"] = df_po_show["Rekomendasi Pesan (kg)"].apply(format_indonesia_kg)
                    df_po_show["Estimasi Biaya"] = df_po_show["Estimasi Biaya (Rp)"].apply(format_rupiah)

                    st.dataframe(
                        df_po_show[["Nama Bahan Baku", "Kategori", "Stok Aktual (kg)", "Batas ROP (kg)", "Rekomendasi Pesan (kg)", "Estimasi Biaya"]],
                        use_container_width=True,
                        hide_index=True
                    )

                    total_est_biaya = df_po["Estimasi Biaya (Rp)"].sum()
                    st.markdown(f"### **Total Pengajuan Anggaran: {format_rupiah(total_est_biaya)}**")

                    st.markdown("<br>", unsafe_allow_html=True)
                    col_ttd1, col_ttd2 = st.columns(2)
                    with col_ttd1:
                        st.write("Disiapkan Oleh (Gudang/Logistik),")
                        st.markdown("<br><br>", unsafe_allow_html=True)
                        st.write(f"**( {petugas_nama} )**")
                    with col_ttd2:
                        st.write("Persetujuan Anggaran (Finance/Manager),")
                        st.markdown("<br><br>", unsafe_allow_html=True)
                        st.write("**( .................................... )**")

                csv_po = df_po.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Draft Catatan Internal (Lengkap/CSV)",
                    data=csv_po,
                    file_name=f"Draft_Internal_PO_{nomor_po.replace('/', '_')}.csv",
                    mime="text/csv",
                    key="btn_dl_int_total"
                )
        else:
            st.info("✅ Tidak ada bahan baku yang berada di bawah Reorder Point (ROP). Belum ada draft PO yang perlu diterbitkan.")

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
        "💡 Silakan upload file Excel data Inventory Anda di panel sebelah kiri "
        "untuk Mulai Perhitungan & mengaktifkan sistem monitoring."
    )

    st.markdown("---")
    st.markdown("### 📋 Rincian Data Parameter yang Dibutuhkan System")
    st.markdown(
        "Agar sistem dapat menghitung Reorder Point ($s$), Ukuran Pemesanan ($Q$), dan "
        "Maximum Stock ($S$), file Excel Anda wajib memiliki kolom berikut:"
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
        * **Apa itu?** Durasi waktu sejak barang dipesan sampai tiba di gudang (dalam satuan TAHUN/HARI).
        * **Cara dapatnya:** Jika supplier butuh waktu 30 hari, isi angka 30 (sistem akan mengonversi otomatis).
        """)

    with col_b:
        st.markdown("""
        **4. Biaya Pesan ($A$) - Ordering Cost**
        * **Apa itu?** Biaya yang keluar setiap kali Anda melakukan 1 kali pemesanan.
        * **Contoh:** Biaya administrasi, ongkos kirim, telepon, pemeriksaan barang.

        **5. Biaya Simpan ($h$) - Holding Cost**
        * **Apa itu?** Biaya untuk menyimpan 1 unit barang di gudang selama 1 TAHUN.
        * **Contoh:** Biaya sewa tempat, listrik gudang, perawatan, atau kerusakan per unit.

        **6. Biaya Kekurangan ($C_u$) - Shortage Cost**
        * **Apa itu?** Kerugian atau estimasi biaya denda/kehilangan profit jika persediaan gudang HABIS saat produksi berjalan.
        """)

    st.markdown("---")
    st.success(
        "💡 **Sudah Dijelaskan Tapi Masih Bingung wkwkwk. Tenang, Udah Disediakan "
        "Template Kosongan Kok 🥳** Klik tombol **'📥 Unduh Template Excel "
        "Otomatis'** di menu sebelah kiri untuk mengunduh template Excel siap isi!"
    )
