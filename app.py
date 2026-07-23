import io
import math
import os
from datetime import datetime
import openpyxl
import pandas as pd
import plotly.express as px
from scipy.stats import norm
import streamlit as st

# Impor library ReportLab untuk pembuatan PDF
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

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
# 2. FUNGSI HELPER & FORMATTING ANGKA
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

def format_indonesia_satuan(nilai, satuan="pcs"):
    """Format angka desimal Indonesia dengan satuan dinamis dari Excel."""
    if nilai == "-" or pd.isna(nilai) or nilai is None:
        return "-"
    try:
        val = float(nilai)
        formatted_val = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{formatted_val} {satuan}"
    except:
        return str(nilai)

def format_rupiah(nilai):
    try:
        val = float(nilai)
        return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(nilai)

def muat_riwayat():
    kolom_standar = ["Waktu", "Tanggal", "Bulan", "Bahan Baku", "Aktivitas", "Jumlah", "Satuan"]
    if os.path.exists(FILE_RIWAYAT):
        try:
            df_log = pd.read_excel(FILE_RIWAYAT)
            for col in kolom_standar:
                if col not in df_log.columns:
                    df_log[col] = 0.0 if col == "Jumlah" else ""
            df_log["Waktu"] = pd.to_datetime(df_log["Waktu"])
            return df_log
        except:
            pass
    return pd.DataFrame(columns=kolom_standar)

def catat_transaksi(bahan, aktivitas, jumlah, satuan="pcs"):
    df_log = muat_riwayat()
    sekarang = datetime.now()
    data_baru = pd.DataFrame([{
        "Waktu": sekarang,
        "Tanggal": sekarang.strftime("%Y-%m-%d"),
        "Bulan": sekarang.strftime("%Y-%m"),
        "Bahan Baku": bahan,
        "Aktivitas": aktivitas,
        "Jumlah": jumlah,
        "Satuan": satuan
    }])
    df_log = pd.concat([df_log, data_baru], ignore_index=True)
    df_log.to_excel(FILE_RIWAYAT, index=False)

# --- FUNGSI GENERATOR DOKUMEN PDF RESMI ---
def buat_surat_po_pdf(bahan, kategori, satuan, qty_pesan, vendor_nama, no_po, petugas_nama, tanggal_str, catatan):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle('TitleStyle', parent=styles['Heading1'], alignment=1, fontSize=16, spaceAfter=4)
    style_sub = ParagraphStyle('SubStyle', parent=styles['Normal'], alignment=1, fontSize=10, textColor=colors.HexColor('#555555'))
    style_normal = ParagraphStyle('NormStyle', parent=styles['Normal'], fontSize=10, leading=14)
    style_bold = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=10, leading=14, fontName="Helvetica-Bold")

    # Header Surat
    story.append(Paragraph("<b>SURAT PEMESANAN BARANG (PURCHASE ORDER)</b>", style_title))
    story.append(Paragraph(f"Item: <b>{bahan}</b> (Kategori {kategori})", style_sub))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#333333'), spaceAfter=15))

    # Informasi PO
    info_data = [
        [Paragraph(f"<b>Kepada Yth:</b><br/>{vendor_nama}<br/><b>Tanggal:</b> {tanggal_str}", style_normal),
         Paragraph(f"<b>No. PO:</b> {no_po}<br/><b>Diterbitkan Oleh:</b> {petugas_nama}", style_normal)]
    ]
    t_info = Table(info_data, colWidths=[260, 260])
    t_info.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(t_info)
    story.append(Spacer(1, 15))

    # Tabel Barang
    table_data = [
        ["No", "Nama Bahan Baku", "Jumlah Pesanan"],
        ["1", str(bahan), format_indonesia_satuan(qty_pesan, satuan)]
    ]
    t_items = Table(table_data, colWidths=[40, 320, 160])
    t_items.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F2F2F2')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
        ('PADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_items)
    story.append(Spacer(1, 15))

    story.append(Paragraph(f"<b>Catatan Pengiriman:</b> {catatan}", style_normal))
    story.append(Spacer(1, 40))

    # Tanda Tangan
    ttd_data = [
        [Paragraph("<b>Hormat Kami,</b>", style_normal), Paragraph("<b>Disetujui Oleh,</b>", style_normal)],
        ["", ""],
        ["", ""],
        [Paragraph(f"<b>( {petugas_nama} )</b>", style_bold), Paragraph("<b>( .................................... )</b>", style_bold)]
    ]
    t_ttd = Table(ttd_data, colWidths=[260, 260])
    t_ttd.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_ttd)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def buat_memo_internal_pdf(df_po, nomor_po, petugas_nama, today_str):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle('TitleStyle', parent=styles['Heading1'], alignment=1, fontSize=15, textColor=colors.HexColor('#D32F2F'), spaceAfter=4)
    style_sub = ParagraphStyle('SubStyle', parent=styles['Normal'], alignment=1, fontSize=9, textColor=colors.HexColor('#666666'))
    style_normal = ParagraphStyle('NormStyle', parent=styles['Normal'], fontSize=9, leading=12)
    style_header = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=9, leading=11, fontName="Helvetica-Bold")

    story.append(Paragraph("<b>DRAFT MEMO INTERNAL - ANALISIS & ANGGARAN PO</b>", style_title))
    story.append(Paragraph("Dokumen Lampiran Internal / Approval Procurement", style_sub))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#D32F2F'), spaceAfter=12))

    info_data = [
        [Paragraph(f"<b>Ref No. PO:</b> {nomor_po}<br/><b>Tanggal Analisis:</b> {today_str}", style_normal),
         Paragraph(f"<b>Analis Logistics:</b> {petugas_nama}<br/><b>Target:</b> Multiple Suppliers", style_normal)]
    ]
    t_info = Table(info_data, colWidths=[260, 260])
    story.append(t_info)
    story.append(Spacer(1, 12))

    # Tabel Analisis PDF
    headers = ["Nama Bahan", "Kat", "Satuan", "Stok Aktual", "Batas ROP", "Rekom. Pesan", "Estimasi Biaya"]
    table_data = [[Paragraph(f"<b>{h}</b>", style_header) for h in headers]]

    for _, row in df_po.iterrows():
        table_data.append([
            Paragraph(str(row["Nama Bahan Baku"]), style_normal),
            Paragraph(str(row["Kategori"]), style_normal),
            Paragraph(str(row["Satuan"]), style_normal),
            Paragraph(format_indonesia_satuan(row["Stok Aktual"], row["Satuan"]), style_normal),
            Paragraph(format_indonesia_satuan(row["Batas ROP"], row["Satuan"]), style_normal),
            Paragraph(format_indonesia_satuan(row["Rekomendasi Pesan"], row["Satuan"]), style_normal),
            Paragraph(format_rupiah(row["Estimasi Biaya (Rp)"]), style_normal)
        ])

    t_po = Table(table_data, colWidths=[100, 30, 45, 80, 80, 85, 100])
    t_po.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#EEEEEE')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
        ('PADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_po)
    story.append(Spacer(1, 10))

    total_est = df_po["Estimasi Biaya (Rp)"].sum()
    story.append(Paragraph(f"<b>Total Pengajuan Anggaran: {format_rupiah(total_est)}</b>", ParagraphStyle('Total', parent=styles['Heading2'], fontSize=12)))
    story.append(Spacer(1, 30))

    ttd_data = [
        [Paragraph("<b>Disiapkan Oleh (Gudang/Logistik),</b>", style_normal), Paragraph("<b>Persetujuan Anggaran (Finance/Manager),</b>", style_normal)],
        ["", ""],
        ["", ""],
        [Paragraph(f"<b>( {petugas_nama} )</b>", style_normal), Paragraph("<b>( .................................... )</b>", style_normal)]
    ]
    t_ttd = Table(ttd_data, colWidths=[260, 260])
    story.append(t_ttd)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def buat_template_excel():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "data verifikasi"

    headers = ["Jenis Barang", "D", "Sigma", "L", "A", "h", "Cu", "pi", "Satuan"]
    ws.append(headers)

    ws.append([
        "Beras",
        "=SUM(M14:X14)",
        "=STDEV.S(M14:X14)*SQRT(12)",
        4.0,
        500.0,
        40.0,
        3060.0,
        3000.0,
        "kg"
    ])

    ws.append([])

    keterangan = [
        ["", "", "", "", "", "", "", "", "", "", "Keterangan", ""],
        ["", "", "", "", "", "", "", "", "", "", "D", "Demand/Permintaan"],
        ["", "", "", "", "", "", "", "", "", "", "Sigma", "Deviasi"],
        ["", "", "", "", "", "", "", "", "", "", "L", "Akar Leadtime"],
        ["", "", "", "", "", "", "", "", "", "", "A", "Biaya Per pesan"],
        ["", "", "", "", "", "", "", "", "", "", "h", "Biaya Simpan"],
        ["", "", "", "", "", "", "", "", "", "", "Cu", "Biaya Backorder/Kekurangan"],
        ["", "", "", "", "", "", "", "", "", "", "Pi", "Harga Bahan Baku"],
        ["", "", "", "", "", "", "", "", "", "", "Satuan", "Satuan Ukur Barang (pcs/kg/box/dll)"],
    ]
    for row in keterangan:
        ws.append(row)

    ws.append([])
    ws.append(["", "", "", "", "", "", "", "", "", "", "", "Kolom Perhitungan sigma "])
    ws.append(["", "", "", "", "", "", "", "", "", "", "", "Jenis Barang", "Bulan Periode Pengamatan"])
    
    header_bulan = [
        "", "", "", "", "", "", "", "", "", "", "", "",
        "Januari ", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember", "Total "
    ]
    ws.append(header_bulan)

    data_bulanan_beras = [
        "", "", "", "", "", "", "", "", "", "", "", "Beras ",
        1200, 1150, 1300, 1250, 1400, 1100, 1200, 1250, 1180, 1220, 1300, 1450, "=SUM(M14:X14)"
    ]
    ws.append(data_bulanan_beras)

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
        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file, delimiter=";")
        else:
            df_raw = pd.read_excel(uploaded_file)

        df_raw.columns = df_raw.columns.astype(str).str.strip()

        col_nama = "Jenis Barang" if "Jenis Barang" in df_raw.columns else "Bahan Baku"
        df = df_raw.dropna(subset=[col_nama]).copy()

        df["Bahan_Nama"] = df[col_nama].astype(str).str.strip()
        df["D_num"] = df["D"].apply(bersihkan_angka)
        df["Sigma_num"] = df["Sigma"].apply(bersihkan_angka)

        df["L_hari"] = df["L"].apply(bersihkan_angka)
        df["L_num"] = df["L_hari"].apply(lambda val: val / 365.0 if val > 1.0 else val)

        df["A_num"] = df["A"].apply(bersihkan_angka)
        df["h_num"] = df["h"].apply(bersihkan_angka)
        
        df["Cu_num"] = df["Cu"].apply(bersihkan_angka) if "Cu" in df_raw.columns else 0.0

        if "pi" in df_raw.columns:
            df["pi_num"] = df["pi"].apply(bersihkan_angka)
        else:
            df["pi_num"] = df["Cu_num"]

        satuan_col = [col for col in df_raw.columns if col.strip().lower() == 'satuan']
        if satuan_col:
            df["Satuan"] = df[satuan_col[0]].fillna("pcs").astype(str).str.strip()
            df["Satuan"] = df["Satuan"].replace("", "pcs")
        else:
            df["Satuan"] = "pcs"

        # Otomasi ABC Analysis
        df["Nilai_Penyerapan"] = df["D_num"] * df["pi_num"]
        df_abc = df.sort_values(by="Nilai_Penyerapan", ascending=False).copy()
        total_penyerapan = df_abc["Nilai_Penyerapan"].sum()

        if total_penyerapan > 0:
            df_abc["Persen"] = (df_abc["Nilai_Penyerapan"] / total_penyerapan) * 100
            df_abc["Kumulatif"] = df_abc["Persen"].cumsum()
        else:
            df_abc["Persen"] = 0.0
            df_abc["Kumulatif"] = 0.0

        def tentukan_kategori_standar(row):
            if row["Kumulatif"] <= 80.0:
                return "A"
            elif row["Kumulatif"] <= 95.0:
                return "B"
            else:
                return "C"

        df_abc["Kategori_Fix"] = df_abc.apply(tentukan_kategori_standar, axis=1)
        kat_mapping = dict(zip(df_abc["Bahan_Nama"], df_abc["Kategori_Fix"]))
        df["Kategori"] = df["Bahan_Nama"].map(kat_mapping)

        # Iterasi Hadley-Whitin
        hasil_perhitungan = []
        df_log_ada = muat_riwayat()

        for idx, row in df.iterrows():
            bahan = row["Bahan_Nama"]
            D = row["D_num"]
            sigma = row["Sigma_num"]
            L = row["L_num"]
            A = row["A_num"]
            h = row["h_num"]
            pi = row["pi_num"]
            satuan_item = row["Satuan"]
            kategori = row["Kategori"]

            Cu = row["Cu_num"]
            if Cu <= 0:
                Cu = 2.0 * pi if pi > 0 else (h * 5.0 if h > 0 else 1000.0)

            metode = "Continuous Review (s,S)" if kategori == "A" else "Continuous Review (s,Q)"

            sigma_L = sigma * math.sqrt(L) if sigma > 0 and L > 0 else 1.0
            demand_L = D * L
            
            q_prev = math.sqrt((2 * A * D) / h) if h > 0 else 1.0
            r_prev = demand_L

            for i in range(max_iter):
                denom = Cu * D
                alpha = (h * q_prev) / denom if denom > 0 else 0.05
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
            
            denom_f = Cu * D if Cu * D > 0 else 1.0
            alpha_f = min(max((h * q_prev) / denom_f, 1e-6), 1 - 1e-6)
            Z_f = norm.ppf(1 - alpha_f)
            N_f = sigma_L * (norm.pdf(Z_f) - Z_f * (1 - norm.cdf(Z_f)))
            OK = Cu * (D / q_prev if q_prev > 0 else 0) * N_f

            OT_final = OP + OS + OK
            stok_awal = round(q_prev + r_prev, 2) if kategori == "A" else round(q_prev + SS_final, 2)

            if not df_log_ada.empty and "Jumlah" in df_log_ada.columns:
                log_bahan = df_log_ada[df_log_ada["Bahan Baku"] == bahan]
                if not log_bahan.empty:
                    masuk = log_bahan[log_bahan["Aktivitas"] == "Barang Datang (Stok Masuk)"]["Jumlah"].sum()
                    keluar = log_bahan[log_bahan["Aktivitas"] == "Diambil Produksi (Stok Keluar)"]["Jumlah"].sum()
                    stok_awal = stok_awal + masuk - keluar

            st.session_state["stok_realtime"][bahan] = float(stok_awal)

            hasil_perhitungan.append({
                "Bahan Baku": bahan,
                "Kategori": kategori,
                "Metode": metode,
                "Satuan": satuan_item,
                "Q_opt": round(q_prev, 2),
                "s_opt": round(r_prev, 2),
                "S_max": S_final,
                "SS": round(SS_final, 2),
                "Total Cost (OT)": round(OT_final, 2),
                "pi_num": pi,
            })

        st.session_state["data_gudang"] = pd.DataFrame(hasil_perhitungan)
    except Exception as e:
        st.error(f"Gagal memproses data template Excel: {e}")

# ============================================================
# 5. TAMPILAN DASHBOARD & PAPAN PERHATIAN STOK
# ============================================================
if st.session_state["data_gudang"] is not None:
    df_hasil = st.session_state["data_gudang"].copy()
    df_hasil["Stok Saat Ini"] = df_hasil["Bahan Baku"].map(st.session_state["stok_realtime"])

    st.markdown("### 📢 PAPAN PERHATIAN STOK PERSEDIAAN")

    df_kritis_po = []
    item_aman = []
    item_warning = []
    item_kritis = []

    for idx, row in df_hasil.iterrows():
        stok_aktual = row["Stok Saat Ini"]
        rop = row["s_opt"]
        satuan = row["Satuan"]
        selisih = stok_aktual - rop
        pct_rop = (stok_aktual / rop * 100) if rop > 0 else 0
        qty_rekomendasi = (row["S_max"] - stok_aktual) if row["Kategori"] == "A" else row["Q_opt"]
        label_metode = "(s,S)" if row["Kategori"] == "A" else "(s,Q)"

        row_data = {
            "row": row,
            "stok_aktual": stok_aktual,
            "rop": rop,
            "satuan": satuan,
            "selisih": selisih,
            "pct_rop": pct_rop,
            "qty_rekomendasi": qty_rekomendasi,
            "label_metode": label_metode
        }

        if stok_aktual <= rop:
            item_kritis.append(row_data)
            df_kritis_po.append({
                "Nama Bahan Baku": row["Bahan Baku"],
                "Kategori": row["Kategori"],
                "Satuan": satuan,
                "Stok Aktual": stok_aktual,
                "Batas ROP": rop,
                "Defisit": abs(selisih),
                "Rekomendasi Pesan": qty_rekomendasi,
                "Estimasi Biaya (Rp)": qty_rekomendasi * row["pi_num"]
            })
        elif 100 < pct_rop <= 120:
            item_warning.append(row_data)
        else:
            item_aman.append(row_data)

    if item_kritis:
        for data in item_kritis:
            row = data["row"]
            defisit = abs(data["selisih"])
            pct_defisit = ((data["rop"] - data["stok_aktual"]) / data["rop"] * 100) if data["rop"] > 0 else 0

            with st.status(f"🚨 **{row['Bahan Baku']}** — Status Kritis: Defisit {format_indonesia_satuan(defisit, data['satuan'])} ({pct_defisit:.1f}% di bawah ROP)", state="error", expanded=True):
                c_m1, c_m2, c_m3 = st.columns(3)
                with c_m1:
                    st.caption("STATUS STOK AKTUAL")
                    st.metric(label="Stok Saat Ini", value=format_indonesia_satuan(data["stok_aktual"], data["satuan"]), delta=f"-{format_indonesia_satuan(defisit, data['satuan'])} di bawah ROP", delta_color="inverse")
                with c_m2:
                    st.caption("POSISI PERSENTASE")
                    st.metric(label="Rasio terhadap ROP", value=f"{data['pct_rop']:.1f}%", delta=f"-{pct_defisit:.1f}% Defisit", delta_color="inverse")
                with c_m3:
                    st.caption("REKOMENDASI PEMESANAN")
                    st.metric(label="Jumlah Harus Dipesan", value=format_indonesia_satuan(data["qty_rekomendasi"], data["satuan"]))
                    st.caption(f"👉 **Rekomendasi Berdasarkan {data['label_metode']}**")

    if item_warning:
        for data in item_warning:
            row = data["row"]
            q_formatted = format_indonesia_satuan(data["qty_rekomendasi"], data["satuan"])
            with st.status(f"⚠️ **{row['Bahan Baku']}** — Mendekati ROP (Sisa Selisih Stok dari ROP +{format_indonesia_satuan(data['selisih'], data['satuan'])})", state="running", expanded=True):
                c_m1, c_m2, c_m3 = st.columns(3)
                with c_m1:
                    st.caption("STATUS STOK AKTUAL")
                    st.write(f"**Stok Saat Ini:** {format_indonesia_satuan(data['stok_aktual'], data['satuan'])}")
                    st.write(f"**Batas ROP ($s$):** {format_indonesia_satuan(data['rop'], data['satuan'])}")
                with c_m2:
                    st.caption("POSISI PERSENTASE")
                    st.write(f"**Level Stok:** {data['pct_rop']:.1f}% dari ROP")
                    st.write(f"**Selisih Stok dari ROP:** +{format_indonesia_satuan(data['selisih'], data['satuan'])}")
                with c_m3:
                    st.caption("STATUS AKSI")
                    st.info(f"👀 Siapkan Pemesanan Sejumlah Q = {q_formatted} (Rekomendasi Berdasarkan {data['label_metode']})")

    if item_aman:
        nama_bahan_aman = ", ".join([d["row"]["Bahan Baku"] for d in item_aman])
        with st.expander(f"✅ **{len(item_aman)} Barang dalam Kondisi Aman** (Stok > 120% di atas ROP): {nama_bahan_aman}", expanded=False):
            st.caption("Berikut adalah rincian persediaan yang saat ini berada pada batas aman di atas Reorder Point (ROP):")
            for data in item_aman:
                row = data["row"]
                with st.status(f"✅ **{row['Bahan Baku']}** — Status Aman (Selisih Stok dari ROP +{format_indonesia_satuan(data['selisih'], data['satuan'])})", state="complete", expanded=True):
                    c_a1, c_a2, c_a3 = st.columns(3)
                    with c_a1:
                        st.caption("STATUS STOK AKTUAL")
                        st.write(f"**Stok Saat Ini:** {format_indonesia_satuan(data['stok_aktual'], data['satuan'])}")
                        st.write(f"**Batas ROP ($s$):** {format_indonesia_satuan(data['rop'], data['satuan'])}")
                    with c_a2:
                        st.caption("POSISI PERSENTASE")
                        st.write(f"**Level Stok:** {data['pct_rop']:.1f}% dari ROP")
                        st.write(f"**Selisih Stok dari ROP:** +{format_indonesia_satuan(data['selisih'], data['satuan'])}")
                    with c_a3:
                        st.caption("STATUS AKSI")
                        st.success("✅ **Stok Optimal** (Tidak Perlu Pemesanan)")

    st.markdown("---")

    # PANEL TRANSAKSI MUTASI
    st.markdown("### 📥 📤 Panel Transaksi Gudang")
    c1, c2, c3 = st.columns(3)
    with c1:
        pilih_bahan = st.selectbox("Pilih Jenis Barang:", df_hasil["Bahan Baku"].tolist())
        satuan_terpilih = df_hasil[df_hasil["Bahan Baku"] == pilih_bahan]["Satuan"].values[0]
    with c2:
        jenis_transaksi = st.radio("Aktivitas:", ["Barang Datang (Stok Masuk)", "Diambil Produksi (Stok Keluar)"])
    with c3:
        jumlah_mutasi = st.number_input(f"Jumlah ({satuan_terpilih}):", min_value=0.0, step=10.0)
        submit_button = st.button("Simpan Transaksi 💾")

    if submit_button and jumlah_mutasi > 0:
        if jenis_transaksi == "Barang Datang (Stok Masuk)":
            catat_transaksi(pilih_bahan, jenis_transaksi, jumlah_mutasi, satuan_terpilih)
            st.session_state["stok_realtime"][pilih_bahan] += jumlah_mutasi
            st.toast(f"Berhasil mencatat stok masuk ({satuan_terpilih})!", icon="📥")
        else:
            catat_transaksi(pilih_bahan, jenis_transaksi, jumlah_mutasi, satuan_terpilih)
            st.session_state["stok_realtime"][pilih_bahan] -= jumlah_mutasi
            st.toast(f"Berhasil mencatat stok keluar ({satuan_terpilih})!", icon="📤")
        st.rerun()

    st.markdown("""
    <style>
    button[data-baseweb="tab"] p {
        font-size: 26px !important;
        font-weight: bold !important;
    }
    </style>
    """, unsafe_allow_html=True)

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
            return "🚨 HARUS REORDER!" if r["Stok Saat Ini"] <= r["s_opt"] else "✅ Stok Aman"
            
        df_tampilan["Status Gudang"] = df_tampilan.apply(status_gudang_text, axis=1)
        df_tampilan["Stok Saat Ini Tampil"] = df_tampilan.apply(lambda r: format_indonesia_satuan(r["Stok Saat Ini"], r["Satuan"]), axis=1)
        df_tampilan["s_opt (Batas Aman/r)"] = df_tampilan.apply(lambda r: format_indonesia_satuan(r["s_opt"], r["Satuan"]), axis=1)
        df_tampilan["Q_opt (Pemesanan)"] = df_tampilan.apply(lambda r: format_indonesia_satuan(r["Q_opt"], r["Satuan"]), axis=1)
        df_tampilan["S_max (Target Maksimum)"] = df_tampilan.apply(lambda r: format_indonesia_satuan(r["S_max"], r["Satuan"]) if pd.notna(r["S_max"]) else "-", axis=1)
        df_tampilan["Safety Stock (SS)"] = df_tampilan.apply(lambda r: format_indonesia_satuan(r["SS"], r["Satuan"]), axis=1)
        df_tampilan["Total Cost (OT)"] = df_tampilan["Total Cost (OT)"].apply(format_rupiah)

        kolom_tampil = ["Bahan Baku", "Kategori", "Satuan", "Stok Saat Ini Tampil", "Status Gudang", "s_opt (Batas Aman/r)", "Q_opt (Pemesanan)", "S_max (Target Maksimum)", "Safety Stock (SS)", "Total Cost (OT)"]
        st.dataframe(df_tampilan[kolom_tampil], use_container_width=True)

        fig = px.bar(
            df_hasil,
            x="Bahan Baku",
            y=["Stok Saat Ini", "s_opt"],
            barmode="group",
            title="Komparasi Visual Stok Riil Terhadap Batas Reorder Point (s)",
            labels={"value": "Volume / Jumlah", "variable": "Parameter Gudang"},
            color_discrete_sequence=["#3399FF", "#FF3333"],
        )
        st.plotly_chart(fig, use_container_width=True)

    with tab_po:
        st.markdown("<h2 style='font-size: 28px; font-weight: bold;'>📜 Generator Dokumen Purchase Order (PO)</h2>", unsafe_allow_html=True)

        if not PDF_AVAILABLE:
            st.error("⚠️ Library 'reportlab' belum terinstall. Silakan jalankan `pip install reportlab` di terminal Anda.")
        else:
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

                sub_po_tab1, sub_po_tab2 = st.tabs(["📄 1. Surat PO Resmi untuk Supplier (PDF)", "🔒 2. Draft Catatan Internal Perusahaan (PDF)"])

                with sub_po_tab1:
                    st.info("🔒 Setiap bahan baku diterbitkan dalam Surat PO PDF terpisah secara vertikal.")
                    
                    for i, bahan in enumerate(df_po["Nama Bahan Baku"].unique()):
                        row = df_po[df_po["Nama Bahan Baku"] == bahan].iloc[0]
                        no_po = f"PO/{bahan[:3].upper()}/{datetime.now().strftime('%Y%m%d')}/{i+1:02d}"
                        
                        df_item = pd.DataFrame([{
                            "No": 1, 
                            "Nama Bahan Baku": bahan, 
                            "Jumlah Pesanan": format_indonesia_satuan(row["Rekomendasi Pesan"], row["Satuan"])
                        }])
                        
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

                        bytes_pdf_po = buat_surat_po_pdf(
                            bahan=bahan,
                            kategori=row["Kategori"],
                            satuan=row["Satuan"],
                            qty_pesan=row["Rekomendasi Pesan"],
                            vendor_nama=f"PT. Supplier {bahan} Utama",
                            no_po=no_po,
                            petugas_nama=petugas_nama,
                            tanggal_str=today_str,
                            catatan=catatan_po
                        )

                        st.download_button(
                            f"📄 Download Surat PO ({bahan}) [.pdf]",
                            data=bytes_pdf_po,
                            file_name=f"Surat_PO_{bahan}.pdf",
                            mime="application/pdf",
                            type="primary",
                            key=f"dl_pdf_{i}"
                        )
                        st.markdown("<hr style='border:1px dashed #555;'><br>", unsafe_allow_html=True)

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
                        df_po_show["Stok Aktual Tampil"] = df_po_show.apply(lambda r: format_indonesia_satuan(r["Stok Aktual"], r["Satuan"]), axis=1)
                        df_po_show["Batas ROP Tampil"] = df_po_show.apply(lambda r: format_indonesia_satuan(r["Batas ROP"], r["Satuan"]), axis=1)
                        df_po_show["Rekomendasi Pesan Tampil"] = df_po_show.apply(lambda r: format_indonesia_satuan(r["Rekomendasi Pesan"], r["Satuan"]), axis=1)
                        df_po_show["Estimasi Biaya"] = df_po_show["Estimasi Biaya (Rp)"].apply(format_rupiah)

                        st.dataframe(
                            df_po_show[["Nama Bahan Baku", "Kategori", "Satuan", "Stok Aktual Tampil", "Batas ROP Tampil", "Rekomendasi Pesan Tampil", "Estimasi Biaya"]],
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

                    bytes_pdf_internal = buat_memo_internal_pdf(
                        df_po=df_po,
                        nomor_po=nomor_po,
                        petugas_nama=petugas_nama,
                        today_str=today_str
                    )

                    st.download_button(
                        label="📄 Download Draft Catatan Internal [.pdf]",
                        data=bytes_pdf_internal,
                        file_name=f"Draft_Internal_PO_{nomor_po.replace('/', '_')}.pdf",
                        mime="application/pdf",
                        key="btn_dl_pdf_int_total"
                    )
            else:
                st.success("✅ Tidak ada bahan baku yang berada di bawah Reorder Point (ROP). Belum ada draft PO yang perlu diterbitkan.")

    with tab2:
        st.markdown("### 📜 Log Laporan Mutasi Barang Gudang")
        df_log_tampil = muat_riwayat()
        if not df_log_tampil.empty and "Jumlah" in df_log_tampil.columns:
            df_filtered = df_log_tampil.sort_values(by="Waktu", ascending=False).copy()
            if "Satuan" not in df_filtered.columns:
                df_filtered["Satuan"] = "pcs"
            df_filtered["Jumlah Tampil"] = df_filtered.apply(lambda r: format_indonesia_satuan(r["Jumlah"], r["Satuan"]), axis=1)
            
            st.dataframe(df_filtered[["Waktu", "Bahan Baku", "Aktivitas", "Jumlah Tampil"]], use_container_width=True)

    with tab3:
        if st.button("Hapus & Reset Master Data Gudang 🗑️"):
            st.session_state["data_gudang"] = None
            st.session_state["stok_realtime"] = {}
            if os.path.exists(FILE_RIWAYAT):
                os.remove(FILE_RIWAYAT)
            st.rerun()

else:
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
        * **Apa itu?** Total jumlah pemakaian barang dalam 1 tahun.
        * **Cara dapatnya:** Jumlahkan total pemakaian dari bulan Januari sampai Desember.

        **2. Standar Deviasi / Sigma ($\sigma$) - Fluktuasi Pemakaian**
        * **Apa itu?** Angka yang menunjukkan seberapa fluktuatif pemakaian barang tiap bulannya.
        * **Cara dapatnya:** Gunakan rumus Excel `=STDEV.S(data_12_bulan)*SQRT(12)`.

        **3. Lead Time ($L$) - Waktu Tunggu Pengiriman**
        * **Apa itu?** Durasi waktu sejak barang dipesan sampai tiba di gudang.
        * **Cara dapatnya:** Jika supplier butuh waktu 30 hari, isi angka 30.
        """)

    with col_b:
        st.markdown("""
        **4. Biaya Pesan ($A$) - Ordering Cost**
        * **Apa itu?** Biaya yang keluar setiap kali Anda melakukan 1 kali pemesanan.

        **5. Biaya Simpan ($h$) - Holding Cost**
        * **Apa itu?** Biaya untuk menyimpan 1 unit barang di gudang selama 1 TAHUN.

        **6. Biaya Kekurangan ($C_u$) - Shortage Cost & Satuan**
        * **Biaya Kekurangan:** Estimasi biaya denda jika persediaan gudang HABIS saat produksi berjalan.
        * **Satuan (Kolom I):** Satuan ukur barang (contoh: `pcs`, `kg`, `box`, `roll`, `rim`).
        """)

    st.markdown("---")
    st.success(
        "💡 Bingung mulai dari mana? haha, Saya Sebagai Developer juga sempat linglung menatap Excel kosong cukup lama sebelum sistem ini jadi 😵‍💫. Daripada mengulangi perjuangan yang sama, klik tombol **'📥 Unduh Template Excel Otomatis'** di Panel sebelah kiri dan biarkan template yang kerja keras. Saya sudah capek ngodingnya, Anda tidak perlu capek bikin kolomnya. Cukup Unduh & Isi, lalu kirim dan nikmati hidup tanpa drama 'kolom tidak sesuai format'.😝"
    )
