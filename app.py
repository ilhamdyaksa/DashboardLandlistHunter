import streamlit as st
import pandas as pd
import plotly.express as px
import io
from streamlit_gsheets import GSheetsConnection

# Konfigurasi Halaman
st.set_page_config(
    page_title="Dashboard Pipeline Landlist VGREEN",
    page_icon="⚡",
    layout="wide"
)

# ---------------------------------------------------------
# KONEKSI KE GOOGLE SHEETS (Satu Sumber Kebenaran)
# ---------------------------------------------------------
# Membuka koneksi menggunakan st.connection (Butuh library st-gsheets-connection & .streamlit/secrets.toml)
conn = st.connection("gsheets", type=GSheetsConnection)

# Membaca data dari Google Sheets secara langsung (ttl=0 memastikan data fresh setiap reload)
try:
    df = conn.read(worksheet="Sheet1", ttl=0)
    # Bersihkan baris kosong (NaN) yang mungkin terbaca dari Sheets
    df = df.dropna(how="all")
except Exception as e:
    st.error(f"Gagal terhubung ke Google Sheets. Pastikan kredensial dan URL di secrets.toml benar. Error: {e}")
    st.stop()

# ---------------------------------------------------------
# DATABASE KOTA & PROVINSI INDONESIA
# ---------------------------------------------------------
wilayah_indonesia = {
    "Nanggroe Aceh Darussalam": ["Banda Aceh", "Lhokseumawe", "Aceh Besar", "Pidie"],
    "Sumatera Utara": ["Medan", "Deli Serdang", "Binjai", "Pematangsiantar"],
    "Sumatera Barat": ["Padang", "Bukittinggi", "Payakumbuh", "Pariaman"],
    "Riau": ["Pekanbaru", "Dumai", "Bengkalis", "Kampar"],
    "Kepulauan Riau": ["Batam", "Tanjungpinang", "Bintan", "Karimun"],
    "Jambi": ["Jambi", "Sungai Penuh", "Batanghari", "Bungo"],
    "Sumatera Selatan": ["Palembang", "Lubuklinggau", "Prabumulih", "Banyuasin"],
    "Bangka Belitung": ["Pangkalpinang", "Bangka", "Belitung"],
    "Bengkulu": ["Bengkulu", "Rejang Lebong", "Mukomuko"],
    "Lampung": ["Bandar Lampung", "Metro", "Lampung Selatan", "Lampung Tengah"],
    "DKI Jakarta": ["Jakarta Pusat", "Jakarta Utara", "Jakarta Barat", "Jakarta Selatan", "Jakarta Timur", "Kepulauan Seribu"],
    "Jawa Barat": ["Bandung", "Bogor", "Depok", "Bekasi", "Karawang", "Cianjur", "Garut"],
    "Banten": ["Serang", "Tangerang", "Tangerang Selatan", "Cilegon", "Pandeglang"],
    "Jawa Tengah": ["Semarang", "Surakarta", "Magelang", "Banyumas", "Klaten"],
    "DI Yogyakarta": ["Yogyakarta", "Sleman", "Bantul", "Gunungkidul", "Kulon Progo"],
    "Jawa Timur": ["Surabaya", "Malang", "Sidoarjo", "Gresik", "Kediri"],
    "Bali": ["Denpasar", "Badung", "Gianyar", "Buleleng", "Tabanan"],
    "Nusa Tenggara Barat": ["Mataram", "Lombok Barat", "Lombok Tengah", "Sumbawa"],
    "Nusa Tenggara Timur": ["Kupang", "Manggarai", "Sikka", "Ende"],
    "Kalimantan Barat": ["Pontianak", "Singkawang", "Kubu Raya", "Ketapang"],
    "Kalimantan Tengah": ["Palangka Raya", "Kotawaringin Timur", "Kapuas"],
    "Kalimantan Selatan": ["Banjarmasin", "Banjarbaru", "Banjar", "Tabalong"],
    "Kalimantan Timur": ["Samarinda", "Balikpapan", "Bontang", "Kutai Kartanegara"],
    "Kalimantan Utara": ["Tarakan", "Bulungan", "Nunukan"],
    "Sulawesi Utara": ["Manado", "Bitung", "Tomohon", "Minahasa"],
    "Gorontalo": ["Gorontalo", "Bone Bolango", "Boalemo"],
    "Sulawesi Tengah": ["Palu", "Donggala", "Poso", "Banggai"],
    "Sulawesi Barat": ["Mamuju", "Majene", "Polewali Mandar"],
    "Sulawesi Selatan": ["Makassar", "Parepare", "Palopo", "Gowa"],
    "Sulawesi Tenggara": ["Kendari", "Baubau", "Konawe", "Kolaka"],
    "Maluku": ["Ambon", "Tual", "Maluku Tengah", "Buru"],
    "Maluku Utara": ["Ternate", "Tidore Kepulauan", "Halmahera Utara"],
    "Papua": ["Jayapura", "Biak Numfor", "Keerom"],
    "Papua Barat": ["Manokwari", "Sorong", "Fakfak"],
    "Papua Selatan": ["Merauke", "Boven Digoel", "Mappi"],
    "Papua Tengah": ["Nabire", "Mimika", "Puncak Jaya"],
    "Papua Pegunungan": ["Jayawijaya", "Yahukimo", "Tolikara"],
    "Papua Barat Daya": ["Sorong", "Raja Ampat", "Tambrauw"]
}
all_cities = [city for cities in wilayah_indonesia.values() for city in cities]

# ---------------------------------------------------------
# FUNGSI EXPORT EXCEL (UNTUK TEMPLATE)
# ---------------------------------------------------------
def generate_excel_template():
    template_df = pd.DataFrame(columns=[
        "Site ID", "Nama Lokasi", "Provinsi", "Kota/Kabupaten",
        "Jumlah EVCS", "Jumlah BSS", "PIC Hunter", 
        "Status Approach", "Status Deal", "Kelengkapan Dokumen", 
        "Approval VGREEN", "Tanggal Update"
    ])
    
    template_df.loc[0] = [
        "VGR-CONTOH-01", "SPBU Contoh", "Jawa Barat", "Bogor", 
        2, 1, "Tim Contoh", 
        "Done", "Deal", "Submitted Complete", 
        "Approved", "2026-10-01"
    ]
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        template_df.to_excel(writer, index=False, sheet_name='Template_Landlist')
    processed_data = output.getvalue()
    return processed_data

# Header Dashboard
st.title("⚡ Dashboard Pipeline Pengadaan Lahan VGREEN")
st.caption("Monitoring status pipeline: Approach -> Deal -> Dokumen Lengkap -> Approval VGREEN (Synced with Google Sheets)")
st.markdown("---")

# ---------------------------------------------------------
# 1. KPI METRICS
# ---------------------------------------------------------
total_titik = len(df)
total_approach = len(df[df["Status Approach"] == "Done"]) if total_titik > 0 else 0
total_deal = len(df[df["Status Deal"] == "Deal"]) if total_titik > 0 else 0
total_doc = len(df[df["Kelengkapan Dokumen"] == "Submitted Complete"]) if total_titik > 0 else 0
total_approved = len(df[df["Approval VGREEN"] == "Approved"]) if total_titik > 0 else 0
total_evcs = df["Jumlah EVCS"].sum() if total_titik > 0 else 0
total_bss = df["Jumlah BSS"].sum() if total_titik > 0 else 0

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("📍 Total Titik Lahan", total_titik)
col2.metric("🤝 Lahan Approached", total_approach, f"{(total_approach/total_titik*100 if total_titik else 0):.0f}%")
col3.metric("✍️ Lahan Sudah Deal", total_deal, f"{(total_deal/total_titik*100 if total_titik else 0):.0f}%")
col4.metric("📄 Dokumen Lengkap", total_doc, f"{(total_doc/total_titik*100 if total_titik else 0):.0f}%")
col5.metric("✅ Approved VGREEN", total_approved, f"{(total_approved/total_titik*100 if total_titik else 0):.0f}%")

st.markdown("<br>", unsafe_allow_html=True)
st.caption("### Rencana Kuantitas Perangkat")
metric_col1, metric_col2 = st.columns(2)
metric_col1.metric("🔌 Total EVCS Target", total_evcs)
metric_col2.metric("🔋 Total BSS Target", total_bss)

st.markdown("---")

# ---------------------------------------------------------
# 2. VISUALISASI: FUNNEL & FISHBONE
# ---------------------------------------------------------
chart_col1, chart_col2 = st.columns([1, 1])

with chart_col1:
    st.subheader("Funnel Pipeline Status Lahan")
    funnel_data = pd.DataFrame({
        "Tahapan": ["Total Terdaftar", "Approached", "Deal Sepakat", "Dokumen Lengkap", "Approved VGREEN"],
        "Jumlah": [total_titik, total_approach, total_deal, total_doc, total_approved]
    })
    fig_funnel = px.funnel(funnel_data, x="Jumlah", y="Tahapan", color="Tahapan",
                           color_discrete_sequence=px.colors.sequential.Teal)
    fig_funnel.update_layout(showlegend=False, height=450)
    st.plotly_chart(fig_funnel, use_container_width=True)

with chart_col2:
    st.subheader("Analisis Kendala Approval (Fishbone / Root Cause)")
    fishbone_data = dict(
        character=["Gagal Capai Target", "Manusia (Tim)", "Metode", "Material (Lahan)", "Mesin/Sistem", "Lingkungan",
                   "Kapasitas Nego Kurang", "Telat Update Status", "Screening Lama", "Birokrasi Harga", 
                   "Dokumen Tdk Lengkap", "Harga Overpriced", "Portal Error", "Aplikasi Delay", 
                   "Penolakan Warga", "Sengketa Lahan"],
        parent=["", "Gagal Capai Target", "Gagal Capai Target", "Gagal Capai Target", "Gagal Capai Target", "Gagal Capai Target",
                "Manusia (Tim)", "Manusia (Tim)", "Metode", "Metode", 
                "Material (Lahan)", "Material (Lahan)", "Mesin/Sistem", "Mesin/Sistem", 
                "Lingkungan", "Lingkungan"],
        value=[100, 20, 20, 20, 20, 20, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10]
    )
    fig_fishbone = px.sunburst(
        fishbone_data, names='character', parents='parent', values='value',
        title="Struktur Akar Penyebab (Klik untuk membedah kategori)",
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    fig_fishbone.update_layout(height=450, margin=dict(t=30, l=0, r=0, b=0))
    st.plotly_chart(fig_fishbone, use_container_width=True)

# ---------------------------------------------------------
# 3. DAILY PROGRESS CHART
# ---------------------------------------------------------
st.markdown("---")
st.subheader("📈 Daily Progress Chart (Pergerakan Harian Lahan)")

if not df.empty and "Tanggal Update" in df.columns:
    df_chart = df.copy()
    df_chart['Tanggal Update'] = pd.to_datetime(df_chart['Tanggal Update']).dt.date
    daily_data = df_chart.groupby(['Tanggal Update', 'Approval VGREEN']).size().reset_index(name='Jumlah Titik')

    fig_daily = px.line(
        daily_data, x="Tanggal Update", y="Jumlah Titik", color="Approval VGREEN",
        markers=True, title="Tren Pergerakan Status Lahan Harian",
        labels={"Tanggal Update": "Tanggal", "Jumlah Titik": "Total Lahan"},
        color_discrete_sequence=px.colors.qualitative.Set1
    )
    fig_daily.update_xaxes(type='category', tickformat="%Y-%m-%d")
    fig_daily.update_layout(height=400, xaxis_title="Tanggal", yaxis_title="Jumlah Titik Lahan")
    st.plotly_chart(fig_daily, use_container_width=True)
else:
    st.info("Belum ada data untuk menampilkan grafik pergerakan harian.")

st.markdown("---")

# ---------------------------------------------------------
# 4. SIDEBAR: FORM MANUAL & UPLOAD EXCEL (WRITE TO GOOGLE SHEETS)
# ---------------------------------------------------------
st.sidebar.header("📥 Input Massal (Upload Excel)")

excel_template = generate_excel_template()
st.sidebar.download_button(
    label="1. ⬇️ Download Template Excel",
    data=excel_template,
    file_name="Template_Input_Landlist.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.sidebar.markdown("2. Upload File Excel yang sudah diisi:")
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_uploaded = pd.read_excel(uploaded_file)
        expected_columns = [
            "Site ID", "Nama Lokasi", "Provinsi", "Kota/Kabupaten",
            "Jumlah EVCS", "Jumlah BSS", "PIC Hunter", 
            "Status Approach", "Status Deal", "Kelengkapan Dokumen", 
            "Approval VGREEN", "Tanggal Update"
        ]
        
        if all(col in df_uploaded.columns for col in expected_columns):
            df_uploaded['Tanggal Update'] = pd.to_datetime(df_uploaded['Tanggal Update']).dt.strftime('%Y-%m-%d')
            df_uploaded = df_uploaded.fillna("")
            
            if st.sidebar.button("💾 Konfirmasi & Simpan ke Google Sheets"):
                with st.spinner('Menyimpan ke Google Sheets...'):
                    # 1. Hapus duplikat di dalam file Excel yang diunggah itu sendiri
                    df_uploaded = df_uploaded.drop_duplicates(subset=['Site ID'], keep='last')
                    
                    # 2. Filter data baru: Hanya ambil Site ID yang BELUM ada di Google Sheets (df)
                    if not df.empty and 'Site ID' in df.columns:
                        existing_site_ids = df['Site ID'].tolist()
                    else:
                        existing_site_ids = []
                        
                    df_baru = df_uploaded[~df_uploaded['Site ID'].isin(existing_site_ids)]
                    
                    if not df_baru.empty:
                        # Gabungkan data lama dan baru (yang sudah terfilter), lalu update ke Sheets
                        updated_df = pd.concat([df, df_baru], ignore_index=True)
                        conn.update(worksheet="Sheet1", data=updated_df)
                        st.cache_data.clear() # Bersihkan cache agar data terbaru langsung ditarik
                        st.sidebar.success(f"Berhasil menambahkan {len(df_baru)} data baru ke Google Sheets!")
                    else:
                        st.sidebar.warning("Tidak ada data baru ditambahkan. Semua Site ID dalam file Excel sudah ada di database.")
                st.rerun()
        else:
            st.sidebar.error("Format kolom tidak sesuai dengan template!")
    except Exception as e:
        st.sidebar.error(f"Error membaca file: {e}")

st.sidebar.markdown("---")

st.sidebar.header("➕ Input Manual (Satu per Satu)")
with st.sidebar.form("form_tambah_lahan", clear_on_submit=True):
    site_id = st.text_input("Site ID (Contoh: VGR-BGR-004)")
    nama_lokasi = st.text_input("Nama Lokasi / Calon Partner")
    
    provinsi_input = st.selectbox("Provinsi", list(wilayah_indonesia.keys()))
    kota_input = st.selectbox("Kota / Kabupaten", all_cities)
    
    st.markdown("**Kuantitas Perangkat**")
    col_evcs, col_bss = st.columns(2)
    with col_evcs:
        jumlah_evcs = st.number_input("Jumlah EVCS", min_value=0, value=0, step=1)
    with col_bss:
        jumlah_bss = st.number_input("Jumlah BSS", min_value=0, value=0, step=1)
        
    pic_hunter = st.text_input("PIC Hunter")
    status_approach = st.selectbox("Status Approach", ["Not Started", "In Progress", "Done"])
    status_deal = st.selectbox("Status Deal", ["Pending", "Negotiating", "Deal", "Drop"])
    status_dokumen = st.selectbox("Kelengkapan Dokumen", ["Pending", "Incomplete", "Submitted Complete"])
    status_vgreen = st.selectbox("Approval VGREEN", ["Pending", "Under Review", "Revision", "Approved", "Rejected"])
    tanggal_update = st.date_input("Tanggal Update")
    
    submitted = st.form_submit_button("Simpan Manual")
    if submitted:
        if site_id and nama_lokasi:
            # Pengecekan Duplikasi: Apakah Site ID sudah ada di database?
            if not df.empty and 'Site ID' in df.columns and site_id in df['Site ID'].values:
                st.error(f"Gagal Menyimpan! Site ID '{site_id}' sudah terdaftar di database.")
            else:
                new_row = pd.DataFrame([{
                    "Site ID": site_id, "Nama Lokasi": nama_lokasi, "Provinsi": provinsi_input,
                    "Kota/Kabupaten": kota_input, "Jumlah EVCS": int(jumlah_evcs), "Jumlah BSS": int(jumlah_bss),
                    "PIC Hunter": pic_hunter, "Status Approach": status_approach, "Status Deal": status_deal,
                    "Kelengkapan Dokumen": status_dokumen, "Approval VGREEN": status_vgreen,
                    "Tanggal Update": pd.to_datetime(tanggal_update).strftime('%Y-%m-%d')
                }])
                
                with st.spinner('Menyimpan ke Google Sheets...'):
                    updated_df = pd.concat([df, new_row], ignore_index=True)
                    conn.update(worksheet="Sheet1", data=updated_df)
                    st.cache_data.clear()
                    
                st.success(f"Titik {site_id} berhasil disimpan ke Google Sheets!")
                st.rerun()
        else:
            st.error("Site ID dan Nama Lokasi wajib diisi.")

# ---------------------------------------------------------
# 5. TABEL DETAIL & FILTER
# ---------------------------------------------------------
st.subheader("📋 Database & Tracking Detail Lahan")

if not df.empty:
    filter_col1, filter_col2, filter_col3 = st.columns(3)
    with filter_col1:
        filter_prov = st.multiselect("Filter Provinsi", options=df["Provinsi"].unique(), default=df["Provinsi"].unique())
    with filter_col2:
        filter_kota = st.multiselect("Filter Kota/Kabupaten", options=df["Kota/Kabupaten"].unique(), default=df["Kota/Kabupaten"].unique())
    with filter_col3:
        filter_approval = st.multiselect("Filter Status Approval VGREEN", options=df["Approval VGREEN"].unique(), default=df["Approval VGREEN"].unique())

    filtered_df = df[
        (df["Provinsi"].isin(filter_prov)) &
        (df["Kota/Kabupaten"].isin(filter_kota)) & 
        (df["Approval VGREEN"].isin(filter_approval))
    ]

    st.dataframe(filtered_df, use_container_width=True)

    csv = filtered_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Unduh Data saat ini (CSV)",
        data=csv,
        file_name='landlist_vgreen_progress.csv',
        mime='text/csv'
    )
else:
    st.info("Database masih kosong. Silakan input data melalui form atau upload Excel di sebelah kiri.")