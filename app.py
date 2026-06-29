import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import json
import os
from copy import deepcopy
import zipfile
import xml.etree.ElementTree as ET
import streamlit.components.v1 as components

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

# ==========================================
# KONFIGURASI HALAMAN STREAMLIT
# ==========================================
st.set_page_config(page_title="WebGIS Pendidikan Medan", layout="wide")
st.title("🗺️ Pemetaan Kesenjangan Kapasitas Pendidikan Dasar Negeri Kota Medan")
st.markdown("""
Aplikasi WebGIS ini mengelompokkan 21 Kecamatan di Kota Medan menggunakan algoritma **K-Means++** berdasarkan data kapasitas infrastruktur (Jumlah Sekolah, Guru, dan Peserta Didik) dari sistem Dapodik.
""")

TAB_INPUT, TAB_PREVIEW = st.tabs(["1. Input & Proses", "2. Preview WebGIS"])

GEOJSON_PATH = "medan_kecamatan.geojson"

if "processed_df" not in st.session_state:
    st.session_state["processed_df"] = None
if "raw_df" not in st.session_state:
    st.session_state["raw_df"] = None
if "raw_file_names" not in st.session_state:
    st.session_state["raw_file_names"] = []
if "geojson_data" not in st.session_state:
    st.session_state["geojson_data"] = None
if "geojson_name" not in st.session_state:
    st.session_state["geojson_name"] = None


st.markdown(
    """
    <style>
    .section-card {
        background: linear-gradient(135deg, rgba(13,110,253,0.08), rgba(25,135,84,0.08));
        border: 1px solid rgba(0,0,0,0.06);
        border-radius: 16px;
        padding: 16px 18px;
        margin-bottom: 12px;
    }
    .small-note {
        color: #6c757d;
        font-size: 0.92rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_geojson_from_disk(path, mtime):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_geojson_data():
    if "geojson_data" in st.session_state:
        if st.session_state["geojson_data"] is not None:
            return st.session_state["geojson_data"]

    if os.path.exists(GEOJSON_PATH):
        mtime = os.path.getmtime(GEOJSON_PATH)
        return load_geojson_from_disk(GEOJSON_PATH, mtime)

    return None


def get_geojson_name_field(geojson_data):
    if not geojson_data or "features" not in geojson_data or not geojson_data["features"]:
        return "KECAMATAN"

    properties = geojson_data["features"][0].get("properties", {})
    for key in ["KECAMATAN", "kecamatan", "Kecamatan", "NAMA_KEC", "NAMA", "NAME"]:
        if key in properties:
            return key

    return list(properties.keys())[0] if properties else "KECAMATAN"


def update_geojson_from_upload(uploaded_geojson):
    geojson_bytes = uploaded_geojson.getvalue()
    geojson_obj = json.loads(geojson_bytes.decode("utf-8"))

    with open(GEOJSON_PATH, "wb") as f:
        f.write(geojson_bytes)

    st.session_state["geojson_data"] = geojson_obj
    st.session_state["geojson_name"] = uploaded_geojson.name
    return geojson_obj

def read_dapodik_file(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        return normalize_dapodik_columns(df)

    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        return normalize_dapodik_columns(df)
    except Exception:
        uploaded_file.seek(0)
        df = read_xlsx_without_openpyxl(uploaded_file)
        return normalize_dapodik_columns(df)


def normalize_dapodik_columns(df):
    def clean_text(value):
        return str(value).replace("\ufeff", "").replace("\n", " ").replace("\r", " ").strip().lower()

    def infer_column(target_names, substrings):
        for col in df.columns:
            key = clean_text(col)
            if key in target_names or any(sub in key for sub in substrings):
                return col
        return None

    rename_map = {}

    kecamatan_col = infer_column(
        {"kecamatan", "kec"},
        ["kecamatan", "kecamatan", "kec ", "kec.", "wilayah", "district"],
    )
    sekolah_col = infer_column(
        {"nama sekolah", "sekolah"},
        ["nama sekolah", "sekolah", "school"],
    )
    pd_col = infer_column(
        {"pd", "peserta didik", "siswa"},
        ["pd", "peserta didik", "jumlah siswa", "siswa"],
    )
    guru_col = infer_column(
        {"guru"},
        ["guru", "pendidik"],
    )

    if kecamatan_col is not None:
        rename_map[kecamatan_col] = "Kecamatan"
    if sekolah_col is not None:
        rename_map[sekolah_col] = "Nama Sekolah"
    if pd_col is not None:
        rename_map[pd_col] = "PD"
    if guru_col is not None:
        rename_map[guru_col] = "Guru"

    df = df.rename(columns=rename_map)

    # Bersihkan nama kolom supaya konsisten
    df.columns = [str(col).replace("\ufeff", "").strip() for col in df.columns]
    
    # =====================================================================
    # PERBAIKAN: HAPUS KOLOM DUPLIKAT YANG MEMBUAT PANDAS CRASH
    # Mengabaikan kolom-kolom ganda/tersembunyi bawaan excel Dapodik
    # =====================================================================
    df = df.loc[:, ~df.columns.duplicated()]
    
    return df

def read_xlsx_without_openpyxl(uploaded_file):
    ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    ns_rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    ns_doc_rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

    def col_to_index(col_letters):
        result = 0
        for ch in col_letters:
            result = result * 26 + (ord(ch.upper()) - ord("A") + 1)
        return result - 1

    def cell_value(cell, shared_strings):
        cell_type = cell.attrib.get("t")
        value_node = cell.find(f"{ns_main}v")
        if value_node is None:
            inline = cell.find(f"{ns_main}is/{ns_main}t")
            return inline.text if inline is not None else None

        value = value_node.text
        if value is None:
            return None

        if cell_type == "s":
            try:
                return shared_strings[int(value)]
            except Exception:
                return value
        if cell_type == "b":
            return value == "1"
        if cell_type == "str":
            return value
        try:
            num = float(value)
            return int(num) if num.is_integer() else num
        except Exception:
            return value

    uploaded_file.seek(0)
    with zipfile.ZipFile(uploaded_file) as zf:
        workbook_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

        rels = {}
        for rel in rels_xml.findall(f"{ns_rel}Relationship"):
            rels[rel.attrib["Id"]] = rel.attrib["Target"]

        sheet = workbook_xml.find(f"{ns_main}sheets/{ns_main}sheet")
        if sheet is None:
            return pd.DataFrame()

        rel_id = sheet.attrib.get(f"{ns_doc_rel}id")
        sheet_target = rels.get(rel_id)
        if not sheet_target:
            return pd.DataFrame()

        shared_strings = []
        try:
            shared_xml = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in shared_xml.findall(f"{ns_main}si"):
                text_parts = [t.text or "" for t in si.findall(f".//{ns_main}t")]
                shared_strings.append("".join(text_parts))
        except KeyError:
            shared_strings = []

        sheet_xml = ET.fromstring(zf.read(f"xl/{sheet_target}"))
        rows = []
        for row in sheet_xml.findall(f".//{ns_main}row"):
            row_values = {}
            max_col = -1
            for cell in row.findall(f"{ns_main}c"):
                ref = cell.attrib.get("r", "")
                letters = "".join(ch for ch in ref if ch.isalpha())
                if not letters:
                    continue
                idx = col_to_index(letters)
                max_col = max(max_col, idx)
                row_values[idx] = cell_value(cell, shared_strings)

            if max_col >= 0:
                rows.append([row_values.get(i) for i in range(max_col + 1)])

        if not rows:
            return pd.DataFrame()

        expected_keywords = ["npsn", "nama sekolah", "pd", "guru"]

        def row_score(row):
            values = [str(x).strip().lower() for x in row if x not in (None, "")]
            score = 0
            for keyword in expected_keywords:
                if any(keyword in value for value in values):
                    score += 1
            return score, len(values)

        header_idx = 0
        best_score = (-1, -1)
        for idx, row in enumerate(rows[:20]):
            score = row_score(row)
            if score > best_score:
                best_score = score
                header_idx = idx

        headers_raw = rows[header_idx]
        data_rows = rows[header_idx + 1 :]

        widths = [len(headers_raw)]
        widths.extend(len(r) for r in data_rows)
        max_width = max(widths)
        headers = []
        for i in range(max_width):
            value = headers_raw[i] if i < len(headers_raw) else None
            header_name = str(value).strip() if value not in (None, "") else f"Unnamed_{i}"
            headers.append(header_name)

        normalized_data = []
        for row in data_rows:
            normalized_row = list(row) + [None] * (max_width - len(row))
            normalized_data.append(normalized_row[:max_width])

        if not normalized_data:
            return pd.DataFrame(columns=headers)

        return pd.DataFrame(normalized_data, columns=headers)

def process_dapodik_data(df):
    # =========================================================================
    # PERBAIKAN 1: Ekstrak Kecamatan dari Nama File jika tidak ada di Excel
    # =========================================================================
    if "Kecamatan" not in df.columns and "Sumber_File" in df.columns:
        # Menghapus ekstensi file (misal: "Kec. Medan Barat.xlsx" menjadi "Kec. Medan Barat")
        df["Kecamatan"] = df["Sumber_File"].apply(lambda x: os.path.splitext(str(x))[0].strip())
    
    # =========================================================================
    # PERBAIKAN 2: Hapus baris 'Total' yang sering ikut terdownload di Dapodik
    # =========================================================================
    if "Nama Sekolah" in df.columns:
        df = df[~df["Nama Sekolah"].astype(str).str.lower().str.contains("total", na=False)]
    if "NPSN" in df.columns:
        df = df[df["NPSN"].notna()] # Menghapus baris jika kolom NPSN kosong
        
    # Pastikan kolom wajib terpenuhi
    required_columns = ["Kecamatan", "PD", "Guru"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            f"Kolom wajib tidak ditemukan: {missing_columns}. Kolom yang tersedia: {list(df.columns)}"
        )

    df = df.dropna(subset=["Kecamatan", "PD", "Guru"])
    df["PD"] = pd.to_numeric(df["PD"], errors="coerce")
    df["Guru"] = pd.to_numeric(df["Guru"], errors="coerce")
    df = df.dropna(subset=["Kecamatan", "PD", "Guru"])

    df_agg = df.groupby("Kecamatan").agg(
        Jumlah_Sekolah=("Nama Sekolah", "count"),
        Jumlah_PD=("PD", "sum"),
        Jumlah_Guru=("Guru", "sum"),
    ).reset_index()

    df_agg["Rasio_PD_Sekolah"] = df_agg["Jumlah_PD"] / df_agg["Jumlah_Sekolah"]
    df_agg["Rasio_PD_Guru"] = df_agg["Jumlah_PD"] / df_agg["Jumlah_Guru"]

    X = df_agg[["Jumlah_Sekolah", "Rasio_PD_Sekolah", "Rasio_PD_Guru", "Jumlah_PD"]]
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    kmeans = KMeans(n_clusters=4, init="k-means++", random_state=42, n_init=10)
    df_agg["Klaster"] = kmeans.fit_predict(X_scaled)

    klaster_label = {
        0: "Klaster 0 (Sangat Kritis)",
        1: "Klaster 1 (Kurang Memadai)",
        2: "Klaster 2 (Memadai / Aman)",
        3: "Klaster 3 (Sangat Berlebih)",
    }
    df_agg["Nama_Klaster"] = df_agg["Klaster"].map(klaster_label)
    return df_agg


def prepare_map_geojson(geojson_data, df_agg, geojson_name_field):
    geojson_map = deepcopy(geojson_data)
    lookup = df_agg.set_index("Kecamatan").to_dict(orient="index")

    for feature in geojson_map.get("features", []):
        props = feature.setdefault("properties", {})
        kecamatan = props.get(geojson_name_field)
        row = lookup.get(kecamatan)
        if row:
            props.update(row)
        else:
            props["Jumlah_Sekolah"] = None
            props["Jumlah_PD"] = None
            props["Jumlah_Guru"] = None
            props["Klaster"] = None
            props["Nama_Klaster"] = "Data tidak tersedia"

    return geojson_map


def build_map_html(geojson_data, df_agg, geojson_name_field):
    m = folium.Map(location=[3.5952, 98.6722], zoom_start=11, tiles="CartoDB positron", prefer_canvas=True)

    # Fokus pada 4 klaster utama agar peta ringan dan konsisten dengan hasil kalkulasi
    cluster_colors = {
        0: "#d73027",  # sangat kritis
        1: "#fc8d59",  # kurang memadai
        2: "#91cf60",  # memadai
        3: "#1a9850",  # sangat berlebih
    }

    def style_function(feature):
        klaster = feature["properties"].get("Klaster")
        return {
            "fillColor": cluster_colors.get(klaster, "#9e9e9e"),
            "color": "#5a5a5a",
            "weight": 1,
            "fillOpacity": 0.75,
        }

    tooltip = folium.GeoJsonTooltip(
        fields=[geojson_name_field, "Jumlah_Sekolah", "Jumlah_PD", "Jumlah_Guru", "Nama_Klaster"],
        aliases=["Kecamatan", "Jumlah Sekolah", "Jumlah PD", "Jumlah Guru", "Status Klaster"],
        localize=True,
        sticky=False,
        labels=True,
    )

    popup = folium.GeoJsonPopup(
        fields=[geojson_name_field, "Jumlah_Sekolah", "Jumlah_PD", "Jumlah_Guru", "Nama_Klaster"],
        aliases=["Kecamatan", "Jumlah Sekolah", "Jumlah PD", "Jumlah Guru", "Status Klaster"],
        localize=True,
        labels=True,
        style="background-color: white;",
        max_width=320,
    )

    folium.GeoJson(
        geojson_data,
        name="Klaster Kecamatan",
        style_function=style_function,
        tooltip=tooltip,
        popup=popup,
    ).add_to(m)

    legend_html = """
    <div style="
        position: fixed;
        bottom: 35px;
        left: 35px;
        z-index: 9999;
        background: white;
        border: 1px solid rgba(0,0,0,0.12);
        border-radius: 12px;
        padding: 12px 14px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.12);
        font-size: 13px;
        min-width: 210px;
    ">
        <div style="font-weight: 700; margin-bottom: 8px;">Legenda Klaster</div>
        <div><span style="display:inline-block;width:12px;height:12px;background:#d73027;margin-right:8px;border-radius:3px"></span>Sangat Kritis</div>
        <div><span style="display:inline-block;width:12px;height:12px;background:#fc8d59;margin-right:8px;border-radius:3px"></span>Kurang Memadai</div>
        <div><span style="display:inline-block;width:12px;height:12px;background:#91cf60;margin-right:8px;border-radius:3px"></span>Memadai / Aman</div>
        <div><span style="display:inline-block;width:12px;height:12px;background:#1a9850;margin-right:8px;border-radius:3px"></span>Sangat Berlebih</div>
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))

    folium.LayerControl(position="topright").add_to(m)
    return m.get_root().render()


with TAB_INPUT:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("📥 Upload Data dan GeoJSON")
    st.caption("Data Dapodik disimpan dulu. Proses klasterisasi hanya dijalankan saat tombol ditekan.")
    col_left, col_right = st.columns(2)

    with col_left:
        uploaded_files = st.file_uploader(
            "1) Upload Data Dapodik (CSV/Excel) - bisa banyak file sekaligus",
            type=["csv", "xlsx"],
            accept_multiple_files=True,
        )

        if uploaded_files:
            st.caption("File yang dipilih akan disimpan dulu. Proses klasterisasi hanya berjalan saat tombol dipencet.")
            st.write("File aktif:")
            st.write([f.name for f in uploaded_files])

            if st.button("Simpan file Dapodik"):
                try:
                    df_list = []
                    file_names = []
                    for file in uploaded_files:
                        df_temp = read_dapodik_file(file)
                        df_temp["Sumber_File"] = file.name
                        df_list.append(df_temp)
                        file_names.append(file.name)

                    raw_df = pd.concat(df_list, ignore_index=True)
                    st.session_state["raw_df"] = raw_df
                    st.session_state["raw_file_names"] = file_names
                    st.session_state["processed_df"] = None

                    st.success(f"✅ {len(uploaded_files)} file berhasil disimpan. Data lama sudah diganti.")
                    st.info("Silakan klik tombol Proses Data Dapodik untuk menjalankan klasterisasi.")
                except Exception as e:
                    st.error(f"⚠️ Terjadi kesalahan saat menyimpan file: {e}")

    with col_right:
        uploaded_geojson = st.file_uploader(
            "2) Upload GeoJSON Batas Kecamatan",
            type=["geojson", "json"],
        )

        if uploaded_geojson is not None:
            if st.button("Gunakan / simpan GeoJSON yang diupload"):
                try:
                    geojson_obj = update_geojson_from_upload(uploaded_geojson)
                    st.success(f"✅ GeoJSON berhasil diperbarui dari {uploaded_geojson.name}")
                    st.caption(f"Aktif: {uploaded_geojson.name}")
                    st.session_state["geojson_data"] = geojson_obj
                except Exception as e:
                    st.error(f"⚠️ GeoJSON gagal diperbarui: {e}")

    st.info("Upload satu atau beberapa file Dapodik sekaligus. Hasilnya akan digabung, diproses, lalu ditampilkan di tab preview.")

    with st.expander("Status file aktif", expanded=True):
        if st.session_state["raw_file_names"]:
            st.success(f"File Dapodik tersimpan: {len(st.session_state['raw_file_names'])} file")
            st.caption(", ".join(st.session_state["raw_file_names"]))
        else:
            st.warning("Belum ada file Dapodik yang disimpan.")

        if st.session_state["geojson_name"]:
            st.success(f"GeoJSON aktif: {st.session_state['geojson_name']}")
        elif os.path.exists(GEOJSON_PATH):
            st.success(f"GeoJSON aktif: {GEOJSON_PATH}")
        else:
            st.warning("GeoJSON belum diupload. Silakan upload file batas kecamatan.")

        if st.session_state["raw_df"] is not None:
            st.success("Data Dapodik sudah tersimpan dan siap diproses.")

        if st.button("Proses Data Dapodik", type="primary", disabled=st.session_state["raw_df"] is None):
            try:
                with st.spinner("Memproses data dan menjalankan algoritma K-Means++..."):
                    processed_df = process_dapodik_data(st.session_state["raw_df"])
                    st.session_state["processed_df"] = processed_df

                st.success("✅ Data berhasil diproses.")
                st.subheader("📊 Hasil Klasterisasi per Kecamatan")
                st.dataframe(
                    processed_df[["Kecamatan", "Jumlah_Sekolah", "Jumlah_PD", "Jumlah_Guru", "Nama_Klaster"]],
                    use_container_width=True,
                )

                csv_bytes = processed_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="Unduh hasil klasterisasi (CSV)",
                    data=csv_bytes,
                    file_name="hasil_klasterisasi_kecamatan_medan.csv",
                    mime="text/csv",
                )
            except Exception as e:
                st.error(f"⚠️ Terjadi kesalahan saat memproses file: {e}")

        if st.session_state["processed_df"] is None:
            st.info("👈 Upload lalu simpan file Dapodik terlebih dahulu, kemudian tekan tombol Proses Data Dapodik.")
    st.markdown('</div>', unsafe_allow_html=True)


with TAB_PREVIEW:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("📍 Dashboard WebGIS")
    st.caption("Peta ini selaras dengan hasil K-Means++ dan dibuat ringan agar tidak reload berulang saat digeser.")

    processed_df = st.session_state.get("processed_df")
    geojson_data = load_geojson_data()

    if processed_df is None:
        st.warning("Belum ada data hasil proses. Silakan buka tab Input & Proses dan upload file Dapodik terlebih dahulu.")
    elif geojson_data is None:
        st.error(f"File GeoJSON '{GEOJSON_PATH}' tidak ditemukan. Upload GeoJSON terlebih dahulu atau letakkan file di folder yang sama.")
    else:
        try:
            geojson_name_field = get_geojson_name_field(geojson_data)
            map_geojson = prepare_map_geojson(geojson_data, processed_df, geojson_name_field)

            # Ringkasan dashboard agar tampilan selaras dengan hasil kalkulasi
            total_kec = processed_df["Kecamatan"].nunique()
            total_sekolah = int(processed_df["Jumlah_Sekolah"].sum())
            total_pd = int(processed_df["Jumlah_PD"].sum())
            total_guru = int(processed_df["Jumlah_Guru"].sum())

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Kecamatan", total_kec)
            m2.metric("Sekolah", total_sekolah)
            m3.metric("Peserta Didik", f"{total_pd:,}".replace(",", "."))
            m4.metric("Guru", f"{total_guru:,}".replace(",", "."))

            left, right = st.columns([2, 1])
            with left:
                map_html = build_map_html(map_geojson, processed_df, geojson_name_field)
                components.html(map_html, height=650, scrolling=False)

            with right:
                st.markdown("### Ringkasan Klaster")
                cluster_order = [0, 1, 2, 3]
                cluster_summary = (
                    processed_df.groupby("Klaster")
                    .agg(
                        Kecamatan=("Kecamatan", "count"),
                        Jumlah_Sekolah=("Jumlah_Sekolah", "sum"),
                        Jumlah_PD=("Jumlah_PD", "sum"),
                        Jumlah_Guru=("Jumlah_Guru", "sum"),
                    )
                    .reindex(cluster_order)
                    .fillna(0)
                    .reset_index()
                )
                cluster_names = {
                    0: "Sangat Kritis",
                    1: "Kurang Memadai",
                    2: "Memadai / Aman",
                    3: "Sangat Berlebih",
                }
                for _, row in cluster_summary.iterrows():
                    klaster = int(row["Klaster"])
                    label = cluster_names.get(klaster, "Tidak Diketahui")
                    st.markdown(
                        f"""
                        <div style="padding:12px 14px;border-radius:12px;border:1px solid rgba(0,0,0,0.08);margin-bottom:10px;">
                            <div style="font-weight:700;">Klaster {klaster} - {label}</div>
                            <div style="font-size:13px;color:#666;">{int(row['Kecamatan'])} kecamatan</div>
                            <div style="font-size:13px;color:#666;">Sekolah: {int(row['Jumlah_Sekolah'])} | PD: {int(row['Jumlah_PD'])} | Guru: {int(row['Jumlah_Guru'])}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            st.caption(f"GeoJSON field yang dipakai sebagai pengikat wilayah: {geojson_name_field}")
        except Exception as e:
            st.error(f"⚠️ Terjadi kesalahan saat merender peta: {e}")
    st.markdown('</div>', unsafe_allow_html=True)