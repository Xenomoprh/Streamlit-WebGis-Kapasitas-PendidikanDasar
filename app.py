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
if "geojson_data" not in st.session_state:
    st.session_state["geojson_data"] = None
if "geojson_name" not in st.session_state:
    st.session_state["geojson_name"] = None


def load_geojson_data():
    if "geojson_data" in st.session_state:
        if st.session_state["geojson_data"] is not None:
            return st.session_state["geojson_data"]

    if os.path.exists(GEOJSON_PATH):
        with open(GEOJSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)

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
    return df


def read_xlsx_without_openpyxl(uploaded_file):
    """Read the first sheet of an .xlsx file without loading workbook styles.

    This is a fallback for damaged Excel files whose style definitions break
    openpyxl/pandas parsing. It extracts only cell values from XML.
    """

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

        expected_keywords = ["kecamatan", "nama sekolah", "pd", "guru"]

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


with TAB_INPUT:
    st.subheader("📥 Upload Data dan GeoJSON")
    col_left, col_right = st.columns(2)

    with col_left:
        uploaded_files = st.file_uploader(
            "1) Upload Data Dapodik (CSV/Excel) - bisa banyak file sekaligus",
            type=["csv", "xlsx"],
            accept_multiple_files=True,
        )

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
        if st.session_state["geojson_name"]:
            st.success(f"GeoJSON aktif: {st.session_state['geojson_name']}")
        elif os.path.exists(GEOJSON_PATH):
            st.success(f"GeoJSON aktif: {GEOJSON_PATH}")
        else:
            st.warning("GeoJSON belum diupload. Silakan upload file batas kecamatan.")

        if uploaded_files:
            try:
                df_list = []
                for file in uploaded_files:
                    df_temp = read_dapodik_file(file)
                    df_temp["Sumber_File"] = file.name
                    df_list.append(df_temp)

                raw_df = pd.concat(df_list, ignore_index=True)
                st.session_state["raw_df"] = raw_df

                st.success(f"✅ {len(uploaded_files)} file berhasil diunggah dan digabung!")
                st.caption("Semua file akan diproses menjadi satu tabel gabungan sebelum klasterisasi.")

                with st.spinner("Memproses data dan menjalankan algoritma K-Means++..."):
                    processed_df = process_dapodik_data(raw_df)
                    st.session_state["processed_df"] = processed_df

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
                st.error(f"⚠️ Terjadi kesalahan saat membaca atau memproses file: {e}")
        else:
            st.info("👈 Silakan upload satu atau beberapa file Data Dapodik (Excel/CSV) untuk memulai analisis.")


with TAB_PREVIEW:
    st.subheader("📍 Peta Sebaran Kesenjangan Kapasitas Pendidikan")

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

            m = folium.Map(location=[3.5952, 98.6722], zoom_start=11, tiles="CartoDB positron")

            def style_function(feature):
                klaster = feature["properties"].get("Klaster")
                warna = {
                    0: "#d73027",
                    1: "#fc8d59",
                    2: "#91cf60",
                    3: "#1a9850",
                }.get(klaster, "#9e9e9e")
                return {
                    "fillColor": warna,
                    "color": "#444444",
                    "weight": 1,
                    "fillOpacity": 0.7,
                }

            tooltip = folium.GeoJsonTooltip(
                fields=["Kecamatan", "Jumlah_Sekolah", "Jumlah_PD", "Jumlah_Guru", "Nama_Klaster"],
                aliases=["Kecamatan", "Jumlah Sekolah", "Jumlah PD", "Jumlah Guru", "Status Klaster"],
                localize=True,
                sticky=False,
                labels=True,
            )

            popup = folium.GeoJsonPopup(
                fields=["Kecamatan", "Jumlah_Sekolah", "Jumlah_PD", "Jumlah_Guru", "Nama_Klaster"],
                aliases=["Kecamatan", "Jumlah Sekolah", "Jumlah PD", "Jumlah Guru", "Status Klaster"],
                localize=True,
                labels=True,
                style="background-color: white;",
                max_width=300,
            )

            folium.GeoJson(
                map_geojson,
                name="Klaster Kecamatan",
                style_function=style_function,
                tooltip=tooltip,
                popup=popup,
            ).add_to(m)

            folium.LayerControl().add_to(m)
            st_folium(m, width=900, height=500)

            st.caption(f"GeoJSON field yang dipakai sebagai pengikat wilayah: {geojson_name_field}")
        except Exception as e:
            st.error(f"⚠️ Terjadi kesalahan saat merender peta: {e}")