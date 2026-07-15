import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import json
import os
from copy import deepcopy
from io import BytesIO
import zipfile
import xml.etree.ElementTree as ET
import streamlit.components.v1 as components
import re
from difflib import get_close_matches, SequenceMatcher

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

from sklearn.metrics import davies_bouldin_score

# ==========================================
# KONFIGURASI HALAMAN STREAMLIT
# ==========================================
st.set_page_config(page_title="WebGIS Pendidikan Medan", layout="wide")

st.title("Pemetaan Kesenjangan Kapasitas Pendidikan Dasar Negeri Kota Medan")
st.caption(
    "Aplikasi memproses data Dapodik, mengekstrak sekolah SD dan SMP negeri, lalu menampilkan hasil klaster K-Means++ ke dalam WebGIS yang terhubung ke GeoJSON kecamatan."
)

st.sidebar.title("WebGIS Pendidikan Medan")
st.sidebar.caption("Dashboard klasterisasi kapasitas SD & SMP negeri")
st.sidebar.divider()

nav_choice = st.sidebar.radio(
    "Navigasi",
    ["Ekstrak data sekolah SD dan SMP negeri", "Menampilkan WebGIS", "Evaluasi Model"],
    index=0,
)

st.sidebar.caption("Gunakan menu untuk berpindah antara ekstraksi data bersih dan tampilan peta hasil klaster.")

st.divider()
st.write("21 kecamatan • K-Means++ • GeoJSON linked • SD & SMP negeri")

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
if "extracted_df" not in st.session_state:
    st.session_state["extracted_df"] = None
if "extraction_done" not in st.session_state:
    st.session_state["extraction_done"] = False
if "extraction_summary" not in st.session_state:
    st.session_state["extraction_summary"] = None


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
    for key in ["KECAMATAN", "kecamatan", "Kecamatan", "nm_kecamatan", "NM_KECAMATAN", "NAMA_KEC", "NAMA", "NAME"]:
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


def extract_geojson_names(geojson_data):
    if not geojson_data or "features" not in geojson_data:
        return []

    field_name = get_geojson_name_field(geojson_data)
    names = []
    for feature in geojson_data.get("features", []):
        props = feature.get("properties", {})
        value = props.get(field_name)
        if value:
            names.append(str(value).strip())
    return names


def build_extraction_summary(df):
    total_rows = len(df)
    sd_count = 0
    smp_count = 0

    if "BP" in df.columns:
        bp_values = df["BP"].astype(str).str.strip().str.upper()
        sd_count = int(bp_values.eq("SD").sum())
        smp_count = int(bp_values.eq("SMP").sum())

    return {
        "total_rows": int(total_rows),
        "sd_count": sd_count,
        "smp_count": smp_count,
    }


def build_cluster_report_df(df_agg):
    report_columns = [
        "Kecamatan",
        "Jumlah_Sekolah",
        "Jumlah_PD",
        "Jumlah_Guru",
        "Rasio_PD_Sekolah",
        "Rasio_PD_Guru",
        "Klaster",
        "Nama_Klaster",
        "Ranking",
        "Prioritas_Intervensi",
    ]

    extra_columns = [
        "Rata_Rata_Sekolah_Per_Kec",
        "Rata_Rata_PD_Per_Kec",
        "Rata_Rata_Guru_Per_Kec",
    ]

    available_columns = [col for col in report_columns + extra_columns if col in df_agg.columns]
    report_df = df_agg[available_columns].copy()

    if "Klaster" in report_df.columns:
        report_df["Klaster"] = report_df["Klaster"].astype("Int64")

    return report_df.sort_values("Kecamatan").reset_index(drop=True)


def create_excel_download_bytes(df, sheet_name="Klaster"):
    buffer = BytesIO()
    writer_engines = ["openpyxl", "xlsxwriter"]
    last_error = None

    for engine in writer_engines:
        try:
            with pd.ExcelWriter(buffer, engine=engine) as writer:
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            buffer.seek(0)
            return buffer.getvalue()
        except Exception as exc:
            last_error = exc
            buffer.seek(0)
            buffer.truncate(0)

    raise RuntimeError(f"Excel export gagal: {last_error}")


def evaluate_kmeans_candidates(df_agg, chosen_k=4, max_k=8):
    feature_columns = ["Jumlah_Sekolah", "Rasio_PD_Sekolah", "Rasio_PD_Guru", "Jumlah_PD"]
    features = df_agg[feature_columns].dropna().copy()

    if len(features) < 3:
        raise ValueError("Data terlalu sedikit untuk evaluasi klaster.")

    max_k = min(max_k, len(features))
    if max_k < 2:
        raise ValueError("Minimal diperlukan 2 data untuk evaluasi K-Means.")

    scaler = StandardScaler()
    features_scaled = scaler.fit_transform(features)

    rows = []
    for k in range(2, max_k + 1):
        model = KMeans(n_clusters=k, init="k-means++", random_state=42, n_init=10)
        labels = model.fit_predict(features_scaled)
        
        # Menghitung Davies-Bouldin Index (DBI)
        dbi = davies_bouldin_score(features_scaled, labels) if len(set(labels)) > 1 else None

        rows.append(
            {
                "k": k,
                "inertia": float(model.inertia_),
                "dbi": float(dbi) if dbi is not None else None,
            }
        )

    evaluation_df = pd.DataFrame(rows)
    
    # Ambil nilai WCSS dan DBI secara spesifik untuk K yang dipilih via Elbow Method
    chosen_row = evaluation_df.loc[evaluation_df["k"] == chosen_k].iloc[0]

    # Kembalikan dataframe evaluasi, K terpilih, WCSS, dan DBI
    return evaluation_df, chosen_k, float(chosen_row["inertia"]), float(chosen_row["dbi"])


def infer_kecamatan_from_filename(file_name, candidate_names=None):
    """Ambil nama kecamatan dari nama file secara otomatis."""
    base_name = os.path.splitext(file_name)[0]

    if candidate_names:
        matched_name = best_name_match(base_name, candidate_names)
        if matched_name:
            return matched_name

    patterns = [
        r"kec(?:amatan)?\.?\s*([a-zA-Z0-9\s\.\-]+?)(?:\s*-\s*|$)",
        r"data\s+sekolah\s+(.+?)(?:\s*-\s*|$)",
        r"(.+?)(?:\s*-\s*dapodik|\s*-\s*data|$)",
    ]

    for pattern in patterns:
        match = re.search(pattern, base_name, flags=re.IGNORECASE)
        if match:
            extracted = match.group(1).strip()
            extracted_match = best_name_match(extracted, candidate_names or [])
            if extracted_match:
                return extracted_match
            if extracted:
                return extracted

    return None

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

    kecamatan_col = infer_column({"kecamatan", "kec"}, ["kecamatan", "kec ", "kec.", "wilayah", "district"])
    sekolah_col = infer_column({"nama sekolah", "sekolah"}, ["nama sekolah", "sekolah", "school"])
    pd_col = infer_column({"pd", "peserta didik", "siswa"}, ["pd", "peserta didik", "jumlah siswa", "siswa"])
    guru_col = infer_column({"guru"}, ["guru", "pendidik"])
    pegawai_col = infer_column({"pegawai"}, ["pegawai", "tenaga kependidikan", "tendik"])
    rombel_col = infer_column({"rombel"}, ["rombel", "jumlah rombel"])
    rkelas_col = infer_column({"r.kelas", "r kelas"}, ["r.kelas", "r kelas", "ruang kelas", "kelas"])
    rlab_col = infer_column({"r.lab", "r lab"}, ["r.lab", "r lab", "laboratorium", "lab"])
    rperpus_col = infer_column({"r.perpus", "r perpus"}, ["r.perpus", "r perpus", "perpus", "perpustakaan"])
    
    # Menangkap kolom BP (Bentuk Pendidikan) dan Status
    bp_col = infer_column({"bp", "bentuk pendidikan"}, ["bp", "bentuk pendidikan"])
    status_col = infer_column({"status", "status sekolah"}, ["status"])

    if kecamatan_col is not None: rename_map[kecamatan_col] = "Kecamatan"
    if sekolah_col is not None: rename_map[sekolah_col] = "Nama Sekolah"
    if pd_col is not None: rename_map[pd_col] = "PD"
    if guru_col is not None: rename_map[guru_col] = "Guru"
    if pegawai_col is not None: rename_map[pegawai_col] = "Pegawai"
    if rombel_col is not None: rename_map[rombel_col] = "Rombel"
    if rkelas_col is not None: rename_map[rkelas_col] = "R.Kelas"
    if rlab_col is not None: rename_map[rlab_col] = "R.Lab"
    if rperpus_col is not None: rename_map[rperpus_col] = "R.Perpus"
    if bp_col is not None: rename_map[bp_col] = "BP"
    if status_col is not None: rename_map[status_col] = "Status"

    df = df.rename(columns=rename_map)

    # Bersihkan nama kolom supaya konsisten
    df.columns = [str(col).replace("\ufeff", "").strip() for col in df.columns]
    
    # HAPUS KOLOM DUPLIKAT YANG MEMBUAT PANDAS CRASH
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

        expected_keywords = ["npsn", "nama sekolah", "pd", "guru", "bp", "status"]

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

def best_name_match(source_name, candidate_names):
    """Cari kecamatan kandidat terbaik berdasarkan nama sumber."""
    if not source_name or not candidate_names:
        return None

    source_norm = normalize_name(source_name)
    candidate_map = {normalize_name(name): name for name in candidate_names}

    if source_norm in candidate_map:
        return candidate_map[source_norm]

    close = get_close_matches(source_norm, list(candidate_map.keys()), n=1, cutoff=0.72)
    if close:
        return candidate_map[close[0]]

    best_candidate = None
    best_score = 0.0
    for candidate_norm, candidate_name in candidate_map.items():
        score = SequenceMatcher(None, source_norm, candidate_norm).ratio()
        if score > best_score:
            best_score = score
            best_candidate = candidate_name

    return best_candidate if best_score >= 0.72 else None


def assign_dynamic_labels(df_agg):
    """
    Fungsi Post-Hoc Cluster Interpretation (Lexicographical Ordering)
    Mengurutkan klaster berdasarkan Rasio PD/Sekolah sebagai indikator utama.
    Jika seri/berdekatan, Rasio PD/Guru menjadi pembanding sekunder.
    Mengembalikan Label, Deskripsi, Warna, Ranking, dan Prioritas Intervensi.
    """
    # 1. Hitung rata-rata parameter per klaster
    profile = df_agg.groupby("Klaster").agg(
        Avg_Rasio_Sekolah=("Rasio_PD_Sekolah", "mean"),
        Avg_Rasio_Guru=("Rasio_PD_Guru", "mean")
    ).reset_index()

    # 2. Pengurutan Bertingkat (Lexicographical Ordering)
    # Urutkan dari rasio terendah (relatif aman) ke rasio tertinggi (kritis)
    profile = profile.sort_values(
        by=["Avg_Rasio_Sekolah", "Avg_Rasio_Guru"], 
        ascending=[True, True]
    ).reset_index(drop=True)

    dynamic_names = {}
    dynamic_descriptions = {}
    dynamic_colors = {}
    dynamic_ranks = {}
    dynamic_priorities = {}

    # Kumpulan Label Akademis
    labels = ["Sangat Memadai", "Memadai", "Kurang Memadai", "Sangat Kritis"]
    
    # Warna terikat mutlak pada ranking (Hijau Tua -> Merah Tua)
    colors = ["#1a9641", "#a6d96a", "#fdae61", "#d7191c"] 
    
    # Deskripsi murni berdasarkan indikator penelitian (tanpa asumsi)
    desc = [
        "Klaster ini memiliki karakteristik dengan tingkat tekanan kapasitas layanan pendidikan yang relatif paling rendah berdasarkan hasil klasterisasi.",

        "Klaster ini memiliki karakteristik dengan tingkat tekanan kapasitas layanan pendidikan yang relatif rendah berdasarkan hasil klasterisasi.",

        "Klaster ini memiliki karakteristik dengan tingkat tekanan kapasitas layanan pendidikan yang relatif tinggi berdasarkan hasil klasterisasi.",

        "Klaster ini memiliki karakteristik dengan tingkat tekanan kapasitas layanan pendidikan yang relatif paling tinggi berdasarkan hasil klasterisasi."
    ]

    # 3. Pemasangan dinamis ke nomor Klaster asli
    for idx in range(len(profile)):
        c_id = profile.loc[idx, 'Klaster']
        if idx < 4:
            dynamic_names[c_id] = f"{labels[idx]}"
            dynamic_descriptions[c_id] = desc[idx]
            dynamic_colors[c_id] = colors[idx]
            
            # Ranking Kapasitas (1 = Sangat Memadai s/d 4 = Sangat Kritis)
            dynamic_ranks[c_id] = idx + 1               
            
            # Prioritas Intervensi (4 = Aman s/d 1 = Prioritas Utama/Sangat Kritis)
            dynamic_priorities[c_id] = 4 - idx

    return dynamic_names, dynamic_descriptions, dynamic_colors, dynamic_ranks, dynamic_priorities


def process_dapodik_data(df, geojson_names=None):
    if "Nama Sekolah" in df.columns:
        df = df[~df["Nama Sekolah"].astype(str).str.lower().str.contains("total", na=False)]
    if "NPSN" in df.columns:
        df = df[df["NPSN"].notna()]
        
    # =========================================================================
    # FILTER HANYA SD & SMP NEGERI
    # =========================================================================
    if "BP" in df.columns:
        df = df[df["BP"].astype(str).str.strip().str.upper().isin(["SD", "SMP"])]
    if "Status" in df.columns:
        df = df[df["Status"].astype(str).str.strip().str.upper() == "NEGERI"]
        
    required_columns = ["Kecamatan", "PD", "Guru"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Kolom wajib tidak ditemukan: {missing_columns}")

    df = df.dropna(subset=["Kecamatan", "PD", "Guru"])
    df["PD"] = pd.to_numeric(df["PD"], errors="coerce")
    df["Guru"] = pd.to_numeric(df["Guru"], errors="coerce")
    df = df.dropna(subset=["Kecamatan", "PD", "Guru"])

    # Kunci normalisasi dipakai untuk memastikan hasil klaster nyambung ke GeoJSON
    df["Kec_Mapping"] = df["Kecamatan"].apply(normalize_name)

    # AGREGASI DATA
    df_agg = df.groupby("Kec_Mapping", as_index=False).agg(
        Kecamatan=("Kecamatan", "first"),
        Jumlah_Sekolah=("Nama Sekolah", "count"),
        Jumlah_PD=("PD", "sum"),
        Jumlah_Guru=("Guru", "sum"),
    )

    df_agg["Rasio_PD_Sekolah"] = df_agg["Jumlah_PD"] / df_agg["Jumlah_Sekolah"]
    df_agg["Rasio_PD_Guru"] = df_agg["Jumlah_PD"] / df_agg["Jumlah_Guru"]

    X = df_agg[["Jumlah_Sekolah", "Rasio_PD_Sekolah", "Rasio_PD_Guru", "Jumlah_PD"]]
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    kmeans = KMeans(n_clusters=4, init="k-means++", random_state=42, n_init=10)
    df_agg["Klaster"] = kmeans.fit_predict(X_scaled)

    dynamic_names, dynamic_descriptions, dynamic_colors, dynamic_ranks, dynamic_priorities = assign_dynamic_labels(df_agg)
    df_agg["Nama_Klaster"] = df_agg["Klaster"].map(dynamic_names)
    df_agg["Deskripsi_Klaster"] = df_agg["Klaster"].map(dynamic_descriptions)
    df_agg["Cluster_Color"] = df_agg["Klaster"].map(dynamic_colors)
    df_agg["Ranking"] = df_agg["Klaster"].map(dynamic_ranks)
    df_agg["Prioritas_Intervensi"] = df_agg["Klaster"].map(dynamic_priorities)
    
    # PERHITUNGAN RATA-RATA UNTUK PROFIL
    df_agg["Rata_Rata_Sekolah_Per_Kec"] = df_agg.groupby("Klaster")["Jumlah_Sekolah"].transform("mean")
    df_agg["Rata_Rata_PD_Per_Kec"] = df_agg.groupby("Klaster")["Jumlah_PD"].transform("mean")
    df_agg["Rata_Rata_Guru_Per_Kec"] = df_agg.groupby("Klaster")["Jumlah_Guru"].transform("mean")
    
    return df_agg


def profile_clusters(df_agg):
    """Hitung profil rata-rata indikator untuk setiap klaster.
    
    Returns: DataFrame dengan statistik klaster (mean dan count)
    """
    cluster_profile = df_agg.groupby("Klaster").agg(
        Jumlah_Kecamatan=("Kecamatan", "count"),
        Rata_Rata_Sekolah=("Jumlah_Sekolah", "mean"),
        Rata_Rata_PD=("Jumlah_PD", "mean"),
        Rata_Rata_Guru=("Jumlah_Guru", "mean"),
        Rata_Rasio_PD_Sekolah=("Rasio_PD_Sekolah", "mean"),
        Rata_Rasio_PD_Guru=("Rasio_PD_Guru", "mean"),
    ).reset_index()
    
    return cluster_profile


def normalize_name(text):
    """Normalisasi nama kecamatan menggunakan regex untuk pencocokan akurat.
    
    Menghapus spasi, tanda baca, kata 'kecamatan', 'kec', 'medan', 
    dan mengubah menjadi lowercase alphanumeric murni.
    """
    if not text:
        return ""
    tokens = re.split(r"[^a-z0-9]+", str(text).lower())
    stopwords = {"kecamatan", "kec", "medan"}
    cleaned_tokens = [token for token in tokens if token and token not in stopwords]
    return "".join(cleaned_tokens)


def prepare_map_geojson(geojson_data, df_agg, geojson_name_field):
    geojson_map = deepcopy(geojson_data)

    lookup = df_agg.set_index("Kec_Mapping").to_dict(orient="index")

    for feature in geojson_map.get("features", []):
        props = feature.setdefault("properties", {})

        # GeoJSON dan data Dapodik sama-sama dinormalisasi agar penggabungan stabil
        kec_raw = normalize_name(props.get(geojson_name_field, ""))

        if kec_raw in lookup:
            row = lookup[kec_raw]
            row_copy = row.copy()
            if "Klaster" in row_copy:
                row_copy["Klaster"] = int(row_copy["Klaster"]) if pd.notna(row_copy["Klaster"]) else None
            row_copy["cluster_color"] = row_copy.get("Cluster_Color", "#cccccc")
            row_copy["match_mode"] = "normalisasi_nama"
            props.update(row_copy)
        else:
            props["Jumlah_Sekolah"] = props["Jumlah_PD"] = props["Jumlah_Guru"] = None
            props["Rasio_PD_Sekolah"] = props["Rasio_PD_Guru"] = None
            props["Klaster"] = None
            props["Nama_Klaster"] = "Data tidak tersedia"
            props["Deskripsi_Klaster"] = ""
            props["Rata_Rata_Sekolah_Per_Kec"] = props["Rata_Rata_PD_Per_Kec"] = None
            props["Rata_Rata_Guru_Per_Kec"] = None
            props["match_mode"] = "tidak cocok"
            props["cluster_color"] = "#cccccc"

    return geojson_map


def build_map_html(geojson_data, df_agg, geojson_name_field):
    m = folium.Map(location=[3.5952, 98.6722], zoom_start=11, tiles="CartoDB positron", prefer_canvas=True)

    def style_function(feature):
        props = feature["properties"]
        cluster_color = props.get("cluster_color")
        if cluster_color is None:
            cluster_color = "#cccccc"
        return {
            "fillColor": cluster_color,
            "color": "#444444",
            "weight": 1.5,
            "fillOpacity": 0.8,
        }

    # Format tooltip dan popup tanpa menampilkan NaN
    tooltip = folium.GeoJsonTooltip(
        fields=[geojson_name_field, "Jumlah_Sekolah", "Jumlah_PD", "Jumlah_Guru", "Rasio_PD_Sekolah", "Rasio_PD_Guru", "Rata_Rata_Sekolah_Per_Kec", "Rata_Rata_PD_Per_Kec", "Nama_Klaster"],
        aliases=["Kecamatan", "Jumlah Sekolah", "Jumlah PD", "Jumlah Guru", "Rasio PD/Sekolah", "Rasio PD/Guru", "Rata2 Sekolah/Klaster", "Rata2 PD/Klaster", "Status Klaster"],
        localize=True,
        sticky=False,
        labels=True,
    )

    popup = folium.GeoJsonPopup(
        fields=[geojson_name_field, "Jumlah_Sekolah", "Jumlah_PD", "Jumlah_Guru", "Rasio_PD_Sekolah", "Rasio_PD_Guru", "Rata_Rata_Sekolah_Per_Kec", "Rata_Rata_PD_Per_Kec", "Nama_Klaster", "match_mode"],
        aliases=["Kecamatan", "Jumlah Sekolah", "Jumlah PD", "Jumlah Guru", "Rasio PD/Sekolah", "Rasio PD/Guru", "Rata2 Sekolah/Klaster", "Rata2 PD/Klaster", "Status Klaster", "Mode Match"],
        localize=True,
        labels=True,
        style="background-color: white; border-radius: 8px; padding: 10px;",
        max_width=360,
    )

    folium.GeoJson(
        geojson_data,
        name="Klaster Kecamatan",
        style_function=style_function,
        tooltip=tooltip,
        popup=popup,
    ).add_to(m)

    legend_profile = df_agg.groupby("Klaster").agg(
        Avg_Rasio_Sekolah=("Rasio_PD_Sekolah", "mean"),
        Avg_Rasio_Guru=("Rasio_PD_Guru", "mean"),
        Nama_Klaster=("Nama_Klaster", "first"),
        Cluster_Color=("Cluster_Color", "first"),
    ).reset_index().sort_values(
        by=["Avg_Rasio_Sekolah", "Avg_Rasio_Guru"],
        ascending=[True, True],
    )

    legend_items = "".join(
        [
            f"<div><span style='display:inline-block;width:12px;height:12px;background:{row['Cluster_Color']};margin-right:8px;border-radius:3px'></span>{row['Nama_Klaster']}</div>"
            for _, row in legend_profile.iterrows()
        ]
    )

    legend_html = f"""
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
        {legend_items}
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))

    folium.LayerControl(position="topright").add_to(m)
    return m.get_root().render()


if nav_choice == "Ekstrak data sekolah SD dan SMP negeri":
    st.subheader("Ekstrak Data Dapodik")
    st.caption("Upload file Dapodik dan GeoJSON, lalu jalankan ekstraksi secara manual. Tidak ada penyaringan otomatis saat file dipilih.")
    col_left, col_right = st.columns([1.2, 0.8])

    with col_left:
        uploaded_files = st.file_uploader(
            "1) Upload Data Dapodik (CSV/Excel) - bisa banyak file sekaligus",
            type=["csv", "xlsx"],
            accept_multiple_files=True,
        )

        if uploaded_files:
            st.caption("File yang dipilih akan disimpan dulu. Nama kecamatan akan dikenali otomatis dari nama file atau kolom yang tersedia.")
            st.markdown("**File aktif**")
            st.write([f.name for f in uploaded_files])

            if st.button("Simpan file Dapodik", use_container_width=True):
                try:
                    df_list = []
                    file_names = []
                    geojson_for_mapping = load_geojson_data()
                    geojson_names_for_mapping = extract_geojson_names(geojson_for_mapping)

                    for file in uploaded_files:
                        df_temp = read_dapodik_file(file)
                        df_temp["Sumber_File"] = file.name

                        inferred_kecamatan = infer_kecamatan_from_filename(file.name, geojson_names_for_mapping)
                        if inferred_kecamatan:
                            if "Kecamatan" not in df_temp.columns:
                                df_temp["Kecamatan"] = inferred_kecamatan
                            else:
                                df_temp["Kecamatan"] = df_temp["Kecamatan"].fillna(inferred_kecamatan)

                        df_list.append(df_temp)
                        file_names.append(file.name)

                    raw_df = pd.concat(df_list, ignore_index=True)
                    st.session_state["raw_df"] = raw_df
                    st.session_state["raw_file_names"] = file_names
                    st.session_state["processed_df"] = None
                    st.session_state["extracted_df"] = None
                    st.session_state["extraction_done"] = False
                    st.session_state["extraction_summary"] = None

                    st.success(f"{len(uploaded_files)} file berhasil disimpan. Data lama sudah diganti.")
                    st.info("File tersimpan. Klik tombol ekstraksi untuk membuat data bersih SD & SMP negeri.")
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat menyimpan file: {e}")

    with col_right:
        uploaded_geojson = st.file_uploader(
            "2) Upload GeoJSON Batas Kecamatan",
            type=["geojson", "json"],
        )

        if uploaded_geojson is not None:
            if st.button("Gunakan / simpan GeoJSON yang diupload"):
                try:
                    geojson_obj = update_geojson_from_upload(uploaded_geojson)
                    st.success(f"GeoJSON berhasil diperbarui dari {uploaded_geojson.name}")
                    st.caption(f"Aktif: {uploaded_geojson.name}")
                    st.session_state["geojson_data"] = geojson_obj
                except Exception as e:
                    st.error(f"GeoJSON gagal diperbarui: {e}")

    st.info("Upload satu atau beberapa file Dapodik sekaligus. Hasilnya akan digabung lalu disiapkan untuk ekstraksi dan proses klaster.")

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

            st.subheader("Hasil Ekstraksi Data (SD & SMP Negeri)")
            st.caption("Tekan tombol di bawah ini untuk menjalankan penyaringan data bersih secara manual. Hasil yang ditampilkan dan diunduh difokuskan ke kolom penting untuk algoritma.")

            if st.button("Ekstrak Data SD & SMP Negeri", type="primary", use_container_width=True):
                extracted_df = st.session_state["raw_df"].copy()
                if "BP" in extracted_df.columns:
                    extracted_df = extracted_df[extracted_df["BP"].astype(str).str.strip().str.upper().isin(["SD", "SMP"])]
                if "Status" in extracted_df.columns:
                    extracted_df = extracted_df[extracted_df["Status"].astype(str).str.strip().str.upper() == "NEGERI"]

                extracted_df = extracted_df.dropna(subset=["Kecamatan", "PD", "Guru"])

                if "Kecamatan" in extracted_df.columns:
                    extracted_df["Kecamatan"] = extracted_df["Kecamatan"].astype(str).str.replace("nan", "", regex=False).str.strip()

                st.session_state["extracted_df"] = extracted_df
                st.session_state["extraction_done"] = True

            if st.session_state["extraction_done"] and st.session_state["extracted_df"] is not None:
                extracted_df = st.session_state["extracted_df"]
                extraction_summary = st.session_state.get("extraction_summary") or build_extraction_summary(extracted_df)
                st.session_state["extraction_summary"] = extraction_summary

                summary_m1, summary_m2, summary_m3 = st.columns(3)
                summary_m1.metric("Total hasil ekstraksi", extraction_summary["total_rows"])
                summary_m2.metric("SD berhasil diekstrak", extraction_summary["sd_count"])
                summary_m3.metric("SMP berhasil diekstrak", extraction_summary["smp_count"])

                preferred_columns = ["Kecamatan", "Nama Sekolah", "BP", "Status", "PD", "Guru", "Pegawai", "Rombel", "R.Kelas", "R.Lab", "R.Perpus"]
                preview_columns = [col for col in preferred_columns if col in extracted_df.columns]
                cleaned_export_df = extracted_df[preview_columns].copy()

                st.dataframe(cleaned_export_df, height=220, use_container_width=True, hide_index=True)

                csv_bersih = cleaned_export_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="⬇Unduh Data Bersih SD & SMP Negeri (CSV)",
                    data=csv_bersih,
                    file_name="Data_SD_SMP_Negeri_Medan_Clean.csv",
                    mime="text/csv",
                    use_container_width=True,
                    help="Gunakan file ini sebagai lampiran data skripsi Anda.",
                )
            else:
                st.info("Belum ada data bersih yang diekstrak. Klik tombol ekstraksi untuk menampilkan preview dan unduhan.")
            st.divider()

        if st.button("Proses Data Dapodik", type="primary", disabled=st.session_state["raw_df"] is None):
            try:
                with st.spinner("Memproses data dan menjalankan algoritma K-Means++..."):
                    geojson_for_matching = load_geojson_data()
                    geojson_names = []
                    if geojson_for_matching and "features" in geojson_for_matching:
                        geojson_names = extract_geojson_names(geojson_for_matching)

                    source_df = st.session_state["extracted_df"] if st.session_state["extraction_done"] and st.session_state["extracted_df"] is not None else st.session_state["raw_df"]
                    processed_df = process_dapodik_data(source_df, geojson_names=geojson_names)
                    st.session_state["processed_df"] = processed_df

                st.success("Data berhasil diproses.")
                
                # Tampilkan tabel ringkasan klaster (profil rata-rata)
                st.subheader("Profil Rata-rata Indikator per Klaster")
                cluster_profile = profile_clusters(processed_df)
                cluster_names_dict = processed_df.groupby("Klaster")["Nama_Klaster"].first().to_dict()
                cluster_profile["Nama_Klaster"] = cluster_profile["Klaster"].map(cluster_names_dict)
                display_profile = cluster_profile[["Klaster", "Nama_Klaster", "Jumlah_Kecamatan", "Rata_Rata_Sekolah", "Rata_Rata_PD", "Rata_Rata_Guru", "Rata_Rasio_PD_Sekolah", "Rata_Rasio_PD_Guru"]].copy()
                display_profile.columns = ["Klaster", "Nama Klaster", "Jml Kecamatan", "Rata2 Sekolah", "Rata2 PD", "Rata2 Guru", "Rata2 Rasio PD/Sekolah", "Rata2 Rasio PD/Guru"]
                st.dataframe(display_profile, use_container_width=True, hide_index=True)
                
                st.subheader("Hasil Klasterisasi per Kecamatan")
                st.dataframe(
                    processed_df[["Kecamatan", "Jumlah_Sekolah", "Jumlah_PD", "Jumlah_Guru", "Rasio_PD_Sekolah", "Rasio_PD_Guru", "Nama_Klaster"]],
                    use_container_width=True,
                    hide_index=True,
                )

                csv_bytes = processed_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="Unduh hasil klasterisasi (CSV)",
                    data=csv_bytes,
                    file_name="hasil_klasterisasi_kecamatan_medan.csv",
                    mime="text/csv",
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file: {e}")

        if st.session_state["processed_df"] is None:
            st.info("Upload lalu simpan file Dapodik terlebih dahulu. Jika ingin data bersih, tekan tombol ekstraksi sebelum proses klaster.")
elif nav_choice == "Menampilkan WebGIS":
    st.subheader("Dashboard WebGIS")
    st.caption("Peta ini selaras dengan hasil K-Means++ dan disusun agar tampil seperti dashboard analitis.")

    processed_df = st.session_state.get("processed_df")
    geojson_data = load_geojson_data()

    if processed_df is None:
        st.warning("Belum ada data hasil proses. Silakan kembali ke menu ekstraksi dan jalankan proses data terlebih dahulu.")
    elif geojson_data is None:
        st.error(f"File GeoJSON '{GEOJSON_PATH}' tidak ditemukan. Upload GeoJSON terlebih dahulu atau letakkan file di folder yang sama.")
    else:
        try:
            geojson_name_field = get_geojson_name_field(geojson_data)
            map_geojson = prepare_map_geojson(geojson_data, processed_df, geojson_name_field)
            extraction_summary = st.session_state.get("extraction_summary")

            matched = 0
            missing = 0
            for feature in map_geojson.get("features", []):
                props = feature.get("properties", {})
                if props.get("cluster_color") == "#cccccc" or props.get("Nama_Klaster") == "Data tidak tersedia":
                    missing += 1
                else:
                    matched += 1

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
            st.caption(f"Match GeoJSON ke data klaster: {matched} cocok, {missing} belum cocok")

            if extraction_summary is not None:
                ex1, ex2, ex3 = st.columns(3)
                ex1.metric("Total hasil ekstraksi", extraction_summary.get("total_rows", 0))
                ex2.metric("SD berhasil diekstrak", extraction_summary.get("sd_count", 0))
                ex3.metric("SMP berhasil diekstrak", extraction_summary.get("smp_count", 0))

            st.subheader("Hasil Proses yang Disimpan")
            st.caption("Ringkasan ini tetap ditampilkan saat Anda berpindah ke tab WebGIS karena datanya diambil dari session state.")
            result_preview = build_cluster_report_df(processed_df)
            st.dataframe(result_preview, use_container_width=True, hide_index=True, height=280)

            export_col1, export_col2 = st.columns(2)
            csv_export = result_preview.to_csv(index=False).encode("utf-8")
            export_col1.download_button(
                label="Unduh Data Klastering (CSV)",
                data=csv_export,
                file_name="data_klastering_kecamatan_medan.csv",
                mime="text/csv",
                use_container_width=True,
            )

            try:
                excel_export = create_excel_download_bytes(result_preview, sheet_name="Klaster Medan")
                export_col2.download_button(
                    label="Unduh Data Klastering (Excel)",
                    data=excel_export,
                    file_name="data_klastering_kecamatan_medan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as excel_error:
                export_col2.button("Unduh Data Klastering (Excel)", disabled=True, use_container_width=True)
                export_col2.caption(f"Export Excel belum tersedia: {excel_error}")

            left, right = st.columns([2, 1])
            with left:
                map_html = build_map_html(map_geojson, processed_df, geojson_name_field)
                components.html(map_html, height=650, scrolling=False)

            with right:
                st.subheader("Profil Klaster (Rata-rata Indikator)")
                cluster_profile = profile_clusters(processed_df)
                cluster_names = processed_df.groupby("Klaster")["Nama_Klaster"].first().to_dict()
                cluster_meanings = processed_df.groupby("Klaster")["Deskripsi_Klaster"].first().to_dict()
                
                for _, row in cluster_profile.iterrows():
                    klaster = int(row["Klaster"])
                    label = cluster_names.get(klaster, "Tidak Diketahui")
                    meaning = cluster_meanings.get(klaster, "")
                    n_kec = int(row["Jumlah_Kecamatan"])
                    avg_sekolah = f"{row['Rata_Rata_Sekolah']:.1f}"
                    avg_pd = f"{row['Rata_Rata_PD']:.0f}"
                    avg_guru = f"{row['Rata_Rata_Guru']:.0f}"
                    avg_rasio_pd_sekolah = f"{row['Rata_Rasio_PD_Sekolah']:.1f}"
                    avg_rasio_pd_guru = f"{row['Rata_Rasio_PD_Guru']:.1f}"

                    with st.container():
                        st.markdown(
                            f"<div style='margin-bottom:0.35rem;'><strong>{label}</strong></div>"
                            f"<div style='font-size:0.88rem; line-height:1.35; margin-bottom:0.45rem; color:rgba(255,255,255,0.78);'>{meaning}</div>",
                            unsafe_allow_html=True,
                        )
                        st.caption(
                            f"Kecamatan: {n_kec} | Sekolah: {avg_sekolah} | PD: {avg_pd} | Guru: {avg_guru} | Rasio PD/Sekolah: {avg_rasio_pd_sekolah} | Rasio PD/Guru: {avg_rasio_pd_guru}"
                        )
                        st.markdown("<div style='margin:0.35rem 0 0.2rem 0;'></div>", unsafe_allow_html=True)

            st.caption(f"GeoJSON field yang dipakai sebagai pengikat wilayah: {geojson_name_field}")
        except Exception as e:
            st.error(f"Terjadi kesalahan saat merender peta: {e}")
elif nav_choice == "Evaluasi Model":
    st.subheader("Evaluasi Model K-Means++")
    st.caption("Halaman ini menunjukkan bukti numerik penentuan jumlah klaster menggunakan Elbow Method (WCSS/Inertia), divalidasi menggunakan Davies-Bouldin Index (DBI).")

    processed_df = st.session_state.get("processed_df")

    if processed_df is None:
        st.warning("Belum ada data hasil proses. Jalankan proses klaster terlebih dahulu sebelum membuka evaluasi model.")
    else:
        try:
            # Memanggil fungsi evaluasi dengan menetapkan K=4 sesuai keputusan Elbow
            evaluation_df, current_k, wcss_score, dbi_score = evaluate_kmeans_candidates(processed_df, chosen_k=4)

            best_message = (
                f"Berdasarkan analisis visual titik patahan (Elbow Method) pada grafik penurunan WCSS, "
                f"jumlah klaster optimal yang ditetapkan untuk penelitian ini adalah **K = {current_k}**. "
                f"Davies-Bouldin Index (DBI) digunakan sebagai pengujian validitas klaster, dengan skor sebesar **{dbi_score:.3f}** "
                f"dan nilai WCSS sebesar **{wcss_score:.3f}**."
            )
            st.info(best_message)

            # Menampilkan 3 Metrik Jejer
            metric_col1, metric_col2, metric_col3 = st.columns(3)
            metric_col1.metric("K Terpilih (Elbow Method)", current_k)
            metric_col2.metric("Nilai WCSS (Inertia)", f"{wcss_score:.3f}")
            metric_col3.metric("Skor Validasi DBI", f"{dbi_score:.3f}")

            # Menampilkan Grafik
            chart_left, chart_right = st.columns(2)
            with chart_left:
                st.subheader("Elbow Method (Inertia)")
                st.line_chart(evaluation_df.set_index("k")["inertia"])
            with chart_right:
                st.subheader("Davies-Bouldin Index (DBI)")
                st.line_chart(evaluation_df.set_index("k")["dbi"])

            # Tabel Rincian Data Evaluasi
            st.subheader("Tabel Evaluasi K")
            display_eval = evaluation_df.copy()
            display_eval["inertia"] = display_eval["inertia"].map(lambda x: f"{x:.3f}")
            display_eval["dbi"] = display_eval["dbi"].map(lambda x: f"{x:.3f}" if pd.notna(x) else "-")
            display_eval.columns = ["k", "WCSS (Inertia)", "Davies-Bouldin Index"]
            st.dataframe(display_eval, use_container_width=True, hide_index=True)

            st.caption("Elbow Method (Inertia) mengukur jarak titik data ke pusat klasternya (semakin turun melandai berarti optimal). DBI mengukur seberapa baik pemisahan antar klaster (nilai yang rendah menunjukkan validitas pembagian yang baik).")
        except Exception as e:
            st.error(f"Evaluasi model gagal dijalankan: {e}")