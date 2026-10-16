import streamlit as st
import zipfile
import io
import xml.etree.ElementTree as ET
import pandas as pd
from shapely.geometry import Point, Polygon

# === KONFIGURASI HALAMAN ===
st.set_page_config(page_title="FTTH KML/KMZ to Excel", layout="wide")

st.title("🗺️ Konversi Data KML/KMZ ke Format Excel FTTH (Kolom A - BD)")
st.write(
    "Aplikasi ini mengekstrak folder **HP**, **FAT**, **Pole**, dan **Boundary** dari file KML/KMZ, "
    "mencocokkan tiang (Pole) yang memiliki titik koordinat sama/terdekat dengan FAT, "
    "serta mengurutkan nomor HP dari 1 pada setiap FAT ke dalam kolom **A sampai BD**."
)

# === FUNGSI BANTU NAMA KOLOM EXCEL ===
def get_excel_col_name(col_idx):
    """Mengubah indeks angka (0-55) menjadi huruf kolom Excel (A-BD)."""
    result = ""
    while col_idx >= 0:
        result = chr(col_idx % 26 + 65) + result
        col_idx = col_idx // 26 - 1
    return result

def excel_col_to_idx(col_str):
    """Mengubah huruf kolom Excel (misal 'G', 'AJ', 'BD') menjadi indeks 0-based."""
    idx = 0
    for char in col_str.upper():
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx - 1

# === FUNGSI PARSER KML ===
def parse_kml_data(file_content):
    polygon_list = []
    hp_list = []
    fat_list = []
    pole_list = []

    tree = ET.parse(io.BytesIO(file_content))
    root = tree.getroot()
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}

    def parse_coords(coord_str):
        coords = []
        for tuple_str in coord_str.strip().split():
            parts = tuple_str.split(',')
            if len(parts) >= 2:
                coords.append((float(parts[0]), float(parts[1])))
        return coords

    def walk_nodes(node, current_path):
        for child in node:
            tag = child.tag.split('}')[-1]
            next_path = list(current_path)
            
            if tag in ['Folder', 'Document']:
                name_el = child.find('kml:name', ns)
                if name_el is not None and name_el.text:
                    next_path.append(name_el.text.strip())

            if tag == 'Placemark':
                name_el = child.find('kml:name', ns)
                placemark_name = name_el.text.strip() if name_el is not None and name_el.text else "Untitled Placemark"
                
                # Cek apakah Placemark adalah Polygon (Boundary)
                polygon_elem = child.find('.//kml:Polygon', ns)
                if polygon_elem is not None:
                    coords_elem = polygon_elem.find('.//kml:coordinates', ns)
                    if coords_elem is not None and coords_elem.text:
                        coords = parse_coords(coords_elem.text)
                        if len(coords) >= 3:
                            poly = Polygon(coords)
                            polygon_list.append({
                                "geometry": poly,
                                "name": placemark_name
                            })

                # Cek apakah Placemark adalah Point (HP / FAT / Pole)
                point_elem = child.find('.//kml:Point', ns)
                if point_elem is not None:
                    coords_elem = point_elem.find('.//kml:coordinates', ns)
                    if coords_elem is not None and coords_elem.text:
                        coords = parse_coords(coords_elem.text)
                        if len(coords) >= 1:
                            lon, lat = coords[0]
                            pt = Point(lon, lat)
                            
                            folder_names = [p.lower() for p in next_path]
                            point_data = {
                                "geometry": pt,
                                "name": placemark_name,
                                "lat": lat,
                                "lon": lon
                            }
                            
                            if any(f in ["hp", "homepass"] for f in folder_names):
                                hp_list.append(point_data)
                            elif any(f in ["fat", "odp"] for f in folder_names):
                                fat_list.append(point_data)
                            elif any(f in ["pole", "tiang"] for f in folder_names):
                                pole_list.append(point_data)

            walk_nodes(child, next_path)

    walk_nodes(root, [])
    return polygon_list, hp_list, fat_list, pole_list

# === KONFIGURASI FORMULIR (SIDEBAR) ===
st.sidebar.header("⚙️ Form Isian Data Wilayah & Kode")
st.sidebar.write("Isi data berikut sebelum mengunduh Excel:")

input_prov = st.sidebar.text_input("M : Province", "SUMATERA UTARA")
input_regency = st.sidebar.text_input("N : Regency City", "MEDAN")
input_district = st.sidebar.text_input("O : District", "MEDAN TUNTUNGAN")
input_subdist = st.sidebar.text_input("P : Sub District", "TANJUNG SELAMAT")
input_postal = st.sidebar.text_input("Q : Postal Code", "20251")
input_area = st.sidebar.text_input("W : Area", "MEDAN")
input_cluster = st.sidebar.text_input("Y : Cluster Name", "OPA MEDAN")
input_olt = st.sidebar.text_input("AR : OLT Code", "MDN.100.0802")
input_fdt = st.sidebar.text_input("AS : FDT Code", "BGN-7.068")

st.sidebar.markdown("---")
st.sidebar.subheader("🔧 Pengaturan Format Nama FAT")
fat_col_g_mode = st.sidebar.selectbox(
    "Format Kolom G (Nama FAT):",
    [
        "Gunakan FDT Code (AS) + Kode Boundary (Contoh: BGN-7.068.A01)",
        "Gunakan Nama Asli dari Folder FAT di KML"
    ]
)
fat_col_at_mode = st.sidebar.selectbox(
    "Format Kolom AT (Nama FAT):",
    [
        "Gunakan Nama Asli dari Folder FAT (Contoh: BGN-7.068.A01)",
        "Gunakan FDT Code (AS) + Kode Boundary",
        "Hanya Kode Akhir FAT / Boundary (Contoh: A01)"
    ]
)

# === UPLOAD FILE ===
col_up1, col_up2 = st.columns(2)
with col_up1:
    uploaded_file = st.file_uploader("1️⃣ Pilih File Sumber Data (.kml atau .kmz)", type=["kml", "kmz"])
with col_up2:
    template_excel = st.file_uploader(
        "2️⃣ (Opsional) Upload Contoh.xlsx untuk menyalin Baris 1 (Header) persis sesuai template",
        type=["xlsx"]
    )

if uploaded_file is not None:
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    with st.spinner("Mengekstrak KML/KMZ, memetakan relasi HP-FAT-Pole, dan menyusun Excel..."):
        try:
            kml_content = None
            if file_extension == "kmz":
                with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                    for filename in zip_ref.namelist():
                        if filename.lower().endswith('.kml'):
                            with zip_ref.open(filename) as f:
                                kml_content = f.read()
                            break
            elif file_extension == "kml":
                kml_content = uploaded_file.read()

            if kml_content:
                polygon_list, hp_list, fat_list, pole_list = parse_kml_data(kml_content)

                if not hp_list:
                    st.error("❌ Tidak ditemukan titik di dalam folder HP.")
                else:
                    # 1. PETA: FAT ke BOUNDARY (Polygon)
                    fat_dict = {}
                    for fat in fat_list:
                        matched_poly = None
                        for poly in polygon_list:
                            if poly["geometry"].covers(fat["geometry"]):
                                matched_poly = poly["name"]
                                break
                        fat["boundary"] = matched_poly
                        if matched_poly:
                            fat_dict[matched_poly] = fat

                    # Fallback pencocokan nama jika titik FAT berada sedikit di luar garis poligon
                    for poly in polygon_list:
                        if poly["name"] not in fat_dict:
                            for fat in fat_list:
                                if poly["name"].lower() in fat["name"].lower():
                                    fat["boundary"] = poly["name"]
                                    fat_dict[poly["name"]] = fat
                                    break

                    # 2. PETA: FAT ke POLE (Mencari Pole dengan tikor sama / jarak terdekat)
                    for fat in fat_list:
                        closest_pole = ""
                        min_dist = float('inf')
                        for pole in pole_list:
                            dist = fat["geometry"].distance(pole["geometry"])
                            if dist < min_dist:
                                min_dist = dist
                                closest_pole = pole["name"]
                        fat["closest_pole"] = closest_pole

                    # 3. PETA: HP ke BOUNDARY & PENYUSUNAN DATA PER HP
                    # Pertama, kelompokkan HP berdasarkan Boundary agar urut sesuai nama Boundary (A01, A02, dst)
                    hp_with_boundary = []
                    for hp in hp_list:
                        hp_poly = None
                        for poly in polygon_list:
                            if poly["geometry"].covers(hp["geometry"]):
                                hp_poly = poly["name"]
                                break
                        hp["boundary"] = hp_poly if hp_poly else "ZZZ_OUTSIDE"
                        hp_with_boundary.append(hp)

                    # Urutkan berdasarkan Boundary (A01, A02, ..., A08) dengan mempertahankan urutan asli HP di KML
                    hp_with_boundary.sort(key=lambda x: x["boundary"])

                    results = []
                    fat_hp_counter = {}

                    for hp in hp_with_boundary:
                        hp_poly = hp["boundary"]
                        if hp_poly in fat_dict:
                            fat = fat_dict[hp_poly]
                            fat_raw_name = fat["name"]
                            fat_short_code = fat_raw_name.split('.')[-1] if '.' in fat_raw_name else hp_poly

                            # Hitung urutan HP pada setiap FAT (reset ke 1 setiap ganti FAT)
                            if hp_poly not in fat_hp_counter:
                                fat_hp_counter[hp_poly] = 1
                            else:
                                fat_hp_counter[hp_poly] += 1
                            hp_seq = fat_hp_counter[hp_poly]

                            # Format Kolom G (FDTCODE.FATCODE)
                            if "FDT Code" in fat_col_g_mode:
                                fdt_prefix = input_fdt.strip() if input_fdt.strip() else (
                                    '.'.join(fat_raw_name.split('.')[:-1]) if '.' in fat_raw_name else "FDT"
                                )
                                col_g_val = f"{fdt_prefix}.{fat_short_code}"
                            else:
                                col_g_val = fat_raw_name

                            # Format Kolom AT (Nama FAT)
                            if "Nama Asli" in fat_col_at_mode:
                                col_at_val = fat_raw_name
                            elif "FDT Code" in fat_col_at_mode:
                                fdt_prefix = input_fdt.strip() if input_fdt.strip() else "FDT"
                                col_at_val = f"{fdt_prefix}.{fat_short_code}"
                            else:
                                col_at_val = fat_short_code

                            results.append({
                                "G": col_g_val,
                                "H": hp_seq,
                                "I": fat["closest_pole"],
                                "J": round(fat["lat"], 6),
                                "K": round(fat["lon"], 6),
                                "L": "JL. JALAN",
                                "M": input_prov,
                                "N": input_regency,
                                "O": input_district,
                                "P": input_subdist,
                                "Q": input_postal,
                                "U": "FTTH",
                                "W": input_area,
                                "Y": input_cluster,
                                "AC": "JLN.",
                                "AD": "JALAN",
                                "AJ": hp["name"],
                                "AR": input_olt,
                                "AS": input_fdt,
                                "AT": col_at_val,
                                "AV": hp["lat"],
                                "AW": hp["lon"],
                                "AX": "T",
                                "AY": "F",
                                "AZ": "F",
                                "BA": "AE",
                                "BB": "T",
                                "BC": 31,
                                "BD": "F"
                            })

                    # === PENYUSUNAN TABEL 56 KOLOM (A SAMPAI BD) ===
                    total_cols = 56  # A (0) sampai BD (55)
                    col_letters = [get_excel_col_name(i) for i in range(total_cols)]

                    # Header default baris 1 jika tidak mengunggah template Contoh.xlsx
                    default_headers = {
                        "G": "FAT Code",
                        "H": "HP No",
                        "I": "Pole ID",
                        "J": "FAT Latitude",
                        "K": "FAT Longitude",
                        "L": "Street Prefix",
                        "M": "Province",
                        "N": "Regency City",
                        "O": "District",
                        "P": "Sub District",
                        "Q": "Postal Code",
                        "U": "Technology",
                        "W": "Area",
                        "Y": "Cluster Name",
                        "AC": "Street Short",
                        "AD": "Street Name",
                        "AJ": "HP Name",
                        "AR": "OLT Code",
                        "AS": "FDT Code",
                        "AT": "FAT Name",
                        "AV": "HP Latitude",
                        "AW": "HP Longitude",
                        "AX": "AX",
                        "AY": "AY",
                        "AZ": "AZ",
                        "BA": "BA",
                        "BB": "BB",
                        "BC": "BC",
                        "BD": "BD"
                    }

                    row1_headers = [default_headers.get(letter, "") for letter in col_letters]

                    # Jika user mengunggah Contoh.xlsx, salin Baris 1 (Header) dari Contoh.xlsx
                    if template_excel is not None:
                        try:
                            tpl_df = pd.read_excel(template_excel, header=None, nrows=1)
                            for c_idx in range(min(total_cols, tpl_df.shape[1])):
                                val = tpl_df.iloc[0, c_idx]
                                row1_headers[c_idx] = "" if pd.isna(val) else str(val)
                            st.info("ℹ️ Baris 1 (Header) berhasil disalin langsung dari file template Excel Anda.")
                        except Exception as ex:
                            st.warning(f"Tidak dapat membaca header template Excel: {ex}")

                    # Buat matriks data (setiap baris = 1 HP)
                    matrix_data = []
                    for res in results:
                        row_vals = [""] * total_cols
                        for col_letter, val in res.items():
                            c_idx = excel_col_to_idx(col_letter)
                            if 0 <= c_idx < total_cols:
                                row_vals[c_idx] = val
                        matrix_data.append(row_vals)

                    # DataFrame untuk di-export ke Excel
                    export_df = pd.DataFrame(matrix_data, columns=col_letters)

                    st.success(
                        f"✅ Berhasil memproses **{len(results)} Homepass (HP)** dalam **{len(fat_hp_counter)} Boundary/FAT**!"
                    )

                    # Tampilkan Pratinjau Kolom-Kolom Utama yang Terisi
                    st.subheader("📋 Pratinjau Kolom yang Terisi (G, H, I, J, K, L..Q, U, W, Y, AC, AD, AJ, AR, AS, AT, AV, AW, AX..BD)")
                    filled_cols = [
                        "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q",
                        "U", "W", "Y", "AC", "AD", "AJ", "AR", "AS", "AT", "AV", "AW",
                        "AX", "AY", "AZ", "BA", "BB", "BC", "BD"
                    ]
                    preview_df = export_df[filled_cols].copy()
                    preview_df.columns = [
                        f"{c} ({row1_headers[excel_col_to_idx(c)] or c})" for c in filled_cols
                    ]
                    st.dataframe(preview_df, use_container_width=True)

                    # === EXPORT KE FILE EXCEL (.XLSX) ===
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        worksheet = workbook.add_worksheet('Sheet1')

                        header_format = workbook.add_format({
                            'bold': True,
                            'valign': 'vcenter',
                            'fg_color': '#FFFF00',
                            'border': 1
                        })
                        cell_format = workbook.add_format({
                            'valign': 'vcenter'
                        })

                        # Tulis Baris 1 (Header)
                        for col_idx, h_text in enumerate(row1_headers):
                            if h_text != "":
                                worksheet.write(0, col_idx, h_text, header_format)
                            else:
                                worksheet.write(0, col_idx, "")

                        # Tulis Data mulai Baris 2 (indeks 1)
                        for r_idx, row_vals in enumerate(matrix_data):
                            for c_idx, val in enumerate(row_vals):
                                worksheet.write(r_idx + 1, c_idx, val, cell_format)

                        # Atur lebar kolom otomatis pada kolom yang berisi data
                        worksheet.set_column('G:Q', 18)
                        worksheet.set_column('U:Y', 15)
                        worksheet.set_column('AC:AD', 12)
                        worksheet.set_column('AJ:AJ', 20)
                        worksheet.set_column('AR:AT', 18)
                        worksheet.set_column('AV:AW', 16)
                        worksheet.set_column('AX:BD', 8)

                    output.seek(0)

                    st.download_button(
                        label="📥 Unduh Hasil Convert Excel (.xlsx)",
                        data=output,
                        file_name="Hasil_Convert_FTTH.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file: {e}")
