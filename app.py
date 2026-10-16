import streamlit as st
import zipfile
import os
import io
import xml.etree.ElementTree as ET
import pandas as pd
from shapely.geometry import Point, Polygon

# === KONFIGURASI HALAMAN ===
st.set_page_config(page_title="FTTH KML/KMZ to Excel", layout="wide")

st.title("🗺️ Pemrosesan FTTH: Boundary, FAT, Pole, & HP ke Excel")
st.write("Unggah file **.kmz** atau **.kml** Anda. Sistem akan mengklasifikasikan titik berdasarkan foldernya (HP, Pole, FAT) dan memetakannya ke dalam area Boundary (Polygon).")

# === FUNGSI PARSER KML ===
def parse_kml_data(file_content):
    polygon_list = []
    point_list = []

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

    # Fungsi rekursif untuk menelusuri XML dan melacak hierarki folder
    def walk_nodes(node, current_path):
        for child in node:
            tag = child.tag.split('}')[-1]
            next_path = list(current_path)
            
            # Jika elemen adalah Folder atau Document, tambahkan namanya ke path
            if tag in ['Folder', 'Document']:
                name_el = child.find('kml:name', ns)
                if name_el is not None and name_el.text:
                    next_path.append(name_el.text.strip())

            # Jika elemen adalah Placemark (Titik atau Poligon)
            if tag == 'Placemark':
                name_el = child.find('kml:name', ns)
                placemark_name = name_el.text.strip() if name_el is not None and name_el.text else "Untitled"
                
                # Cek apakah ini Polygon (Boundary)
                polygon_elem = child.find('.//kml:Polygon', ns)
                if polygon_elem is not None:
                    coords_elem = polygon_elem.find('.//kml:coordinates', ns)
                    if coords_elem is not None and coords_elem.text:
                        coords = parse_coords(coords_elem.text)
                        if len(coords) >= 3:
                            poly = Polygon(coords)
                            polygon_list.append({
                                "geometry": poly,
                                "Boundary_Name": placemark_name
                            })

                # Cek apakah ini Point (HP, Pole, FAT)
                point_elem = child.find('.//kml:Point', ns)
                if point_elem is not None:
                    coords_elem = point_elem.find('.//kml:coordinates', ns)
                    if coords_elem is not None and coords_elem.text:
                        coords = parse_coords(coords_elem.text)
                        if len(coords) >= 1:
                            lon, lat = coords[0]
                            pt = Point(lon, lat)
                            
                            # Identifikasi Tipe Point berdasarkan Struktur Folder KML
                            path_str = " ".join(next_path).lower()
                            tipe_titik = "Lainnya"
                            if "hp" in path_str:
                                tipe_titik = "HP"
                            elif "pole" in path_str or "tiang" in path_str:
                                tipe_titik = "Pole"
                            elif "fat" in path_str or "fdt" in path_str:
                                tipe_titik = "FAT"

                            # Hanya masukkan HP, Pole, atau FAT
                            if tipe_titik != "Lainnya":
                                point_list.append({
                                    "geometry": pt,
                                    "Latitude": lat,
                                    "Longitude": lon,
                                    "Point_Name": placemark_name,
                                    "Type": tipe_titik
                                })

            walk_nodes(child, next_path)

    walk_nodes(root, [])
    return polygon_list, point_list

# === KONFIGURASI FORMULIR (SIDEBAR) ===
st.sidebar.header("⚙️ Form Isian Data Wilayah & Kode")
st.sidebar.write("Isi data berikut sebelum mendownload Excel:")

input_prov = st.sidebar.text_input("Province (Kolom M)", "SUMATERA UTARA")
input_regency = st.sidebar.text_input("Regency City (Kolom N)", "MEDAN")
input_district = st.sidebar.text_input("District (Kolom O)", "")
input_subdist = st.sidebar.text_input("Sub District (Kolom P)", "")
input_postal = st.sidebar.text_input("Postal Code (Kolom Q)", "")
input_area = st.sidebar.text_input("Area (Kolom W)", "")
input_cluster = st.sidebar.text_input("Cluster Name (Kolom Y)", "")
input_olt = st.sidebar.text_input("OLT Code (Kolom AR)", "")
input_fdt = st.sidebar.text_input("FDT Code (Kolom AS)", "")

# === UPLOAD FILE ===
uploaded_file = st.file_uploader("Pilih file KML atau KMZ", type=["kml", "kmz"])

if uploaded_file is not None:
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    with st.spinner("Mengekstrak dan memproses koordinat..."):
        try:
            kml_content = None
            
            # Jika KMZ, ekstrak dan cari file KML di dalamnya
            if file_extension == "kmz":
                with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                    for filename in zip_ref.namelist():
                        if filename.endswith('.kml'):
                            with zip_ref.open(filename) as f:
                                kml_content = f.read()
                            break
                if kml_content is None:
                    st.error("❌ File .kml tidak ditemukan di dalam arsip KMZ.")
            
            # Jika KML, langsung baca isinya
            elif file_extension == "kml":
                kml_content = uploaded_file.read()

            if kml_content:
                # Parsing KML
                polygon_list, point_list = parse_kml_data(kml_content)

                if not polygon_list:
                    st.warning("⚠️ Tidak ditemukan area (Polygon/Boundary) di dalam file. Spasial join tidak dapat dilakukan.")
                if not point_list:
                    st.error("❌ Tidak ditemukan titik (HP/Pole/FAT) di file Anda.")
                
                if point_list:
                    results = []
                    for pt_data in point_list:
                        pt_geom = pt_data["geometry"]
                        matched_poly = "Di Luar Boundary"

                        # Cek setiap titik masuk ke poligon mana
                        for poly_data in polygon_list:
                            if poly_data["geometry"].contains(pt_geom):
                                matched_poly = poly_data["Boundary_Name"]
                                break

                        results.append({
                            "Boundary_Area": matched_poly,
                            "Type": pt_data["Type"],
                            "Point_Name": pt_data["Point_Name"],
                            "Longitude": pt_data["Longitude"],
                            "Latitude": pt_data["Latitude"],
                        })

                    df = pd.DataFrame(results)
                    
                    # Urutkan secara sistematis (Boundary -> Tipe (FAT, Pole, HP) -> Nama Titik)
                    tipe_kasta = {"FAT": 1, "Pole": 2, "HP": 3}
                    df["Type_Order"] = df["Type"].map(tipe_kasta)
                    df = df.sort_values(by=["Boundary_Area", "Type_Order", "Point_Name"]).drop(columns=["Type_Order"]).reset_index(drop=True)
                    
                    # Tambahkan nomor urut
                    df.insert(0, 'No', df.index + 1)

                    # Buat Tabel Rekapitulasi (Pivot Table)
                    summary_df = pd.pivot_table(
                        df, 
                        values='Point_Name', 
                        index='Boundary_Area', 
                        columns='Type', 
                        aggfunc='count', 
                        fill_value=0
                    ).reset_index()
                    
                    # Pastikan kolom standar ada meskipun datanya tidak ada
                    for col in ["FAT", "Pole", "HP"]:
                        if col not in summary_df.columns:
                            summary_df[col] = 0
                            
                    summary_df.rename(columns={"FAT": "Total_FAT", "Pole": "Total_Pole", "HP": "Total_HP"}, inplace=True)
                    summary_df = summary_df[["Boundary_Area", "Total_FAT", "Total_Pole", "Total_HP"]]

                    st.success("✅ File berhasil dikonversi dan direkap!")
                    
                    # Tampilkan di UI (Hanya Data Penting)
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        st.subheader("Data Detail (Pratinjau)")
                        st.dataframe(df, use_container_width=True, hide_index=True)
                    with col2:
                        st.subheader("Rekapitulasi per Boundary")
                        st.dataframe(summary_df, use_container_width=True, hide_index=True)

                    # === PEMBUATAN TEMPLATE EXCEL (A - BD) ===
                    # Fungsi pembantu untuk konversi index ke Huruf Excel
                    def get_excel_col_name(col_idx):
                        result = ""
                        while col_idx >= 0:
                            result = chr(col_idx % 26 + 65) + result
                            col_idx = col_idx // 26 - 1
                        return result
                    
                    # Mapping header khusus berdasarkan index (0 = A, 1 = B, 55 = BD)
                    header_mapping = {
                        0: 'No', 1: 'Boundary', 2: 'Type', 3: 'Point_Name', 4: 'Longitude', 5: 'Latitude',
                        11: 'Jalan_1',            # Kolom L
                        12: 'Province',           # Kolom M
                        13: 'Regency City',       # Kolom N
                        14: 'District',           # Kolom O
                        15: 'Sub District',       # Kolom P
                        16: 'Postal Code',        # Kolom Q
                        20: 'Teknologi',          # Kolom U
                        22: 'Area',               # Kolom W
                        24: 'Cluster Name',       # Kolom Y
                        28: 'Jalan_2',            # Kolom AC
                        29: 'Jalan_3',            # Kolom AD
                        43: 'OLT Code',           # Kolom AR
                        44: 'FDT Code',           # Kolom AS
                        49: 'Flag_AX',            # Kolom AX
                        50: 'Flag_AY',            # Kolom AY
                        51: 'Flag_AZ',            # Kolom AZ
                        52: 'Flag_BA',            # Kolom BA
                        53: 'Flag_BB',            # Kolom BB
                        54: 'Flag_BC',            # Kolom BC
                        55: 'Flag_BD'             # Kolom BD
                    }

                    col_names = []
                    for i in range(56): # Membuat 56 kolom (A - BD)
                        if i in header_mapping:
                            col_names.append(header_mapping[i])
                        else:
                            col_names.append(f"Col_{get_excel_col_name(i)}")

                    # Buat DataFrame Lengkap
                    full_df = pd.DataFrame(columns=col_names)
                    
                    # 1. Isi Data Utama
                    full_df['No'] = df['No']
                    full_df['Boundary'] = df['Boundary_Area']
                    full_df['Type'] = df['Type']
                    full_df['Point_Name'] = df['Point_Name']
                    full_df['Longitude'] = df['Longitude']
                    full_df['Latitude'] = df['Latitude']

                    # 2. Isi Data Statis (Sesuai Permintaan)
                    full_df['Jalan_1'] = "JL. JALAN"      # L
                    full_df['Teknologi'] = "FTTH"         # U
                    full_df['Jalan_2'] = "JLN."           # AC
                    full_df['Jalan_3'] = "JALAN"          # AD
                    full_df['Flag_AX'] = "T"              # AX
                    full_df['Flag_AY'] = "F"              # AY
                    full_df['Flag_AZ'] = "F"              # AZ
                    full_df['Flag_BA'] = "AE"             # BA
                    full_df['Flag_BB'] = "T"              # BB
                    full_df['Flag_BC'] = "31"             # BC
                    full_df['Flag_BD'] = "F"              # BD

                    # 3. Isi Data Dinamis (Dari Input User di Sidebar)
                    full_df['Province'] = input_prov           # M
                    full_df['Regency City'] = input_regency    # N
                    full_df['District'] = input_district       # O
                    full_df['Sub District'] = input_subdist    # P
                    full_df['Postal Code'] = input_postal      # Q
                    full_df['Area'] = input_area               # W
                    full_df['Cluster Name'] = input_cluster    # Y
                    full_df['OLT Code'] = input_olt            # AR
                    full_df['FDT Code'] = input_fdt            # AS

                    # Isi NaN dengan string kosong agar rapi di Excel
                    full_df = full_df.fillna("")

                    # === EXPORT EXCEL (XLSXWRITER) ===
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # 1. Sheet Detail Data
                        full_df.to_excel(writer, index=False, sheet_name='Data_FTTH')
                        workbook = writer.book
                        worksheet_detail = writer.sheets['Data_FTTH']

                        # Definisi Format Style Excel
                        header_format = workbook.add_format({
                            'bold': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1
                        })
                        border_format = workbook.add_format({'border': 1})
                        
                        # Terapkan format ke header
                        for col_num, value in enumerate(full_df.columns.values):
                            # Jika nama kolom berawalan Col_, berarti itu kolom kosong bayangan, biarkan headernya kosong atau tetap
                            header_val = value if not str(value).startswith("Col_") else ""
                            worksheet_detail.write(0, col_num, header_val, header_format)
                            
                        # Atur lebar kolom yang ada datanya
                        worksheet_detail.set_column('A:A', 5, border_format)   # No
                        worksheet_detail.set_column('B:B', 15, border_format)  # Boundary
                        worksheet_detail.set_column('C:C', 10, border_format)  # Type
                        worksheet_detail.set_column('D:D', 25, border_format)  # Point_Name
                        worksheet_detail.set_column('E:F', 18, border_format)  # Long/Lat
                        worksheet_detail.set_column('M:Q', 15, border_format)  # Wilayah
                        worksheet_detail.set_column('W:Y', 15, border_format)  # Area/Cluster
                        worksheet_detail.set_column('AR:AS', 15, border_format)# OLT/FDT
                        
                        # Tambahkan auto-filter dari A sampai BD
                        worksheet_detail.autofilter(0, 0, len(full_df), len(full_df.columns) - 1)

                        # 2. Sheet Rekapitulasi
                        summary_df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
                        worksheet_summary = writer.sheets['Rekapitulasi']
                        
                        header_format_sum = workbook.add_format({
                            'bold': True, 'fg_color': '#B4C6E7', 'border': 1, 'align': 'center'
                        })
                        
                        for col_num, value in enumerate(summary_df.columns.values):
                            worksheet_summary.write(0, col_num, value, header_format_sum)

                        worksheet_summary.set_column('A:A', 20, border_format)
                        worksheet_summary.set_column('B:D', 15, workbook.add_format({'border': 1, 'align': 'center'}))

                    output.seek(0)

                    # Tombol Download
                    st.download_button(
                        label="📥 Unduh File Excel (.xlsx)",
                        data=output,
                        file_name="Hasil_Klasifikasi_FTTH.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file: {e}")