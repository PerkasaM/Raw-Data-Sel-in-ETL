"""
ETL Sell-In Bulanan
====================
Cara pakai tiap bulan:
1. Isi bagian CONFIG di bawah (path file & info cycle)
2. Jalankan: python sellin_etl.py
3. Output: file Excel siap pakai di folder yang sama

Flow:
  editsellin  → proses data SAP bulan ini jadi raw_new yang clean
  concatsellin → gabung raw_new ke raw_old + update alamat + handle duplikat DR Number
"""

import pandas as pd

# ============================================================
# CONFIG — update bagian ini setiap bulan
# ============================================================

PATH_RAW_OLD    = r"C:\Users\USER\Documents\MEVAL\Raw data\Raw Data Sell IN - 2023-2026 (C0726) Rev 1.xlsx"
PATH_TEMPLATE   = r"C:\Users\USER\Documents\MEVAL\TEMPLATE\2026\TEMPLATE_SELL_IN_SAP 040426.xlsx"
PATH_SAP        = r"C:\Users\USER\Documents\SAP\SAP GUI\export customermasterlist 26082026.XLSX"
PATH_MTD_YTD    = r"C:\Users\USER\Documents\MEVAL\MTD YTD\2026\C08\MTD YTD REPORT C08 27.08.2026.xlsx"
PATH_SDO_UPDATE = r"C:\Users\USER\Documents\MEVAL\SDO\SDO UPDATE C07_ALL_AREA.xlsx"
PATH_MD_SKU     = r"C:\Users\USER\Documents\MEVAL\Master Data\skuu6.xlsx"
PATH_SPVRSM     = r"C:\Users\USER\Documents\MEVAL\Master Data\spv rsm.xlsx"
PATH_MS_DC      = r"C:\Users\USER\Documents\MEVAL\Master Data\ms dc.xlsx"
PATH_MD_SUBPRODUK     = r"C:\Users\USER\Documents\MEVAL\Master Data\sub produk.xlsx"
PATH_GROUP      = r"C:\Users\USER\Documents\MEVAL\Master Data\Group.xlsx"
PATH_SWM        = r"C:\Users\USER\Documents\MEVAL\Master Data\SWM_grouping_.xlsx"
PATH_PRICELIST  = r"C:\Users\USER\Documents\MEVAL\Master Data\PRICELIST MEVAL 01042026 INT.xlsx"
PATH_SDO_AKTIF  = r"C:\Users\USER\Documents\MEVAL\Generate\selllin\data\SDO aktif.xlsx"
PATH_PRODUCT    = r"C:\Users\USER\Documents\MEVAL\Master Data\Product Data.xlsx"

# Info cycle bulan ini — ganti setiap bulan
CYCLE       = "C08"
DUMMY_CYCLE = "C0826"
MTD_SHEET   = "SAP CUMULATIVE"

# File output
OUTPUT_RAW_NEW = f"raw_new_{DUMMY_CYCLE}.xlsx"
OUTPUT_FINAL   = f"raw_data_{DUMMY_CYCLE}_rev1.xlsx"

# ============================================================
# BAGIAN 1: Helper functions
# ============================================================

def cleaning_type(tipe):
    """Normalisasi customer group dari SAP ke format MD_TOKO."""
    if pd.isna(tipe):
        return None
    t = str(tipe).upper().strip()
    if '2.0' in t or '2' in t:
        return 'RETAIL'
    elif '3.0' in t or '3' in t:
        return 'GROSIR'
    elif '5.0' in t or '5' in t:
        return 'OTHERS'
    elif '16.0' in t or '6' in t:
        return 'SEMI-GROSIR'
    return t


def cleaning_region(region):
    """Normalisasi nama region."""
    if pd.isna(region):
        return None
    r = str(region).upper().strip()
    if 'EAST JAVA' in r:
        return 'EAST JAVA'
    elif 'CENTRAL JAVA' in r:
        return 'CENTRAL JAVA'
    elif 'JABODETABEK' in r:
        return 'JABODETABEK'
    return r


def normalize_customer_code(code):
    """
    Normalisasi Customer Code:
    - Baca sebagai string (sudah dilakukan saat load dengan dtype=str)
    - Angka murni (tanpa huruf) → pad dengan leading zero ke 7 digit
    - Alfanumerik / text → biarkan apa adanya, strip whitespace saja
    Contoh:
      '537'      → '0000537'
      '1006288'  → '1006288'   (sudah 7 digit)
      '000001'   → '000001'    (sudah ada leading zero, bukan angka murni 7 digit → biarkan)
      'B0000118' → 'B0000118'  (ada huruf → biarkan)
      'NULL01'   → 'NULL01'    (text → biarkan)
    """
    if pd.isna(code):
        return None
    s = str(code).strip()
    if s.isdigit():
        return s.zfill(7)
    return s


def cleaning_sdo(sdo):
    """
    Normalisasi nama SDO agar cocok dengan MD_SDO dan file SDO aktif.
    Mapping: nama variasi/typo → nama standar di master.
    """
    if pd.isna(sdo):
        return None
    s = str(sdo).upper().strip()

    # Hapus prefix/suffix yang tidak relevan
    s = s.replace('MO ', '', 1) if s.startswith('MO ') else s  # "MO MOHAMMAD..." → "MOHAMMAD..."

    mapping = {
        # Nama lengkap → nama standar
        'AHMAD IMAM TAUFIQ':                'IMAM TAUFIQ',
        'IMAM TAUFIK':                      'IMAM TAUFIQ',
        "IMAM SHOVI'I":                     'IMAM SHOVII',
        "IMAM SHOVI":                       'IMAM SHOVII',
        'RIZAL ECOM':                       'E-COMMERCE',
        'ANDY WIJAYA':                      'ANDY',
        'ABDUL JABBAR':                     'ABDUL JABAR',
        'ARDIAN DWI P.':                    'ARDIAN DWI P',
        'ARDI YULI KURNAWAN':               'ARDI YULI KURNIAWAN',
        'MULYADI':                          'MULYADI M',
        'ANDI MANGGALA':                    'ANDI MANGGALA PUTRA',
        'MOHAMMAD HENDRA FARIZAL':          'MOHAMMAD HENDRA FARIZAL',
        'MUHAMMAD LANANG GALIH':            'MUHAMMAD LANANG GALIH GUMILANG',
        'LANANG GALIH':                     'MUHAMMAD LANANG GALIH GUMILANG',
        'RISKY WAHYUDI':                    'RIZKI WAHYUDI',
        'RIZKY WAHYUDI':                    'RIZKI WAHYUDI',
        'ABDULLAH RAYHAN':                  'ABDULLAH',
        # Typo umum
        'ABDILLLAH':                        'ABDILLAH',
        'ACH LUKMAN':                       'ACH LUKMAN HAKIM',
        'AGUNG SEDAYO':                     'AGUNG SEDAYU',
        'AHA JUNI':                         'AHA JUNI ADI',
    }
    for k, v in mapping.items():
        if s == k:          # exact match dulu (lebih aman dari substring)
            return v
    # Substring match hanya untuk prefix pendek yang ambigu
    SUBSTR_MAP = {
        'RIZAL ECOM':   'E-COMMERCE',
        'ANDI MANGGALA': 'ANDI MANGGALA PUTRA',
    }
    for k, v in SUBSTR_MAP.items():
        if k in s:
            return v
    return s


def parse_sap_date(series):
    """
    Convert kolom tanggal dari MTD YTD yang bisa berupa:
    - datetime sudah ter-parse
    - Excel serial number (integer, misal 45800)
    - String tanggal
    Mengembalikan Series datetime64.
    """
    # Coba convert biasa dulu
    result = pd.to_datetime(series, errors='coerce')

    # Jika hasilnya semua 1970 atau NaT, berarti format angka Excel serial
    if result.isna().all() or (result.dt.year == 1970).all():
        # Excel serial: 1 = 1 Jan 1900, tapi ada bug Lotus 1900 leap year
        # Pandas bisa handle dengan unit='D' dari origin '1899-12-30'
        result = pd.to_datetime(series, unit='D', origin='1899-12-30', errors='coerce')

    return result


def realqty(internal_code):
    """Hitung multiplier qty berdasarkan kode produk (paket isi 4, 3, 5, dll)."""
    if pd.isna(internal_code):
        return None
    c = str(internal_code)
    if any(x in c for x in ['AB4', 'EB4', 'BL4', 'XB4']):
        return 4
    elif any(x in c for x in ['AB3', 'BBT', 'EB3', 'ER3', 'BB2']):
        return 3
    elif 'BBL' in c:
        return 5
    return 1


# ============================================================
# BAGIAN 2: Load master data
# ============================================================

def load_master_data():
    """Load semua master data yang dibutuhkan."""
    print("Loading master data...")

    cols_toko = ['Customer Code', 'Customer Name', 'Address', 'City',
                 'Provinsi', 'Region', 'Reg2', 'type', 'TypeDummy', 'Chanel']
    md_toko = pd.read_excel(PATH_TEMPLATE, sheet_name='MD_TOKO', usecols=cols_toko)
    md_toko['Customer Code'] = md_toko['Customer Code'].astype(str).apply(normalize_customer_code)

    md_sdo = pd.read_excel(PATH_TEMPLATE, sheet_name='MD_SDO')
    md_sdo['Customer Code'] = md_sdo['Customer Code'].astype(str).apply(normalize_customer_code)

    md_sku = pd.read_excel(PATH_MD_SKU)
    md_sku = md_sku.rename(columns={
        'Description': 'Desc',
        'SUB-PG':      'SUBPG',
        'Short Code':  'Internal Code'
    })

    spvrsm = pd.read_excel(PATH_SPVRSM, sheet_name='Sheet2')

    cols_sap = ["Customer", "Account group", "Name 1", "Street", "City",
                "Search term", "Region", "Language Key", "Transportation zone", "Zip Code"]
    sap = pd.read_excel(PATH_SAP)
    sap = sap.iloc[:, :len(cols_sap)]
    sap.columns = cols_sap
    sap['Customer'] = sap['Customer'].astype(str).apply(normalize_customer_code)

    return md_toko, md_sdo, md_sku, spvrsm, sap


# ============================================================
# BAGIAN 3: Update MD_TOKO dengan customer baru
# ============================================================

def update_md_toko(md_toko, df_mtd, df_unique, sap, spvrsm):
    print("Updating MD_TOKO...")

    df_check = pd.merge(df_mtd[['Customer Code']], md_toko, on='Customer Code', how='left')
    missing  = df_check.loc[df_check['Customer Name'].isna(), ['Customer Code']].drop_duplicates()

    if missing.empty:
        print("  Tidak ada customer baru.")
        return md_toko

    new_rows = pd.merge(missing, sap[['Customer', 'Name 1', 'Street', 'City']],
                        left_on='Customer Code', right_on='Customer', how='left')
    new_rows = new_rows.rename(columns={'Name 1': 'Customer Name', 'Street': 'Address'})
    new_rows = new_rows.drop(columns=['Customer'], errors='ignore')

    new_rows = pd.merge(new_rows,
                        df_unique[['Customer Code', 'RSM', 'CUST GRP', 'CUST TYPE', 'SDO']],
                        on='Customer Code', how='left')
    new_rows['CUST GRP'] = new_rows['CUST GRP'].astype(str).apply(cleaning_type)
    new_rows['SDO']      = new_rows['SDO'].str.upper()

    new_rows = pd.merge(new_rows, spvrsm[['SDO', 'region']], on='SDO', how='left')

    new_rows['Reg2']      = new_rows['region']
    new_rows['TypeDummy'] = new_rows['CUST GRP']
    new_rows = new_rows.rename(columns={'RSM': 'Region', 'CUST GRP': 'type', 'CUST TYPE': 'Chanel'})
    new_rows['Region'] = new_rows['Region'].apply(cleaning_region)

    # Buat city → Region/Reg2 map dari MD_TOKO yang sudah ada
    city_region_map = (md_toko[['City', 'Region']].dropna()
                       .drop_duplicates('City').set_index('City')['Region'])
    city_reg2_map   = (md_toko[['City', 'Reg2']].dropna()
                       .drop_duplicates('City').set_index('City')['Reg2'])

    # Untuk MODERN TRADE dan PROJECT: selalu pakai city lookup (bukan dari RSM)
    mt_project = new_rows['Chanel'].isin(['MODERN TRADE', 'PROJECTS', 'PROJECT'])
    new_rows.loc[mt_project, 'Region'] = new_rows.loc[mt_project, 'City'].map(city_region_map)
    new_rows.loc[mt_project, 'Reg2']   = new_rows.loc[mt_project, 'City'].map(city_reg2_map)

    # Untuk channel lain yang region-nya masih kosong/tidak valid: fallback ke city lookup
    invalid = (~mt_project) & (new_rows['Region'].isna() | new_rows['Region'].isin(['PROJECT', 'MODERN TRADE', '']))
    new_rows.loc[invalid, 'Region'] = new_rows.loc[invalid, 'City'].map(city_region_map)
    new_rows.loc[invalid, 'Reg2']   = new_rows.loc[invalid, 'City'].map(city_reg2_map)

    new_rows.loc[new_rows['Chanel'] == 'MODERN TRADE', 'type'] = 'MODERN TRADE'
    new_rows = new_rows.drop(columns=['SDO', 'region'], errors='ignore')
    new_rows['type'] = new_rows['type'].astype(str)

    cols_order = ['Customer Code', 'Customer Name', 'Address', 'City',
                  'Provinsi', 'Region', 'Reg2', 'type', 'TypeDummy', 'Chanel']
    for col in cols_order:
        if col not in new_rows.columns:
            new_rows[col] = None
    new_rows = new_rows[cols_order]

    md_toko_updated = pd.concat([md_toko, new_rows], ignore_index=True)
    md_toko_updated = md_toko_updated.drop_duplicates('Customer Code', keep='last')

    print(f"  Ditambahkan {len(new_rows)} customer baru ke MD_TOKO.")
    return md_toko_updated


# ============================================================
# BAGIAN 4: Update MD_SDO dengan SDO baru
# ============================================================

def update_md_sdo(md_sdo, df_mtd, sdo_update):
    print("Updating MD_SDO...")

    sdo_update['Customer Code'] = sdo_update['Customer Code'].astype(str).apply(normalize_customer_code)
    md_sdo1 = md_sdo[['Customer Code', 'Customer Name', 'CURRENT SDO']].copy()
    md_sdo1['Customer Code'] = md_sdo1['Customer Code'].astype(str).apply(normalize_customer_code)

    sdo_map = (
        sdo_update[['Customer Code', 'SDO Update']]
        .dropna(subset=['SDO Update'])
        .drop_duplicates('Customer Code', keep='last')
        .set_index('Customer Code')['SDO Update']
    )

    mask_update = md_sdo1['Customer Code'].isin(sdo_map.index)
    md_sdo1.loc[mask_update, 'CURRENT SDO'] = md_sdo1.loc[mask_update, 'Customer Code'].map(sdo_map)
    print(f"  {mask_update.sum()} customer SDO di-update dari file konfirmasi.")

    df_check    = pd.merge(df_mtd[['Customer Code']].drop_duplicates(), md_sdo1,
                           on='Customer Code', how='left')
    missing_sdo = df_check.loc[df_check['CURRENT SDO'].isna(), ['Customer Code']].drop_duplicates()

    if missing_sdo.empty:
        print("  Tidak ada customer baru tanpa SDO.")
        return md_sdo1

    new_sdo = missing_sdo.copy()
    new_sdo['CURRENT SDO'] = new_sdo['Customer Code'].map(sdo_map)

    mtd_sdo_map = (
        df_mtd[['Customer Code', 'SDO']].dropna(subset=['SDO'])
        .drop_duplicates('Customer Code', keep='last')
        .set_index('Customer Code')['SDO']
        .str.upper().apply(cleaning_sdo)
    )
    fallback_mask = new_sdo['CURRENT SDO'].isna()
    new_sdo.loc[fallback_mask, 'CURRENT SDO'] = new_sdo.loc[fallback_mask, 'Customer Code'].map(mtd_sdo_map)

    print(f"  {len(new_sdo)} customer baru ditambahkan "
          f"({(~fallback_mask).sum()} dari konfirmasi, {fallback_mask.sum()} fallback MTD).")

    md_sdo_updated = pd.concat([md_sdo1, new_sdo], ignore_index=True)
    md_sdo_updated = md_sdo_updated.drop_duplicates('Customer Code', keep='last')
    return md_sdo_updated


# ============================================================
# BAGIAN 5: Proses data bulan ini
# ============================================================

def process_new_month(md_toko, md_sdo, md_sdo_updated, md_sku, spvrsm,
                      md_yee, md_group, md_swm, df):
    print("Processing data bulan ini...")

    df = df.reset_index(drop=True)

    # ----------------------------------------------------------------
    # PENTING: Embed semua kolom dari df ke raw_new SEBELUM merge apapun.
    # Ini mencegah mismatch panjang saat merge mengembangkan baris.
    # Setelah merge, kolom-kolom ini sudah ikut terbawa otomatis.
    # ----------------------------------------------------------------
    raw_new = pd.DataFrame({
        'Customer Code':        df['CUSTOMER#'].astype(str).apply(normalize_customer_code),
        '_row_idx':             df.index,                          # jangkar untuk join balik ke df jika perlu
        'Sales Group':          df['SALES GRP'],
        'PLANT':                df['PLANT'],
        'STORELOC':             df['STORELOC'],
        'SDO Name':             df['SDO'].astype(str).str.upper().apply(cleaning_sdo),
        'SKU Code':             df['MATERIAL'],
        'Total (with VAT)':     pd.to_numeric(df['NET VALUE WITH VAT'], errors='coerce'),
        'QTY':                  pd.to_numeric(df['QUANTITY'], errors='coerce'),
        'Total Net (Non VAT)':  pd.to_numeric(df['NET VALUE'], errors='coerce'),
        'DR Number':            df['DOCNUM'],
        'SAP Date':             parse_sap_date(df['ACTUAL PGI DATE']),
        'Status':               df['TYPE'],
    })

    # --- Join MD_TOKO (type → Type) ---
    md_toko_join = md_toko.rename(columns={'type': 'Type'}).drop_duplicates('Customer Code')
    raw_new = pd.merge(raw_new, md_toko_join, on='Customer Code', how='left')

    # Fallback Customer Name dari kolom NAME1 di MTD jika tidak ada di MD_TOKO
    if 'NAME1' in df.columns:
        mtd_name_map = (df[['CUSTOMER#', 'NAME1']]
                        .rename(columns={'CUSTOMER#': 'Customer Code', 'NAME1': '_name_fallback'})
                        .drop_duplicates('Customer Code')
                        .set_index('Customer Code')['_name_fallback'])
        missing_name = raw_new['Customer Name'].isna()
        raw_new.loc[missing_name, 'Customer Name'] = raw_new.loc[missing_name, 'Customer Code'].map(mtd_name_map)

    # --- Join SDO Update ---
    raw_new = pd.merge(
        raw_new,
        md_sdo_updated[['Customer Code', 'CURRENT SDO']].rename(columns={'CURRENT SDO': 'SDO Update'})
        .drop_duplicates('Customer Code'),
        on='Customer Code', how='left'
    )

    # --- Join Status SDO ---
    md_sdoaktif = (md_sdo[['SDO Name', 'STATUS SDO']].copy()
                   .rename(columns={'STATUS SDO': 'Status SDO'})
                   .drop_duplicates('SDO Name')
                   .dropna(thresh=1))
    raw_new = pd.merge(raw_new, md_sdoaktif[['SDO Name', 'Status SDO']], on='SDO Name', how='left')

    # --- Join SPV, ASM, RSM ---
    spvrsm_clean = (
        spvrsm.rename(columns={'SDO': 'SDO Update', 'spv': 'SPV', 'asm': 'ASM', 'rsm': 'RSM'})
        .drop_duplicates('SDO Update')
    )
    spvrsm_cols = ['SDO Update'] + [c for c in ['SPV', 'ASM', 'RSM'] if c in spvrsm_clean.columns]
    raw_new['SDO Update'] = raw_new['SDO Update'].apply(cleaning_sdo)
    raw_new = pd.merge(raw_new, spvrsm_clean[spvrsm_cols], on='SDO Update', how='left')

    # --- Join MD_SKU ---
    raw_new = pd.merge(
        raw_new,
        md_sku[['SKU Code', 'Internal Code', 'Desc', 'Brand', 'PG', 'SUBPG']].drop_duplicates('SKU Code'),
        on='SKU Code', how='left'
    )
    raw_new['Internal Code'] = raw_new['Internal Code'].astype(str).str.strip().str.upper()

    # --- Join sub produk → PG 1 ---
    md_yee_join = (
        md_yee.rename(columns={'SHORT CODE': 'Internal Code', 'SUBPG2': 'PG 1'})
        [['Internal Code', 'PG 1']].drop_duplicates('Internal Code')
    )
    md_yee_join['Internal Code'] = md_yee_join['Internal Code'].astype(str).str.strip().str.upper()
    raw_new = pd.merge(raw_new, md_yee_join, on='Internal Code', how='left')

    # --- Join SWM Grouping → SUBPG1 ---
    md_swm_join = (
        md_swm.rename(columns={'Short Code': 'Internal Code'})
        [['Internal Code', 'SUBPG1']].drop_duplicates('Internal Code')
    )
    md_swm_join['Internal Code'] = md_swm_join['Internal Code'].astype(str).str.strip().str.upper()
    raw_new = pd.merge(raw_new, md_swm_join, on='Internal Code', how='left')
    raw_new['SUBPG1'] = raw_new['SUBPG1'].fillna('-')

    # --- Join Group ---
    md_group_join = (
        md_group.rename(columns={'CUSTOMER CODE1': 'Customer Code', 'GROUP': 'GROUP'})
        [['Customer Code', 'GROUP']].drop_duplicates('Customer Code')
    )
    md_group_join['Customer Code'] = md_group_join['Customer Code'].astype(str).apply(normalize_customer_code)
    raw_new = pd.merge(raw_new, md_group_join, on='Customer Code', how='left')

    # --- Kalkulasi tambahan ---
    raw_new['REG/FREE'] = 'REG'
    raw_new.loc[raw_new['Total (with VAT)'] < 1000, 'REG/FREE'] = 'FREE'

    raw_new['Year']        = raw_new['SAP Date'].dt.strftime('%Y')
    raw_new['Month']       = raw_new['SAP Date'].dt.strftime('%b')
    raw_new['Cycle']       = CYCLE
    raw_new['Dummy Cycle'] = DUMMY_CYCLE

    # --- REAL QTY ---
    raw_new['isi']      = raw_new['Internal Code'].apply(realqty)
    raw_new['REAL QTY'] = raw_new['QTY'] * raw_new['isi']
    raw_new = raw_new.drop(columns=['isi', '_row_idx'], errors='ignore')

    # --- Hapus Total Net = 0 ---
    before  = len(raw_new)
    raw_new = raw_new[raw_new['Total Net (Non VAT)'] != 0].reset_index(drop=True)
    print(f"  Hapus {before - len(raw_new)} baris Total Net (Non VAT) = 0.")

    # --- Hapus kolom duplikat dari merge ---
    raw_new = raw_new.loc[:, ~raw_new.columns.duplicated()]

    # --- Reorder kolom sesuai raw_old ---
    col_order = [
        'Customer Code', 'Customer Name', 'Address', 'City', 'Provinsi', 'Region', 'Reg2',
        'Type', 'TypeDummy', 'Chanel', 'Sales Group', 'PLANT', 'STORELOC',
        'SDO Name', 'SDO Update', 'Status SDO',
        'SKU Code', 'Internal Code', 'Desc', 'Brand', 'PG', 'SUBPG',
        'Total (with VAT)', 'REG/FREE', 'DR Number', 'SAP Date', 'QTY',
        'Total Net (Non VAT)', 'Year', 'Month', 'Cycle', 'Dummy Cycle', 'Status',
        'REAL QTY', 'PG 1', 'SUBPG1', 'GROUP', 'SPV', 'ASM', 'RSM'
    ]
    col_order = [c for c in col_order if c in raw_new.columns]
    raw_new   = raw_new[col_order]

    print(f"  raw_new selesai: {len(raw_new)} baris.")
    return raw_new


# ============================================================
# BAGIAN 6: Concat + update alamat + update nama MT
# ============================================================

def update_master_data(raw_old, raw_new):
    print("Merging raw_old + raw_new...")
    raw_old = raw_old.loc[:, ~raw_old.columns.duplicated()]
    raw_new = raw_new.loc[:, ~raw_new.columns.duplicated()]

    doc_revisi       = raw_new['DR Number'].unique()
    raw_old_filtered = raw_old[~raw_old['DR Number'].isin(doc_revisi)]

    df_final = pd.concat([raw_old_filtered, raw_new], ignore_index=True)
    print(f"  Total baris final: {len(df_final)} ({len(raw_old_filtered)} lama + {len(raw_new)} baru).")
    return df_final


def update_alamat(df_final, raw_new):
    print("Updating alamat dari raw_new (hanya baris yang belum punya alamat Google Maps)...")

    # Ambil alamat Google dari raw_new — hanya yang mengandung tanda '+'
    google_addr = (
        raw_new[['Customer Code', 'Address']]
        .dropna(subset=['Address'])
        .loc[raw_new['Address'].astype(str).str.contains(r'\+', regex=True)]
        .drop_duplicates('Customer Code', keep='last')
        .set_index('Customer Code')['Address']
    )

    if google_addr.empty:
        print("  Tidak ada alamat Google Maps di raw_new, skip update.")
        return df_final

    # Hanya update baris yang:
    # 1. Customer Code-nya ada di mapping alamat Google
    # 2. Alamatnya belum punya tanda '+' (belum Google Maps)
    has_customer  = df_final['Customer Code'].isin(google_addr.index)
    no_plus       = ~df_final['Address'].astype(str).str.contains(r'\+', regex=True, na=True)
    update_mask   = has_customer & no_plus

    df_final.loc[update_mask, 'Address'] = df_final.loc[update_mask, 'Customer Code'].map(google_addr)
    print(f"  {update_mask.sum()} baris alamat diupdate ({google_addr.shape[0]} customer dengan alamat Google).")
    return df_final


def update_desc_brand(raw_old, ms_dc):
    ms_dc = ms_dc.rename(columns={
        'Short Code':  'Internal Code',
        'Brand':       'Brand_dc',
        'PG':          'PG_dc',
        'SUB-PG':      'SUBPG_dc',
        'Description': 'Desc_dc'
    })
    raw_old['Internal Code'] = raw_old['Internal Code'].astype(str).str.strip().str.upper()
    ms_dc['Internal Code']   = ms_dc['Internal Code'].astype(str).str.strip().str.upper()

    cols_to_use = ['Internal Code'] + [c for c in ['Desc_dc', 'Brand_dc', 'PG_dc', 'SUBPG_dc']
                                       if c in ms_dc.columns]
    df_merge = raw_old.merge(ms_dc[cols_to_use], on='Internal Code', how='left')

    if 'Desc_dc'  in df_merge.columns: raw_old['Desc']  = df_merge['Desc_dc'].combine_first(raw_old.get('Desc'))
    if 'Brand_dc' in df_merge.columns: raw_old['Brand'] = df_merge['Brand_dc'].combine_first(raw_old.get('Brand'))
    if 'PG_dc'    in df_merge.columns: raw_old['PG']    = df_merge['PG_dc'].combine_first(raw_old.get('PG'))
    if 'SUBPG_dc' in df_merge.columns: raw_old['SUBPG'] = df_merge['SUBPG_dc'].combine_first(raw_old.get('SUBPG'))
    return raw_old


def update_customer_name_mt(df_final, md_group):
    print("Updating Customer Name (semua chanel)...")
    md_group_join = md_group.rename(columns={
        'CUSTOMER CODE1': 'Customer Code',
        'CUSTOMER NAME':  'Customer Name New'
    })
    md_group_join['Customer Code'] = md_group_join['Customer Code'].astype(str).apply(normalize_customer_code)
    name_map = md_group_join.drop_duplicates('Customer Code').set_index('Customer Code')['Customer Name New']

    in_group    = df_final['Customer Code'].isin(name_map.index)
    update_mask = in_group   # semua chanel, tidak dibatasi MODERN TRADE

    new_names = df_final.loc[update_mask, 'Customer Code'].map(name_map)
    changed   = (new_names.values != df_final.loc[update_mask, 'Customer Name'].values)

    actual_update = update_mask.copy()
    actual_update[update_mask] = changed
    df_final.loc[actual_update, 'Customer Name'] = df_final.loc[actual_update, 'Customer Code'].map(name_map)

    print(f"  {actual_update.sum()} baris Customer Name diupdate (semua chanel dengan perubahan nama).")
    return df_final


def cleaning_reg2(df_final):
    """
    Update kolom Reg2 DAN Region berdasarkan kombinasi RSM dan ASM setelah data tergabung.
    Reg2 = versi detail (CENTRAL JAVA 1, CENTRAL JAVA 2, PANTURA, dst)
    Region = versi broad (CENTRAL JAVA, JABODETABEK, WEST JAVA, SULAWESI)

    Format rules: (kondisi, Reg2_baru, Region_baru)
    - Jika Region_baru = None → hanya Reg2 yang diupdate, Region tidak disentuh
    """
    print("Cleaning Region & Reg2 berdasarkan RSM/ASM...")

    rsm = df_final['RSM'].str.upper().str.strip() if 'RSM' in df_final.columns else pd.Series('', index=df_final.index)
    asm = df_final['ASM'].str.upper().str.strip() if 'ASM' in df_final.columns else pd.Series('', index=df_final.index)
    rsm = rsm.fillna('')
    asm = asm.fillna('')

    # (kondisi, Reg2_baru, Region_baru)
    # Region_baru = None → hanya Reg2 yang diupdate
    rules = [
        (rsm == 'IMAM SHOVII',                                          'CENTRAL JAVA 3',   'CENTRAL JAVA'),
        ((rsm == 'ARIFIN SANTIKA') & (asm == 'AL JUSTIAN SAPUTRA'),    'JABODETABEK',      'JABODETABEK'),
        ((rsm == 'ARIFIN SANTIKA') & (asm == 'IMAM TAUFIQ'),           'WEST JAVA',        'WEST JAVA'),
        (rsm == 'ANDI HARIS',                                           'CENTRAL JAVA 1',   'CENTRAL JAVA'),
        (asm == 'IWAN SUSANTO',                                         'CENTRAL JAVA 2',   'CENTRAL JAVA'),
        (asm == 'JEKY TIRTA',                                           'SULAWESI',         'SULAWESI'),
        (asm == 'FREEKY JINO',                                          'PANTURA',          None),
    ]

    updated_reg2   = 0
    updated_region = 0
    for cond, val_reg2, val_region in rules:
        # Update Reg2
        mask_reg2 = cond & (df_final['Reg2'] != val_reg2)
        df_final.loc[mask_reg2, 'Reg2'] = val_reg2
        updated_reg2 += mask_reg2.sum()

        # Update Region hanya jika val_region tidak None
        if val_region is not None:
            mask_region = cond & (df_final['Region'] != val_region)
            df_final.loc[mask_region, 'Region'] = val_region
            updated_region += mask_region.sum()

    print(f"  {updated_reg2} baris Reg2 diupdate, {updated_region} baris Region diupdate.")
    return df_final

def update_desc_pricelist(df_final, pricelist):
    """
    Update kolom Desc di seluruh df_final dari pricelist terbaru.
    Hanya Internal Code yang ada di pricelist yang diupdate.
    Source: kolom Short Code (→ Internal Code) dan Description di sheet MEVAL.
    """
    print("Updating Desc dari pricelist...")

    pl = pricelist.rename(columns={'Short Code': 'Internal Code', 'Description': 'Desc_pl'})
    pl['Internal Code'] = pl['Internal Code'].astype(str).str.strip().str.upper()
    pl = pl[['Internal Code', 'Desc_pl']].dropna(subset=['Desc_pl']).drop_duplicates('Internal Code')

    df_final['Internal Code'] = df_final['Internal Code'].astype(str).str.strip().str.upper()
    df_final = df_final.merge(pl, on='Internal Code', how='left')

    # Hanya update yang ada di pricelist (Desc_pl tidak NaN)
    has_pl = df_final['Desc_pl'].notna()
    df_final.loc[has_pl, 'Desc'] = df_final.loc[has_pl, 'Desc_pl']
    df_final = df_final.drop(columns=['Desc_pl'])

    print(f"  {has_pl.sum()} baris Desc diupdate dari pricelist ({pl.shape[0]} produk unik).")
    return df_final


def update_sdo_raw_old(raw_old, sdo_update):
    """
    Update kolom SDO Update di raw_old berdasarkan file konfirmasi SDO.
    Hanya Customer Code yang ada di sdo_update yang diupdate.
    """
    print("Updating SDO Update di raw_old...")

    sdo_map = (
        sdo_update[['Customer Code', 'SDO Update']]
        .dropna(subset=['SDO Update'])
        .drop_duplicates('Customer Code', keep='last')
        .set_index('Customer Code')['SDO Update']
    )

    if 'SDO Update' not in raw_old.columns:
        raw_old['SDO Update'] = None

    raw_old['Customer Code'] = raw_old['Customer Code'].astype(str).apply(normalize_customer_code)
    mask = raw_old['Customer Code'].isin(sdo_map.index)
    raw_old.loc[mask, 'SDO Update'] = raw_old.loc[mask, 'Customer Code'].map(sdo_map)

    print(f"  {mask.sum()} baris SDO Update diupdate di raw_old ({sdo_map.shape[0]} customer unik).")
    return raw_old


def update_spv_asm_rsm(df_final, spvrsm):
    """
    Update kolom SPV, ASM, RSM di seluruh df_final (raw_old + raw_new)
    berdasarkan SDO Update menggunakan file spv rsm Sheet2.
    """
    print("Updating SPV, ASM, RSM dari spvrsm...")

    spvrsm_clean = (
        spvrsm.rename(columns={'SDO': 'SDO Update', 'spv': 'SPV', 'asm': 'ASM', 'rsm': 'RSM'})
        .drop_duplicates('SDO Update')
    )
    cols_avail = ['SDO Update'] + [c for c in ['SPV', 'ASM', 'RSM'] if c in spvrsm_clean.columns]
    spvrsm_clean = spvrsm_clean[cols_avail].set_index('SDO Update')

    df_final['SDO Update'] = df_final['SDO Update'].astype(str).str.strip().apply(cleaning_sdo)
    mask = df_final['SDO Update'].isin(spvrsm_clean.index)

    for col in ['SPV', 'ASM', 'RSM']:
        if col in spvrsm_clean.columns:
            if col not in df_final.columns:
                df_final[col] = None
            df_final.loc[mask, col] = df_final.loc[mask, 'SDO Update'].map(spvrsm_clean[col])

    print(f"  {mask.sum()} baris SPV/ASM/RSM diupdate ({spvrsm_clean.shape[0]} SDO unik di mapping).")
    return df_final


def update_status_sdo(df_final, sdo_aktif):
    """
    Update kolom Status SDO di seluruh df_final berdasarkan file SDO_aktif.
    Kunci lookup: SDO Name (di-normalize dengan cleaning_sdo sebelum lookup).

    Kolom SDO_aktif: SDO Name, STATUS SDO, AREA
    Hanya SDO Name yang ada di file aktif yang Status SDO-nya diupdate.
    Yang tidak ditemukan dibiarkan (tidak dihapus/di-null).
    """
    print("Updating Status SDO dari SDO_aktif...")

    # Normalize SDO Name di master aktif
    sdo_aktif = sdo_aktif.copy()
    sdo_aktif['SDO Name'] = sdo_aktif['SDO Name'].astype(str).str.upper().str.strip().apply(cleaning_sdo)
    sdo_aktif = sdo_aktif.dropna(subset=['SDO Name']).drop_duplicates('SDO Name', keep='last')

    # Mapping SDO Name → STATUS SDO
    status_map = sdo_aktif.set_index('SDO Name')['STATUS SDO']

    # Normalize SDO Name di df_final sebelum lookup
    if 'SDO Name' not in df_final.columns:
        print("  WARN: kolom 'SDO Name' tidak ada di df_final, skip.")
        return df_final

    sdo_name_clean = df_final['SDO Name'].astype(str).str.upper().str.strip().apply(cleaning_sdo)

    # Tandai baris yang SDO Name-nya ada di master aktif
    in_master = sdo_name_clean.isin(status_map.index)

    # Update Status SDO
    if 'Status SDO' not in df_final.columns:
        df_final['Status SDO'] = None

    df_final.loc[in_master, 'Status SDO'] = sdo_name_clean[in_master].map(status_map).values

    n_updated  = in_master.sum()
    n_notfound = (~in_master).sum()
    n_aktif    = (df_final.loc[in_master, 'Status SDO'] == 'AKTIF').sum()
    n_taktif   = (df_final.loc[in_master, 'Status SDO'] == 'TIDAK AKTIF').sum()

    print(f"  {n_updated} baris Status SDO diupdate "
          f"({n_aktif} AKTIF, {n_taktif} TIDAK AKTIF).")
    print(f"  {n_notfound} baris SDO Name tidak ditemukan di master aktif (dibiarkan).")
    return df_final


def update_product_focus_category(df_final, product_data):
    """
    Tambah / update kolom PRODUCT FOCUS dan CATEGORY di df_final
    berdasarkan lookup Internal Code ke Product_Data.xlsx.
    Kolom source: Internal Code, CATEGORY, PRODUCT FOCUS.
    - Internal Code yang tidak ada di product_data → kolom diisi NaN (tidak diubah)
    - Selalu overwrite dengan data terbaru dari product_data
    """
    print("Updating PRODUCT FOCUS & CATEGORY dari Product Data...")

    prod = product_data[['Internal Code', 'CATEGORY', 'PRODUCT FOCUS']].copy()
    prod['Internal Code'] = prod['Internal Code'].astype(str).str.strip().str.upper()
    prod = prod.dropna(subset=['Internal Code']).drop_duplicates('Internal Code', keep='last')

    df_final['Internal Code'] = df_final['Internal Code'].astype(str).str.strip().str.upper()

    # Merge — keep all rows, kolom baru muncul dengan suffix _pd jika sudah ada
    df_final = df_final.merge(
        prod.rename(columns={'CATEGORY': 'CATEGORY_pd', 'PRODUCT FOCUS': 'PRODUCT FOCUS_pd'}),
        on='Internal Code', how='left'
    )

    # Assign ke kolom final (overwrite jika sudah ada, buat baru jika belum)
    df_final['CATEGORY']      = df_final['CATEGORY_pd']
    df_final['PRODUCT FOCUS'] = df_final['PRODUCT FOCUS_pd']
    df_final = df_final.drop(columns=['CATEGORY_pd', 'PRODUCT FOCUS_pd'])

    n_filled = df_final['PRODUCT FOCUS'].notna().sum()
    n_empty  = df_final['PRODUCT FOCUS'].isna().sum()
    print(f"  {n_filled} baris terisi, {n_empty} baris Internal Code tidak ditemukan di product data.")
    return df_final


def check_chanel_mismatch(raw_new, md_toko):
    """
    Bandingkan kolom Chanel di raw_new (hasil process_new_month)
    dengan Chanel di MD_TOKO per Customer Code.
    Tampilkan hanya yang tidak sesuai.

    Return:
        DataFrame berisi Customer Code yang chanelnya berbeda antara
        raw_new dan MD_TOKO, dengan kolom:
        Customer Code, Customer Name, Chanel raw_new, Chanel MD_TOKO
    """
    # Ambil Chanel dari raw_new — satu baris per Customer Code
    raw_unique = (
        raw_new[['Customer Code', 'Customer Name', 'Chanel']]
        .drop_duplicates('Customer Code')
        .rename(columns={'Chanel': 'Chanel_raw_new', 'Customer Name': 'Nama raw_new'})
    )
    raw_unique['Customer Code']   = raw_unique['Customer Code'].astype(str).apply(normalize_customer_code)
    raw_unique['Chanel_raw_new']  = raw_unique['Chanel_raw_new'].astype(str).str.strip().str.upper()

    # Ambil Chanel dari MD_TOKO
    toko = (
        md_toko[['Customer Code', 'Customer Name', 'Chanel']]
        .drop_duplicates('Customer Code')
        .rename(columns={'Chanel': 'Chanel_MD_TOKO', 'Customer Name': 'Nama MD_TOKO'})
    )
    toko['Chanel_MD_TOKO'] = toko['Chanel_MD_TOKO'].astype(str).str.strip().str.upper()

    # Merge
    merged = pd.merge(raw_unique, toko, on='Customer Code', how='left')

    # Tandai yang belum ada di MD_TOKO
    merged['Chanel_MD_TOKO'] = merged['Chanel_MD_TOKO'].fillna('(BELUM ADA DI MD_TOKO)')

    # Filter hanya yang berbeda
    mismatch = merged[
        merged['Chanel_raw_new'] != merged['Chanel_MD_TOKO']
    ].copy().reset_index(drop=True)

    mismatch = mismatch[[
        'Customer Code', 'Nama raw_new', 'Nama MD_TOKO',
        'Chanel_raw_new', 'Chanel_MD_TOKO'
    ]]
    mismatch.columns = [
        'Customer Code', 'Nama di raw_new', 'Nama di MD_TOKO',
        'Chanel raw_new', 'Chanel MD_TOKO'
    ]

    return mismatch


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 55)
    print(f"ETL Sell-In | Cycle: {CYCLE} | {DUMMY_CYCLE}")
    print("=" * 55)

    # 1. Load master utama
    md_toko, md_sdo, md_sku, spvrsm, sap = load_master_data()

    # 2. Load master tambahan
    print("Loading sub produk...")
    md_yee = pd.read_excel(PATH_MD_SUBPRODUK, sheet_name='PG')

    print("Loading SWM Grouping...")
    md_swm = pd.read_excel(PATH_SWM, sheet_name='modern_classic')
    # Pastikan kolom Short Code tersedia untuk rename di process_new_month
    if 'Short Code' not in md_swm.columns and 'SHORT CODE' in md_swm.columns:
        md_swm = md_swm.rename(columns={'SHORT CODE': 'Short Code'})

    print("Loading Pricelist...")
    pricelist = pd.read_excel(PATH_PRICELIST)

    print("Loading SDO Aktif...")
    sdo_aktif = pd.read_excel(PATH_SDO_AKTIF)

    print("Loading Product Data...")
    product_data = pd.read_excel(PATH_PRODUCT)

    print("Loading Group...")
    md_group = pd.read_excel(PATH_GROUP)
    md_group['CUSTOMER CODE1'] = md_group['CUSTOMER CODE1'].astype(str).apply(normalize_customer_code)

    # 3. Load SDO Update
    print("Loading SDO Update...")
    sdo_update = pd.read_excel(PATH_SDO_UPDATE, sheet_name='Sheet1')
    sdo_update['Customer Code'] = sdo_update['Customer Code'].astype(str).apply(normalize_customer_code)

    # 4. Load MTD YTD — sekali, index bersih
    print("Loading MTD YTD...")
    df_mtd = pd.read_excel(PATH_MTD_YTD, sheet_name=MTD_SHEET)
    df_mtd = df_mtd.reset_index(drop=True)
    df_mtd['CUSTOMER#']       = df_mtd['CUSTOMER#'].astype(str).apply(normalize_customer_code)
    df_mtd['ACTUAL PGI DATE'] = parse_sap_date(df_mtd['ACTUAL PGI DATE'])
    df_mtd['SDO']             = df_mtd['SDO'].astype(str).apply(cleaning_sdo)
    df_mtd.loc[df_mtd['CUST TYPE'] == 'PROJECT',['CUST TYPE']] = 'PROJECTS'
    
    df_unique = df_mtd.drop_duplicates('CUSTOMER#').copy()
    df_unique = df_unique.rename(columns={'CUSTOMER#': 'Customer Code'})

    # 5. Update master
    md_toko        = update_md_toko(md_toko, df_unique, df_unique, sap, spvrsm)
    md_sdo_updated = update_md_sdo(md_sdo, df_unique, sdo_update)

    # 6. Proses data bulan ini
    raw_new = process_new_month(
        md_toko, md_sdo, md_sdo_updated, md_sku,
        spvrsm, md_yee, md_group, md_swm, df_mtd
    )

    raw_new.to_excel(OUTPUT_RAW_NEW, index=False)
    print(f"\nraw_new disimpan → {OUTPUT_RAW_NEW}")

    # QC: bandingkan Chanel MTD YTD vs MD_TOKO, simpan ke sheet kedua
    print("Checking chanel mismatch raw_new vs MD_TOKO...")
    mismatch = check_chanel_mismatch(raw_new, md_toko)
    if mismatch.empty:
        print("  Semua chanel sesuai, tidak ada mismatch.")
    else:
        print(f"  Ditemukan {len(mismatch)} Customer Code dengan chanel tidak sesuai.")
        with pd.ExcelWriter(OUTPUT_RAW_NEW, engine='openpyxl', mode='a') as writer:
            mismatch.to_excel(writer, sheet_name='Chanel Mismatch', index=False)
        print(f"  Sheet 'Chanel Mismatch' ditambahkan ke {OUTPUT_RAW_NEW}")

    # 7. Load raw_old + bersihkan
    print("\nLoading raw_old...")
    raw_old = pd.read_excel(PATH_RAW_OLD, sheet_name='Sheet1')
    ms_dc   = pd.read_excel(PATH_MS_DC)
    raw_old = update_desc_brand(raw_old, ms_dc)
    raw_old = update_sdo_raw_old(raw_old, sdo_update)

    before  = len(raw_old)
    raw_old = raw_old[raw_old['Total Net (Non VAT)'] != 0].reset_index(drop=True)
    print(f"  Hapus {before - len(raw_old)} baris Total Net (Non VAT) = 0 dari raw_old.")

    # 8. Gabung + update
    df_final = update_master_data(raw_old, raw_new)
    df_final = update_alamat(df_final, raw_new)
    df_final = update_customer_name_mt(df_final, md_group)

    # Cleaning SDO Name di df_final sebelum lookup Status SDO
    print("Cleaning SDO Name di df_final...")
    if 'SDO Name' in df_final.columns:
        before_clean = df_final['SDO Name'].nunique()
        df_final['SDO Name'] = df_final['SDO Name'].astype(str).str.upper().str.strip().apply(cleaning_sdo)
        after_clean = df_final['SDO Name'].nunique()
        print(f"  SDO Name unik: {before_clean} → {after_clean} setelah normalisasi.")

    # Update Status SDO dari master SDO aktif
    df_final = update_status_sdo(df_final, sdo_aktif)

    # Update SDO Update di seluruh df_final dari file konfirmasi
    print("Updating SDO Update di df_final...")
    sdo_map = (
        sdo_update[['Customer Code', 'SDO Update']]
        .dropna(subset=['SDO Update'])
        .drop_duplicates('Customer Code', keep='last')
        .set_index('Customer Code')['SDO Update']
    )
    df_final['Customer Code'] = df_final['Customer Code'].astype(str).apply(normalize_customer_code)
    mask = df_final['Customer Code'].isin(sdo_map.index)
    df_final.loc[mask, 'SDO Update'] = df_final.loc[mask, 'Customer Code'].map(sdo_map)
    print(f"  {mask.sum()} baris SDO Update diupdate di df_final.")

    df_final = update_spv_asm_rsm(df_final, spvrsm)
    df_final = cleaning_reg2(df_final)
    df_final = update_desc_pricelist(df_final, pricelist)

    # Uppercase Desc dan SUBPG sesuai standar raw_old
    print("Uppercasing kolom Desc, SUBPG dan PG...")
    if 'Desc' in df_final.columns:
        df_final['Desc'] = df_final['Desc'].astype(str).str.upper().replace('NAN', None)
    if 'SUBPG' in df_final.columns:
        df_final['SUBPG'] = df_final['SUBPG'].astype(str).str.upper().replace('NAN', None)
    if 'PG' in df_final.columns:
        df_final['PG'] = df_final['PG'].astype(str).str.upper().replace('NAN',None)

    # Tambah kolom PRODUCT FOCUS dan CATEGORY dari Product Data (paling akhir)
    df_final = update_product_focus_category(df_final, product_data)

    # Hapus baris Chanel = DIRECT MDI sebelum simpan
    before = len(df_final)
    df_final = df_final[
        df_final['Chanel'].astype(str).str.upper().str.strip() != 'DIRECT MDI'
    ].reset_index(drop=True)
    print(f"Hapus {before - len(df_final)} baris Chanel DIRECT MDI.")

    # 9. Simpan
    df_final.to_excel(OUTPUT_FINAL, index=False)
    print(f"\nOutput final disimpan → {OUTPUT_FINAL}")
    print(f"Total baris: {len(df_final)}")
    print("\nSelesai!")


if __name__ == "__main__":
    main()