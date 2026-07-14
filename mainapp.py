import streamlit as st
import geopandas as gpd
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
import fiona
import io
import datetime
import re
import tempfile

# Enable KML support
fiona.drvsupport.supported_drivers['KML'] = 'rw'
fiona.drvsupport.supported_drivers['LIBKML'] = 'rw'

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="County-Agnostic Territory Analyzer", layout="wide")
st.title("Congregation Territory Address Analyzer")

if 'mappings' not in st.session_state: st.session_state['mappings'] = None
if 'excluded_values' not in st.session_state: st.session_state['excluded_values'] = None
if 'excel_data' not in st.session_state: st.session_state['excel_data'] = None

# --- HELPERS ---
def natural_keys(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(text))]

def get_base_address(address_str):
    regex = r'(?i)(apt|unit|ste|#|suite|lot)\s*[a-zA-Z0-9-]*$'
    return re.sub(regex, '', str(address_str)).strip().strip(',')

# --- 2. UPLOAD & MAPPING ---
st.header("Step 1: Upload & Map")
uploaded_shapefile = st.file_uploader("Upload County Shapefile (.zip)", type=["zip"])
uploaded_kml = st.file_uploader("Upload Territory KML File", type=["kml"])

if uploaded_shapefile and not st.session_state['mappings']:
    tfile = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
    tfile.write(uploaded_shapefile.read())
    tfile.close()
    try:
        raw_data = gpd.read_file(f"zip://{tfile.name}")
        if not hasattr(raw_data, 'columns'):
            st.error("Error: Shapefile contains geometry but no data. Ensure .dbf is present.")
            st.stop()
        
        cols = raw_data.columns.tolist()
        method = st.radio("Address format:", ["Single 'Full Address' column", "Separate columns (House #, Street, etc.)"])
        col_map = {'method': method}
        if method == "Single 'Full Address' column": 
            col_map['FullAddress'] = st.selectbox("Select 'Full Address'", cols)
        else:
            for f in ['HouseNo', 'HouseSx', 'Dir', 'Street', 'StType', 'Unit', 'Muni', 'Zip_Code']:
                col_map[f] = st.selectbox(f"Select {f}", cols)
        col_map['Status'] = st.selectbox("Select 'Status' column", cols)
        
        if st.button("Confirm Mapping"):
            st.session_state['mappings'] = col_map
            st.session_state['gdf_path'] = tfile.name
            st.rerun()
    except Exception as e: st.error(f"Upload Error: {e}")

if st.session_state['mappings'] and not st.session_state['excluded_values']:
    raw_data = gpd.read_file(f"zip://{st.session_state['gdf_path']}")
    st.session_state['excluded_values'] = st.multiselect("Excluded Statuses", raw_data[st.session_state['mappings']['Status']].unique().tolist())

# --- 3. ANALYSIS ENGINE ---
def generate_excel_report(joined_gdf, min_goal, max_goal, cong_name):
    output = io.BytesIO()
    joined_gdf['Base_Address'] = joined_gdf['Mailable_Address'].apply(get_base_address)
    excluded_gdf = joined_gdf[joined_gdf['Internal_Status'].isin(st.session_state['excluded_values'])].copy()
    valid_gdf = joined_gdf[~joined_gdf['Internal_Status'].isin(st.session_state['excluded_values'])].copy()
    
    unique_territories = sorted(valid_gdf['Territory_Name'].unique().astype(str), key=natural_keys)
    valid_gdf['Territory_Name'] = pd.Categorical(valid_gdf['Territory_Name'], categories=unique_territories, ordered=True)
    counts = valid_gdf.groupby('Territory_Name', observed=True).size().reset_index(name='Total_Addresses')
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Dashboard, Counts, Address List, etc (Same logic as Milwaukee version)
        pd.DataFrame([["Territory Analysis Output"]]).to_excel(writer, sheet_name="Dashboard", index=False)
        counts.to_excel(writer, sheet_name="Counts", index=False)
        valid_gdf.to_excel(writer, sheet_name="Address List", index=False)
    output.seek(0)
    return output

# --- 4. EXECUTION ---
cong_name = st.text_input("Congregation Name", "Congregation")
goal = st.selectbox("Goal Range", ["25-50", "50-75", "75-100", "100-125", "125-150"])
min_g, max_g = [int(x) for x in goal.split("-")]

if st.session_state['mappings'] and st.session_state['excluded_values'] is not None and uploaded_kml:
    if st.button("Generate Territory Analysis"):
        try:
            # Load
            raw_gdf = gpd.read_file(f"zip://{st.session_state['gdf_path']}")
            kml_gdf = gpd.read_file(uploaded_kml, driver="KML").make_valid()
            
            # Projection Hardening (Crucial Step)
            if raw_gdf.crs != kml_gdf.crs:
                raw_gdf = raw_gdf.to_crs(kml_gdf.crs)
            
            # Standardize Addresses
            mappings = st.session_state['mappings']
            if mappings['method'] == "Single 'Full Address' column":
                raw_gdf['Mailable_Address'] = raw_gdf[mappings['FullAddress']].astype(str)
            else:
                raw_gdf['Mailable_Address'] = (raw_gdf[mappings['HouseNo']].astype(str) + " " + raw_gdf[mappings['Street']].astype(str)).str.strip()
            
            raw_gdf['Internal_Status'] = raw_gdf[mappings['Status']].astype(str)
            
            # Spatial Join
            joined = gpd.sjoin(raw_gdf, kml_gdf.rename(columns={'geometry': 'geometry_terr'}).set_geometry('geometry_terr'), how="inner", predicate="within")
            if len(joined) == 0:
                st.error("No addresses found within the boundaries. Check if your Coordinate Systems match.")
                st.stop()

            # Territory Naming
            name_col = next((c for c in ['Name', 'name', 'Title'] if c in kml_gdf.columns), None)
            joined['Territory_Name'] = joined[name_col] if name_col else "Territory_" + joined.index.astype(str)
            
            st.session_state['excel_data'] = generate_excel_report(joined, min_g, max_g, cong_name)
            st.success("Analysis Complete!")
            st.rerun()
        except Exception as e:
            st.error(f"Processing Error: {e}")

if st.session_state['excel_data']:
    st.download_button("⬇️ Download Excel Analysis", st.session_state['excel_data'], f"{cong_name}_Analysis.xlsx")
