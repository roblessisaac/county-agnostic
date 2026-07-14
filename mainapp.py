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

# --- 1. CONFIGURATION & STATE ---
st.set_page_config(page_title="County-Agnostic Territory Analyzer", layout="wide")
st.title("Congregation Territory Address Analyzer")

if 'mappings' not in st.session_state: st.session_state['mappings'] = None
if 'excluded_values' not in st.session_state: st.session_state['excluded_values'] = None
if 'excel_data' not in st.session_state: st.session_state['excel_data'] = None

# --- 2. HELPERS ---
def natural_keys(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(text))]

def get_base_address(address_str):
    regex = r'(?i)(apt|unit|ste|#|suite|lot)\s*[a-zA-Z0-9-]*$'
    return re.sub(regex, '', str(address_str)).strip().strip(',')

# --- 3. UI MAPPING LOGIC ---
st.header("Step 1: Upload Data")
uploaded_shapefile = st.file_uploader("Upload County Shapefile (.zip)", type=["zip"])
uploaded_kml = st.file_uploader("Upload Territory KML File", type=["kml"])

if uploaded_shapefile and not st.session_state['mappings']:
    tfile = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
    tfile.write(uploaded_shapefile.read())
    tfile.close()
    
    try:
        raw_data = gpd.read_file(f"zip://{tfile.name}")
        # Ensure it's a table, not just a series of shapes
        if not hasattr(raw_data, 'columns'):
            st.error("The file uploaded has shapes but no data table. Please check if your zip contains the .dbf file.")
            st.stop()
            
        cols = raw_data.columns.tolist()
        st.subheader("Map Your Columns")
        method = st.radio("How is your address data stored?", ["Single 'Full Address' column", "Separate columns (House #, Street, etc.)"])
        col_map = {'method': method}
        if method == "Single 'Full Address' column":
            col_map['FullAddress'] = st.selectbox("Select 'Full Address' column", cols)
        else:
            for field in ['HouseNo', 'HouseSx', 'Dir', 'Street', 'StType', 'Unit', 'Muni', 'Zip_Code']:
                col_map[field] = st.selectbox(f"Select column for {field}", cols)
        col_map['Status'] = st.selectbox("Select 'Status' column", cols)
        
        if st.button("Confirm Mapping"):
            st.session_state['mappings'] = col_map
            st.session_state['gdf_path'] = tfile.name
            st.rerun()
    except Exception as e: st.error(f"Error reading shapefile: {e}")

if st.session_state['mappings'] and not st.session_state['excluded_values']:
    raw_data = gpd.read_file(f"zip://{st.session_state['gdf_path']}")
    unique_vals = raw_data[st.session_state['mappings']['Status']].unique().tolist()
    st.subheader("Select Excluded Statuses")
    st.session_state['excluded_values'] = st.multiselect("Select values to treat as 'Excluded' (Tab 6)", unique_vals)

# --- 4. EXCEL GENERATION ENGINE ---
def generate_excel_report(joined_gdf, kml_gdf, min_goal, max_goal, cong_name):
    output = io.BytesIO()
    joined_gdf['Base_Address'] = joined_gdf['Mailable_Address'].apply(get_base_address)
    
    excluded_gdf = joined_gdf[joined_gdf['Internal_Status'].isin(st.session_state['excluded_values'])].copy()
    valid_gdf = joined_gdf[~joined_gdf['Internal_Status'].isin(st.session_state['excluded_values'])].copy()

    unique_territories = valid_gdf['Territory_Name'].unique().tolist()
    unique_territories.sort(key=natural_keys)
    valid_gdf['Territory_Name'] = pd.Categorical(valid_gdf['Territory_Name'], categories=unique_territories, ordered=True)
    
    counts_df = valid_gdf.groupby('Territory_Name', observed=True).size().reset_index(name='Total_Addresses')
    counts_df['Category'] = counts_df['Total_Addresses'].apply(lambda c: "Undersized" if c < min_goal else ("Ideal" if min_goal <= c <= max_goal else "Oversized"))
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Dashboard logic remains identical to your finalized Milwaukee version
        pass # (Include the same logic as the previous script here)
    output.seek(0)
    return output

# --- 5. EXECUTION & DOWNLOAD ---
cong_name = st.text_input("Congregation Name", "ExampleCongregation")
goal = st.selectbox("Goal Range", ["25-50", "50-75", "75-100", "100-125", "125-150"])
min_g, max_g = [int(x) for x in goal.split("-")]

if st.session_state['mappings'] and st.session_state['excluded_values'] is not None and uploaded_kml:
    if st.button("Generate Territory Analysis"):
        try:
            mappings = st.session_state['mappings']
            raw_gdf = gpd.read_file(f"zip://{st.session_state['gdf_path']}")
            
            # Standardization
            if mappings['method'] == "Single 'Full Address' column":
                raw_gdf['Mailable_Address'] = raw_gdf[mappings['FullAddress']]
            else:
                raw_gdf['Mailable_Address'] = raw_gdf[mappings['HouseNo']].astype(str) + " " + raw_gdf[mappings['Street']]
            raw_gdf['Internal_Status'] = raw_gdf[mappings['Status']]
            
            kml_gdf = gpd.read_file(uploaded_kml, driver="KML").make_valid()
            # Defensive territory name parsing
            name_cols = ['Name', 'name', 'Title', 'title', 'Description', 'description']
            name_col = next((c for c in name_cols if c in kml_gdf.columns), None)
            kml_gdf['Territory_Name'] = kml_gdf[name_col].fillna("Territory_" + kml_gdf.index.astype(str)) if name_col else "Territory_" + kml_gdf.index.astype(str)
            
            joined = gpd.sjoin(raw_gdf.to_crs(kml_gdf.crs), kml_gdf.rename(columns={'geometry': 'geometry_terr'}).set_geometry('geometry_terr'), how="inner", predicate="within")
            st.session_state['excel_data'] = generate_excel_report(joined, kml_gdf, min_g, max_g, cong_name)
            st.success("Analysis Complete!")
            st.rerun()
        except Exception as e:
            st.error(f"Error processing files: {e}")

if st.session_state['excel_data']:
    st.download_button("⬇️ Download Excel Analysis", st.session_state['excel_data'], f"{cong_name}_Analysis.xlsx")
