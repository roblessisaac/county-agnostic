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
import os

# Enable KML support
fiona.drvsupport.supported_drivers['KML'] = 'rw'
fiona.drvsupport.supported_drivers['LIBKML'] = 'rw'

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="County-Agnostic Territory Analyzer", layout="wide")
st.title("Congregation Territory Analysis Engine")

# --- 2. HELPERS & UTILITIES ---
def natural_keys(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(text))]

def get_base_address(address_str):
    # Regex to clean units for grouping: matches 'Apt 1', 'Unit 2', '#3', etc.
    regex = r'(?i)(apt|unit|ste|#|suite|lot)\s*[a-zA-Z0-9-]*$'
    return re.sub(regex, '', str(address_str)).strip().strip(',')

# --- 3. MAPPING UI ---
st.header("Step 1: Upload Data")
uploaded_shapefile = st.file_uploader("Upload County Shapefile (.zip)", type=["zip"])
uploaded_kml = st.file_uploader("Upload Territory KML File", type=["kml"])

if 'mappings' not in st.session_state: st.session_state['mappings'] = None
if 'excluded_values' not in st.session_state: st.session_state['excluded_values'] = None
if 'excel_data' not in st.session_state: st.session_state['excel_data'] = None

if uploaded_shapefile and not st.session_state['mappings']:
    tfile = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
    tfile.write(uploaded_shapefile.read())
    tfile.close()
    try:
        raw_data = gpd.read_file(f"zip://{tfile.name}")
        if not hasattr(raw_data, 'columns'):
            st.error("Error: Shapefile contains geometry but no data table. Check if .dbf is included.")
            st.stop()
        
        cols = raw_data.columns.tolist()
        method = st.radio("Address format:", ["Single 'Full Address' column", "Separate columns (House #, Street, etc.)"])
        col_map = {'method': method}
        if method == "Single 'Full Address' column": 
            col_map['FullAddress'] = st.selectbox("Select 'Full Address' column", cols)
        else:
            for f in ['HouseNo', 'HouseSx', 'Dir', 'Street', 'StType', 'Unit', 'Muni', 'Zip_Code']:
                col_map[f] = st.selectbox(f"Select {f} column", cols)
        col_map['Status'] = st.selectbox("Select 'Status' column", cols)
        
        if st.button("Confirm Mapping"):
            st.session_state['mappings'] = col_map
            st.session_state['gdf_path'] = tfile.name
            st.rerun()
    except Exception as e: st.error(f"Upload Error: {e}")

if st.session_state['mappings'] and not st.session_state['excluded_values']:
    raw_data = gpd.read_file(f"zip://{st.session_state['gdf_path']}")
    st.session_state['excluded_values'] = st.multiselect("Select Statuses to Exclude", raw_data[st.session_state['mappings']['Status']].unique().tolist())

# --- 4. ENGINE: EXCEL GENERATION ---
def generate_excel_report(joined_gdf, cong_name, min_goal, max_goal):
    output = io.BytesIO()
    joined_gdf['Base_Address'] = joined_gdf['Mailable_Address'].apply(get_base_address)
    
    excluded_gdf = joined_gdf[joined_gdf['Internal_Status'].isin(st.session_state['excluded_values'])].copy()
    valid_gdf = joined_gdf[~joined_gdf['Internal_Status'].isin(st.session_state['excluded_values'])].copy()
    
    unique_territories = sorted(valid_gdf['Territory_Name'].unique().astype(str), key=natural_keys)
    valid_gdf['Territory_Name'] = pd.Categorical(valid_gdf['Territory_Name'], categories=unique_territories, ordered=True)
    counts = valid_gdf.groupby('Territory_Name', observed=True).size().reset_index(name='Total_Addresses')
    counts['Category'] = counts['Total_Addresses'].apply(lambda c: "Undersized" if c < min_goal else ("Ideal" if min_goal <= c <= max_goal else "Oversized"))
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Dashboard Sheet
        dashboard_data = [
            [f"Territory Analysis: {cong_name}"],
            [f"Generated {datetime.datetime.now().strftime('%B %Y')} by Territory Analysis Engine."],
            [""],
            [f"Total Territories: {len(counts)}"],
            [f"Total Valid Addresses: {counts['Total_Addresses'].sum()}"],
            [f"Excluded Addresses: {len(excluded_gdf)}"],
            [""],
            ["Goal Range:", f"{min_goal}-{max_goal}"]
        ]
        pd.DataFrame(dashboard_data).to_excel(writer, sheet_name="Dashboard", index=False, header=False)
        
        ws_dash = writer.sheets['Dashboard']
        ws_dash['A1'].font = Font(size=20, bold=True)
        
        # Counts Sheet
        counts.to_excel(writer, sheet_name="Counts", index=False)
        ws_counts = writer.sheets['Counts']
        ws_counts.freeze_panes = 'A2'
        
        # Address List Sheet
        valid_gdf[['Territory_Name', 'Mailable_Address']].to_excel(writer, sheet_name="Address List", index=False)
        ws_addr = writer.sheets['Address List']
        ws_addr.freeze_panes = 'A2'
        ws_addr.column_dimensions['B'].width = 55
        
        # Apartments Sheet (5+ units)
        apt = valid_gdf.groupby(['Territory_Name', 'Base_Address'], observed=True).size().reset_index(name='Total Units')
        apt[apt['Total Units'] >= 5].to_excel(writer, sheet_name="Apartments", index=False)
        
        # Audit Sheet
        excluded_gdf[['Territory_Name', 'Mailable_Address', 'Internal_Status']].to_excel(writer, sheet_name="Excluded Audit", index=False)
        ws_audit = writer.sheets['Excluded Audit']
        ws_audit.column_dimensions['B'].width = 55

    output.seek(0)
    return output

# --- 5. EXECUTION & DOWNLOAD ---
cong_name = st.text_input("Congregation Name", "Congregation")
goal = st.selectbox("Goal Range", ["25-50", "50-75", "75-100", "100-125", "125-150"])
min_g, max_g = [int(x) for x in goal.split("-")]

if st.session_state['mappings'] and st.session_state['excluded_values'] is not None and uploaded_kml:
    if st.button("Generate Territory Analysis"):
        try:
            raw_gdf = gpd.read_file(f"zip://{st.session_state['gdf_path']}")
            kml_gdf = gpd.read_file(uploaded_kml, driver="KML").make_valid()
            
            # Projection Hardening
            if raw_gdf.crs != kml_gdf.crs: raw_gdf = raw_gdf.to_crs(kml_gdf.crs)
            
            # Standardization Engine
            mappings = st.session_state['mappings']
            if mappings['method'] == "Single 'Full Address' column":
                raw_gdf['Mailable_Address'] = raw_gdf[mappings['FullAddress']].astype(str)
            else:
                raw_gdf['Mailable_Address'] = (raw_gdf[mappings['HouseNo']].astype(str) + " " + raw_gdf[mappings['Street']].astype(str)).str.strip()
            
            raw_gdf['Internal_Status'] = raw_gdf[mappings['Status']].astype(str)
            
            # Join
            joined = gpd.sjoin(raw_gdf, kml_gdf.rename(columns={'geometry': 'geometry_terr'}).set_geometry('geometry_terr'), how="inner", predicate="within")
            
            # Territory Name
            name_col = next((c for c in ['Name', 'name', 'Title', 'title'] if c in kml_gdf.columns), None)
            joined['Territory_Name'] = joined[name_col] if name_col else "Territory_" + joined.index.astype(str)
            
            st.session_state['excel_data'] = generate_excel_report(joined, cong_name, min_g, max_g)
            st.success("Analysis Complete!")
            st.rerun()
        except Exception as e:
            st.error(f"Processing Error: {e}")

if st.session_state['excel_data']:
    st.download_button("⬇️ Download Excel Analysis", st.session_state['excel_data'], f"{cong_name}_Analysis.xlsx")
