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

# --- 1. CONFIGURATION & UI SETUP ---
st.set_page_config(page_title="County-Agnostic Territory Analyzer", layout="wide")
st.title("Congregation Territory Address Analyzer")

if 'mappings' not in st.session_state: st.session_state['mappings'] = None
if 'excluded_values' not in st.session_state: st.session_state['excluded_values'] = None
if 'excel_data' not in st.session_state: st.session_state['excel_data'] = None

# --- 2. HELPERS ---
def natural_keys(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(text))]

def build_addresses(row):
    house = str(row['Internal_HouseNo']).replace('.0', '').strip() if pd.notna(row['Internal_HouseNo']) and str(row['Internal_HouseNo']).lower() != "nan" else ""
    house_sx = str(row['Internal_HouseSx']).strip() if pd.notna(row['Internal_HouseSx']) and str(row['Internal_HouseSx']).lower() != "nan" else ""
    direction = str(row['Internal_Dir']).strip() if pd.notna(row['Internal_Dir']) and str(row['Internal_Dir']).lower() != "nan" else ""
    street = str(row['Internal_Street']).strip() if pd.notna(row['Internal_Street']) and str(row['Internal_Street']).lower() != "nan" else ""
    st_type = str(row['Internal_StType']).strip() if pd.notna(row['Internal_StType']) and str(row['Internal_StType']).lower() != "nan" else ""
    muni = str(row['Internal_Muni']).strip() if pd.notna(row['Internal_Muni']) and str(row['Internal_Muni']).lower() != "nan" else ""
    zip_c = str(row['Internal_Zip']).strip() if pd.notna(row['Internal_Zip']) and str(row['Internal_Zip']).lower() != "nan" else ""
    unit_val = str(row['Internal_Unit']).strip() if pd.notna(row['Internal_Unit']) and str(row['Internal_Unit']).lower() != "nan" else ""
    unit_str = f" Apt {unit_val}" if unit_val else ""
    full_house_num = f"{house}{house_sx}"
    street_parts = [direction, street, st_type]
    full_street = " ".join([p for p in street_parts if p])
    base_addr_line = f"{full_house_num} {full_street}".strip()
    base_addr = f"{base_addr_line}, {muni}, WI {zip_c}".replace(" ,", ",").strip(" ,")
    mailable_addr_line = f"{base_addr_line}{unit_str}".strip()
    mailable_addr = f"{mailable_addr_line}, {muni}, WI {zip_c}".replace(" ,", ",").strip(" ,")
    return pd.Series([base_addr, mailable_addr])

# --- 3. MAPPING UI ---
st.header("Step 1: Upload Data")
uploaded_shapefile = st.file_uploader("Upload County Shapefile (.zip)", type=["zip"])
uploaded_kml = st.file_uploader("Upload Territory KML File", type=["kml"])

if uploaded_shapefile and not st.session_state['mappings']:
    tfile = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
    tfile.write(uploaded_shapefile.read())
    tfile.close()
    
    try:
        gdf = gpd.read_file(f"zip://{tfile.name}")
        
        # --- THE FIX ---
        if not isinstance(gdf, gpd.GeoDataFrame):
            st.error("The file loaded, but it does not have attribute data (columns). Please verify that your ZIP file contains the required .dbf file along with the .shp file, as the attribute table is missing.")
            st.stop()
        # ----------------
            
        cols = gdf.columns.tolist()
        st.subheader("Map Your Columns")
        col_map = {}
        fields = ['HouseNo', 'HouseSx', 'Dir', 'Street', 'StType', 'Unit', 'Muni', 'Zip_Code', 'Status']
        
        for field in fields:
            # We add an index to the selectbox to make it easier to find columns
            col_map[field] = st.selectbox(f"Select column for {field}", cols, index=cols.index(field) if field in cols else 0)
        
        if st.button("Confirm Mapping"):
            st.session_state['mappings'] = col_map
            st.session_state['gdf_path'] = tfile.name
            st.rerun()
            
    except Exception as e:
        st.error(f"Error reading shapefile: {e}")

if st.session_state['mappings'] and not st.session_state['excluded_values']:
    gdf = gpd.read_file(f"zip://{st.session_state['gdf_path']}")
    unique_vals = gdf[st.session_state['mappings']['Status']].unique().tolist()
    st.subheader("Select Excluded Statuses")
    st.session_state['excluded_values'] = st.multiselect("Select values to treat as 'Excluded' (Tab 6)", unique_vals)

# --- 4. EXCEL GENERATION ENGINE ---
def generate_excel_report(joined_gdf, kml_gdf, min_goal, max_goal, cong_name):
    output = io.BytesIO()
    joined_gdf['Internal_Zip'] = joined_gdf['Internal_Zip'].astype(str).str[:5]
    joined_gdf[['Base_Address', 'Mailable_Address']] = joined_gdf.apply(build_addresses, axis=1)
    
    excluded_gdf = joined_gdf[joined_gdf['Internal_Status'].isin(st.session_state['excluded_values'])].copy()
    valid_gdf = joined_gdf[~joined_gdf['Internal_Status'].isin(st.session_state['excluded_values'])].copy()

    unique_territories = valid_gdf['Territory_Name'].unique().tolist()
    unique_territories.sort(key=natural_keys)
    valid_gdf['Territory_Name'] = pd.Categorical(valid_gdf['Territory_Name'], categories=unique_territories, ordered=True)
    
    counts_df = valid_gdf.groupby('Territory_Name', observed=True).size().reset_index(name='Total_Addresses')
    counts_df = counts_df[counts_df['Total_Addresses'] > 0]
    counts_df['Category'] = counts_df['Total_Addresses'].apply(lambda c: "Undersized" if c < min_goal else ("Ideal" if min_goal <= c <= max_goal else "Oversized"))
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Dashboard Logic
        total_territories = len(counts_df)
        total_addresses = counts_df['Total_Addresses'].sum()
        largest_terr = counts_df.loc[counts_df['Total_Addresses'].idxmax()] if total_territories > 0 else None
        smallest_terr = counts_df.loc[counts_df['Total_Addresses'].idxmin()] if total_territories > 0 else None
        ideal_pct = (len(counts_df[counts_df['Category'] == 'Ideal']) / total_territories) * 100 if total_territories > 0 else 0
        
        dashboard_top = [
            [f"Territory Analysis: {cong_name}"],
            [f"Generated {datetime.datetime.now().strftime('%B %Y')} by Territory Analysis Engine."],
            [""],
            [f"Total Territories: {total_territories}"],
            [f"Total Valid Addresses: {total_addresses}"],
            [f"Excluded Addresses (See Tab 6): {len(excluded_gdf)}"],
            [f"The largest territory has {largest_terr['Total_Addresses']} addresses in it ({largest_terr['Territory_Name']})." if largest_terr is not None else ""],
            [f"The smallest territory has {smallest_terr['Total_Addresses']} addresses in it ({smallest_terr['Territory_Name']})." if smallest_terr is not None else ""],
            [""],
            [f"Goal Range: {min_goal}-{max_goal}"],
            [""] 
        ]
        pd.DataFrame(dashboard_top).to_excel(writer, sheet_name="Dashboard", index=False, header=False)
        
        ranges = ["25-50", "50-75", "75-100", "100-125", "125-150", "150-175"]
        distribution = [[counts_df[(counts_df['Total_Addresses'] >= int(r.split('-')[0])) & (counts_df['Total_Addresses'] <= int(r.split('-')[1]))].shape[0], r] for r in ranges]
        
        ws1 = writer.sheets['Dashboard']
        ws1['A1'].font = Font(size=20, bold=True)
        ws1['A2'].hyperlink = "http://www.territoryanalysis.com/"
        ws1['A2'].font = Font(color="0563C1", underline="single")
        
        # ... [Rest of formatting logic] ...
        
    output.seek(0)
    return output

# --- 5. EXECUTION FLOW ---
if st.session_state['mappings'] and st.session_state['excluded_values'] is not None and uploaded_kml:
    if st.button("Generate Territory Analysis"):
        try:
            mappings = st.session_state['mappings']
            raw_gdf = gpd.read_file(f"zip://{st.session_state['gdf_path']}").rename(columns={
                mappings['HouseNo']: 'Internal_HouseNo', mappings['HouseSx']: 'Internal_HouseSx',
                mappings['Dir']: 'Internal_Dir', mappings['Street']: 'Internal_Street',
                mappings['StType']: 'Internal_StType', mappings['Unit']: 'Internal_Unit',
                mappings['Muni']: 'Internal_Muni', mappings['Zip_Code']: 'Internal_Zip',
                mappings['Status']: 'Internal_Status'
            })
            
            kml_gdf = gpd.read_file(uploaded_kml, driver="KML").make_valid()
            # Defensive Name Parsing
            name_cols = ['Name', 'name', 'Title', 'title', 'Description', 'description']
            name_col = next((col for col in name_cols if col in kml_gdf.columns), None)
            fallback = "Territory_" + kml_gdf.index.astype(str)
            kml_gdf['Territory_Name'] = kml_gdf[name_col].fillna(fallback) if name_col else fallback
            
            parcel_gdf = raw_gdf.to_crs(kml_gdf.crs)
            joined_gdf = gpd.sjoin(parcel_gdf, kml_gdf.rename(columns={'geometry': 'geometry_terr'}).set_geometry('geometry_terr'), how="inner", predicate="within")
            joined_gdf = joined_gdf.dropna(subset=['Territory_Name'])
            
            st.session_state['excel_data'] = generate_excel_report(joined_gdf, kml_gdf, MIN_GOAL, MAX_GOAL, congregation_name.replace(" ", ""))
            st.success("Analysis Complete!")
        except Exception as e:
            st.session_state['excel_data'] = None
            st.error(f"Error processing files: {e}")

if st.session_state['excel_data'] is not None:
    st.write("File is ready for download.")
    st.download_button(
        label="⬇️ Download Excel Analysis",
        data=st.session_state['excel_data'],
        file_name=f"{congregation_name.replace(' ', '')}_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )