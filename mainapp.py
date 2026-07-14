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
    # Strip common unit identifiers to create a "Base Address" for grouping
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
        gdf = gpd.read_file(f"zip://{tfile.name}")
        cols = gdf.columns.tolist()
        
        st.subheader("Data Format")
        method = st.radio("How is your address data stored?", ["Single 'Full Address' column", "Separate columns (House #, Street, etc.)"])
        
        col_map = {'method': method}
        if method == "Single 'Full Address' column":
            col_map['FullAddress'] = st.selectbox("Select 'Full Address' column", cols)
        else:
            fields = ['HouseNo', 'HouseSx', 'Dir', 'Street', 'StType', 'Unit', 'Muni', 'Zip_Code']
            for field in fields:
                col_map[field] = st.selectbox(f"Select column for {field}", cols)
        
        col_map['Status'] = st.selectbox("Select 'Status' column", cols)
        
        if st.button("Confirm Mapping"):
            st.session_state['mappings'] = col_map
            st.session_state['gdf_path'] = tfile.name
            st.rerun()
    except Exception as e: st.error(f"Error: {e}")

if st.session_state['mappings'] and not st.session_state['excluded_values']:
    gdf = gpd.read_file(f"zip://{st.session_state['gdf_path']}")
    unique_vals = gdf[st.session_state['mappings']['Status']].unique().tolist()
    st.subheader("Select Excluded Statuses")
    st.session_state['excluded_values'] = st.multiselect("Select values to treat as 'Excluded' (Tab 6)", unique_vals)

# --- 4. EXCEL GENERATION ENGINE ---
def generate_excel_report(joined_gdf, kml_gdf, min_goal, max_goal, cong_name):
    output = io.BytesIO()
    
    # Standardize Base Address for grouping
    joined_gdf['Base_Address'] = joined_gdf['Mailable_Address'].apply(get_base_address)
    
    # Filter
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
        ideal_pct = (len(counts_df[counts_df['Category'] == 'Ideal']) / total_territories) * 100 if total_territories > 0 else 0
        
        dashboard_top = [
            [f"Territory Analysis: {cong_name}"],
            [f"Generated {datetime.datetime.now().strftime('%B %Y')} by Territory Analysis Engine."],
            [""], [f"Total Territories: {total_territories}"], [f"Total Valid Addresses: {total_addresses}"],
            [f"Excluded Addresses (See Tab 6): {len(excluded_gdf)}"], [""] 
        ]
        pd.DataFrame(dashboard_top).to_excel(writer, sheet_name="Dashboard", index=False, header=False)
        
        # Grid
        distribution = [[counts_df[(counts_df['Total_Addresses'] >= int(r.split('-')[0])) & (counts_df['Total_Addresses'] <= int(r.split('-')[1]))].shape[0], r] for r in ["25-50", "50-75", "75-100", "100-125", "125-150", "150-175"]]
        pd.DataFrame(distribution, columns=["Category", "Range", "Count"]).to_excel(writer, sheet_name="Dashboard", startrow=7, index=False)
        
        # Footer
        ws1 = writer.sheets['Dashboard']
        ws1['A1'].font = Font(size=20, bold=True)
        ws1['A2'].hyperlink = "http://www.territoryanalysis.com/"
        ws1['A2'].font = Font(color="0563C1", underline="single")
        bold_inline = InlineFont(b=True)
        ws1['A15'].value = CellRichText(["About ", TextBlock(bold_inline, f"{ideal_pct:.1f}%"), " of territories fall within this range."])
        
        # Footer Content (No Wrap)
        footer_text = [
            "As a part of this analysis, every address point within your territory was collected & identified.",
            "These addresses, with a little reformatting, can be added to NWS or other programs (Please see http://www.territoryanalysis.com/ to see if your system is supported.)",
            "It's suggested to export this file into a program you can easily edit, like excel or google sheets.",
            "That will allow you to expand cells to read easier, create custom filters to see specific data, and customize the sheet to make it more legible.",
            "",
            "The DASHBOARD tab displays basic statistics about the territory that was analyzed",
            "The COUNTS tab organizes territories by size. This is done by 'counting' workable addresses, not geographical size.",
            "The ADDRESS LIST tab displays every workable address in your territory.",
            "The APARTMENTS tab displays every multifamily above 5 units in your territory.",
            "The BORDER REWRITES tab displays borders within your territory that may benefit from being redrawn.",
            "The EXCLUDED AUDIT tab displays addresses that are NOT counted towards your territory."
        ]
        for i, text in enumerate(footer_text):
            ws1.cell(row=18+i, column=1).value = text

        # Style Sheets 2-6
        # [Insert existing styling logic here...]
        
    output.seek(0)
    return output

# --- 5. EXECUTION FLOW ---
congregation_name = st.text_input("Congregation Name (No Spaces)", "ExampleCongregation")
goal_range = st.selectbox("Goal Range", ["25-50", "50-75", "75-100", "100-125", "125-150"])
MIN_GOAL, MAX_GOAL = [int(x) for x in goal_range.split("-")]

if st.session_state['mappings'] and st.session_state['excluded_values'] is not None and uploaded_kml:
    if st.button("Generate Territory Analysis"):
        try:
            mappings = st.session_state['mappings']
            raw_gdf = gpd.read_file(f"zip://{st.session_state['gdf_path']}")
            
            # Standardization Engine
            if mappings['method'] == "Single 'Full Address' column":
                raw_gdf['Mailable_Address'] = raw_gdf[mappings['FullAddress']]
            else:
                raw_gdf['Mailable_Address'] = raw_gdf[mappings['HouseNo']].astype(str) + " " + raw_gdf[mappings['Street']] 
            
            raw_gdf['Internal_Status'] = raw_gdf[mappings['Status']]
            
            # Spatial Join
            kml_gdf = gpd.read_file(uploaded_kml, driver="KML").make_valid()
            # ... Spatial Join Logic ...
            
            st.session_state['excel_data'] = generate_excel_report(joined_gdf, kml_gdf, MIN_GOAL, MAX_GOAL, congregation_name)
            st.success("Analysis Complete!")
            st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")

if st.session_state['excel_data'] is not None:
    st.write("File is ready for download.")
    st.download_button("⬇️ Download Excel Analysis", st.session_state['excel_data'], "Analysis.xlsx")
