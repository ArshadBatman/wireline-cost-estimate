import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill

st.title("SMARTLog: Wireline Cost Estimator")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# --- Reference Wells ---
reference_wells = {
    "Well A": {
        "Package": "Package A",
        "Hole Sections": {
            '12.25"': {"Quantity": 2, "Total Months": 1, "Depth": 5500},
            '8.5"': {"Quantity": 2, "Total Months": 1, "Depth": 8000}
        },
        "Special Tools": {
            '12.25"': {
                "Well A PEX-AIT (150DegC Maximum)": ["AU14: AUX_SURELOC","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU"],
                "Well A DSI-Dual OBMI (150DegC Maximum)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3",
                                                           "AU2: AUX_PCAL","AU2: AUX_PCAL","PP7: PROC_PETR7","PA7: PROC_ACOU6",
                                                           "PA11: PROC_ACOU13","PA12: PROC_ACOU14","IM3: IMAG_SOBM","PI1: PROC_IMAG1",
                                                           "PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8","PI9: PROC_IMAG9",
                                                           "PI12: PROC_IMAG12","PI13: PROC_IMAG13"]
            },
            '8.5"': {
                "Well A PEX-AIT (150DegC Maximum)": ["AU14: AUX_SURELOC","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU"],
                "Well A DSI-Dual OBMI (150DegC Maximum)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3",
                                                           "AU2: AUX_PCAL","AU2: AUX_PCAL","PP7: PROC_PETR7","PA7: PROC_ACOU6",
                                                           "PA11: PROC_ACOU13","PA12: PROC_ACOU14","IM3: IMAG_SOBM","PI1: PROC_IMAG1",
                                                           "PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8","PI9: PROC_IMAG9",
                                                           "PI12: PROC_IMAG12","PI13: PROC_IMAG13"]
            }
        }
    }
}

# --- Reference Well Selector ---
st.sidebar.header("Reference Well Selection")
selected_well = st.sidebar.selectbox("Reference Well", ["None"] + list(reference_wells.keys()))

# --- Default values if a reference well is selected ---
hole_sizes_defaults = []
hole_data_defaults = {}
special_tools_defaults = {}
if selected_well != "None":
    well_info = reference_wells[selected_well]
    hole_sizes_defaults = list(well_info["Hole Sections"].keys())
    hole_data_defaults = well_info["Hole Sections"]
    special_tools_defaults = well_info["Special Tools"]

if uploaded_file:
    # --- Reset unique tracker when a new file is uploaded ---
    if "last_uploaded_name" not in st.session_state or st.session_state["last_uploaded_name"] != uploaded_file.name:
        st.session_state["unique_tracker"] = set()
        st.session_state["last_uploaded_name"] = uploaded_file.name

    # Optional manual reset button
    if st.sidebar.button("Reset unique-tool usage (AU14, etc.)", key="reset_unique"):
        st.session_state["unique_tracker"] = set()
        st.sidebar.success("Unique-tool tracker cleared.")

    # Read data
    df = pd.read_excel(uploaded_file, sheet_name="Data")

    # Ensure required columns exist
    for col in ["Flat Rate", "Depth Charge (per ft)", "Source"]:
        if col not in df.columns:
            df[col] = 0 if "Rate" in col else "Data"

    # Unique tools across sections
    unique_tools = {"AU14: AUX_SURELOC"}

    # --- Dynamic Hole Section Setup ---
    st.sidebar.header("Hole Sections Setup")
    num_sections = st.sidebar.number_input(
        "Number of Hole Sections",
        min_value=1, max_value=5,
        value=len(hole_sizes_defaults) if hole_sizes_defaults else 2,
        step=1
    )

    hole_sizes = []
    for i in range(num_sections):
        default_size = hole_sizes_defaults[i] if i < len(hole_sizes_defaults) else f"{12.25 - i*3.75:.2f}"
        hole_size = st.sidebar.text_input(f"Hole Section {i+1} Size (inches)", value=default_size)
        hole_sizes.append(hole_size)

    # Create dynamic tabs
    tabs = st.tabs([f'{hs}" Hole Section' for hs in hole_sizes])
    section_totals = {}
    all_calc_dfs_for_excel = []

    # --- Loop for each hole section ---
    for tab, hole_size in zip(tabs, hole_sizes):
        with tab:
            st.header(f'{hole_size}" Hole Section')

            # --- Populate defaults ---
            default_qty = hole_data_defaults.get(hole_size, {}).get("Quantity", 2)
            default_months = hole_data_defaults.get(hole_size, {}).get("Total Months", 1)
            default_depth = hole_data_defaults.get(hole_size, {}).get("Depth", 5500)

            quantity_tools = st.sidebar.number_input(f"Quantity of Tools ({hole_size})", min_value=1, value=default_qty, key=f"qty_{hole_size}")
            total_months = st.sidebar.number_input(f"Total Months ({hole_size})", min_value=0, value=default_months, key=f"months_{hole_size}")
            total_depth = st.sidebar.number_input(f"Total Depth (ft) ({hole_size})", min_value=0, value=default_depth, key=f"depth_{hole_size}")
            total_days = st.sidebar.number_input(f"Total Days ({hole_size})", min_value=0, value=0, key=f"days_{hole_size}")
            total_survey = st.sidebar.number_input(f"Total Survey (ft) ({hole_size})", min_value=0, value=0, key=f"survey_{hole_size}")
            total_hours = st.sidebar.number_input(f"Total Hours ({hole_size})", min_value=0, value=0, key=f"hours_{hole_size}")
            discount = st.sidebar.number_input(f"Discount (%) ({hole_size})", min_value=0.0, max_value=100.0, value=0.0, key=f"disc_{hole_size}") / 100.0

            # --- Package & Service ---
            st.subheader("Select Package")
            package_options = df["Package"].dropna().unique().tolist()
            selected_package = st.selectbox("Choose Package", package_options, key=f"pkg_{hole_size}")
            package_df = df[df["Package"] == selected_package]

            st.subheader("Select Service Name")
            service_options = package_df["Service Name"].dropna().unique().tolist()
            selected_service = st.selectbox("Choose Service Name", service_options, key=f"svc_{hole_size}")
            df_service = package_df[package_df["Service Name"] == selected_service]

            # --- Tool selection with special cases ---
            code_list = df_service["Specification 1"].dropna().unique().tolist()
            special_cases_map = {
                "STANDARD WELLS": special_tools_defaults.get(hole_size, {}),
                "HT WELLS": {}
            }
            special_cases = special_cases_map.get("STANDARD WELLS", {})
            code_list_with_special = list(special_cases.keys()) + code_list

            # Preselect special tools if well is selected
            preselected_tools = list(special_cases.keys()) if selected_well != "None" else []
            selected_codes = st.multiselect("Select Tools (by Specification 1)", code_list_with_special, default=preselected_tools, key=f"tools_{hole_size}")

            # --- Expand selected special cases ---
            expanded_codes = []
            used_special_cases = []
            for code in selected_codes:
                if code in special_cases:
                    expanded_codes.extend(special_cases[code])
                    used_special_cases.append(code)
                else:
                    expanded_codes.append(code)

            df_tools = df_service[df_service["Specification 1"].isin(expanded_codes)].copy()

            # --- Continue with your existing display, calculation, and Excel export logic ---
            # (No changes needed below here)
            # [Keep the rest of your code for display_df, calc_df, section totals, Excel export]
