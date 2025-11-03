import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="SMARTLog: Wireline Cost Estimator", layout="wide")
st.title("SMARTLog: Wireline Cost Estimator")

# --- File upload ---
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    # Reset tracker if new file uploaded
    if "last_uploaded_name" not in st.session_state or st.session_state["last_uploaded_name"] != uploaded_file.name:
        st.session_state["unique_tracker"] = set()
        st.session_state["last_uploaded_name"] = uploaded_file.name

    if st.sidebar.button("Reset Unique Tool Tracker"):
        st.session_state["unique_tracker"] = set()
        st.sidebar.success("Unique-tool tracker cleared.")

    df = pd.read_excel(uploaded_file, sheet_name="Data")

    # Ensure required columns exist
    for col in ["Flat Rate", "Depth Charge (per ft)", "Source"]:
        if col not in df.columns:
            df[col] = 0 if "Rate" in col else "Data"

    unique_tools = {"AU14: AUX_SURELOC"}  # Unique tools tracker

    # --- Reference Well Template ---
    reference_well_data = {
        "name": "Reference Well A",
        "Package": "Package A",
        "Service Name": "STANDARD WELLS",
        "Hole Sections": {
            '12.25" Hole': {
                "Depth": 5500,
                "Tools_LineItems": [
                    "AU14: AUX_SURELOC","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU",
                    "GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3","AU2: AUX_PCAL",
                    "PP7: PROC_PETR7","PA7: PROC_ACOU6","PA11: PROC_ACOU13","PA12: PROC_ACOU14",
                    "IM3: IMAG_SOBM","PI1: PROC_IMAG1","PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8",
                    "PI9: PROC_IMAG9","PI12: PROC_IMAG12","PI13: PROC_IMAG13",
                    "FP25: FPS_SCAR","FP18: FPS_SAMP","FP19: FPS_SPHA","FP23: FPS_TRA",
                    "FP24: FPS_TRK","FP28: FPS_FCHA_1","FP33: FPS_FCHA_6","FP34: FPS_FCHA_7",
                    "FP14: FPS_PUMP","FP42: FPS_PROB_XLD","FP11: FPS_PROB_FO","FP26: FPS_FCON",
                    "DT3: RTDT_PER","PPT12: PROC_PT12","SC2: SC_ADD1","SC2: SC_ADD2"
                ]
            },
            '8.5" Hole': {
                "Depth": 8000,
                "Tools_LineItems": [
                    "AU14: AUX_SURELOC","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU",
                    "GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3","AU2: AUX_PCAL",
                    "PP7: PROC_PETR7","PA7: PROC_ACOU6","PA11: PROC_ACOU13","PA12: PROC_ACOU14",
                    "IM3: IMAG_SOBM","PI1: PROC_IMAG1","PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8",
                    "PI9: PROC_IMAG9","PI12: PROC_IMAG12","PI13: PROC_IMAG13",
                    "FP25: FPS_SCAR","FP18: FPS_SAMP","FP19: FPS_SPHA","FP23: FPS_TRA",
                    "FP24: FPS_TRK","FP28: FPS_FCHA_1","FP33: FPS_FCHA_6","FP34: FPS_FCHA_7",
                    "FP14: FPS_PUMP","FP42: FPS_PROB_XLD","FP11: FPS_PROB_FO","FP26: FPS_FCON",
                    "DT3: RTDT_PER","PPT12: PROC_PT12","SC2: SC_ADD1","SC2: SC_ADD2"
                ]
            }
        },
        "Additional Services_LineItems": ["CO1: CONV_PCL","AU7: AUX_SBOX","PC5: PC_10KH2S","PR1: PR_FP","PR2: PR_BO"],
        "Quantity of Tools": 2,
        "Total Months": 1
    }

    # --- Hole Sections Setup ---
    st.sidebar.header("Hole Sections Setup")
    num_sections = st.sidebar.number_input("Number of Hole Sections", min_value=1, max_value=5, value=2)
    hole_sizes = [st.sidebar.text_input(f"Hole Section {i+1} Size (inches)", value=f"{12.25 - i*3.75:.2f}") for i in range(num_sections)]

    tabs = st.tabs([f'{hs}" Hole Section' for hs in hole_sizes])
    section_totals = {}
    all_calc_dfs_for_excel = []

    special_cases_map = {"STANDARD WELLS": {"XL Rock (150DegC Max)": ["AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2"]}}

    # --- Loop through hole sections ---
    for tab, hole_size in zip(tabs, hole_sizes):
        with tab:
            st.header(f'{hole_size}" Hole Section')
            ref_option = st.selectbox("Reference Well", options=["None", reference_well_data["name"]], key=f"ref_{hole_size}")

            # Sidebar inputs
            st.sidebar.subheader(f"Inputs for {hole_size}\" Section")
            qty = st.sidebar.number_input("Quantity of Tools", min_value=1, value=int(st.session_state.get(f"qty_{hole_size}", 2)), key=f"qty_{hole_size}")
            days = st.sidebar.number_input("Total Days", min_value=0, value=int(st.session_state.get(f"days_{hole_size}", 0)), key=f"days_{hole_size}")
            months = st.sidebar.number_input("Total Months", min_value=0, value=int(st.session_state.get(f"months_{hole_size}", 1)), key=f"months_{hole_size}")
            depth = st.sidebar.number_input("Total Depth (ft)", min_value=0, value=int(st.session_state.get(f"depth_{hole_size}", 5500)), key=f"depth_{hole_size}")
            survey = st.sidebar.number_input("Total Survey (ft)", min_value=0, value=int(st.session_state.get(f"survey_{hole_size}", 0)), key=f"survey_{hole_size}")
            hours = st.sidebar.number_input("Total Hours", min_value=0, value=int(st.session_state.get(f"hours_{hole_size}", 0)), key=f"hours_{hole_size}")
            discount = st.sidebar.number_input("Discount (%)", min_value=0.0, max_value=100.0, value=float(st.session_state.get(f"disc_{hole_size}", 0.0)), key=f"disc_{hole_size}") / 100.0

            # --- Package & Service selection ---
            package_options = df["Package"].dropna().unique().tolist()
            selected_package = st.selectbox("Choose Package", package_options, index=0, key=f"pkg_{hole_size}")
            package_df = df[df["Package"] == selected_package]
            service_options = package_df["Service Name"].dropna().unique().tolist()
            selected_service = st.selectbox("Choose Service Name", service_options, index=0, key=f"svc_{hole_size}")
            df_service = package_df[package_df["Service Name"] == selected_service]

            # --- Tool selection ---
            code_list = df_service["Specification 1"].dropna().unique().tolist()
            special_cases = special_cases_map.get(selected_service, {})
            code_list_with_special = list(special_cases.keys()) + code_list

            if ref_option != "None":
                ref_hole_key = f'{hole_size}" Hole'
                reference_codes = reference_well_data["Hole Sections"].get(ref_hole_key, {}).get("Tools_LineItems", []) + reference_well_data["Additional Services_LineItems"]
                default_selected = [c for c in code_list_with_special if c in reference_codes]
                selected_codes = st.multiselect("Select Tools", code_list_with_special, default=default_selected, key=f"tools_{hole_size}")
            else:
                selected_codes = st.multiselect("Select Tools", code_list_with_special, key=f"tools_{hole_size}")

            # --- Expand special cases ---
            expanded_codes = []
            used_special_cases = []
            for code in selected_codes:
                if code in special_cases:
                    expanded_codes.extend(special_cases[code])
                    used_special_cases.append(code)
                else:
                    expanded_codes.append(code)

            df_tools = df_service[df_service["Specification 1"].isin(expanded_codes)].copy() if expanded_codes else pd.DataFrame(columns=df_service.columns)

            # --- Calculations ---
            if not df_tools.empty:
                calc_df = pd.DataFrame()
                calc_df["Code"] = df_tools["Specification 1"]
                calc_df["Daily Rate"] = pd.to_numeric(df_tools.get("Daily Rate", 0), errors="coerce").fillna(0)
                calc_df["Monthly Rate"] = pd.to_numeric(df_tools.get("Monthly Rate", 0), errors="coerce").fillna(0)
                calc_df["Depth Charge (per ft)"] = pd.to_numeric(df_tools.get("Depth Charge (per ft)", 0), errors="coerce").fillna(0)
                calc_df["Flat Rate"] = pd.to_numeric(df_tools.get("Flat Rate", df_tools.get("Flat Rate", 0)), errors="coerce").fillna(0)
                calc_df["Operating Charge (MYR)"] = (calc_df["Depth Charge (per ft)"] * depth + calc_df["Flat Rate"]) * (1 - discount)
                calc_df["Rental Charge (MYR)"] = qty * ((calc_df["Daily Rate"] * days) + (calc_df["Monthly Rate"] * months)) * (1 - discount)
                calc_df["Total (MYR)"] = calc_df["Operating Charge (MYR)"] + calc_df["Rental Charge (MYR)"]

                # Unique tool logic
                tracker = set(st.session_state.get("unique_tracker", set()))
                for ut in unique_tools:
                    mask = calc_df["Code"] == ut
                    if mask.any():
                        if ut in tracker:
                            calc_df.loc[mask, ["Operating Charge (MYR)", "Rental Charge (MYR)"]] = 0
                        tracker.add(ut)
                st.session_state["unique_tracker"] = tracker

                st.dataframe(calc_df)
                section_totals[hole_size] = calc_df["Total (MYR)"].sum()
                st.write(f"💵 Section Total: {section_totals[hole_size]:,.2f}")
                all_calc_dfs_for_excel.append((hole_size, used_special_cases, df_tools, special_cases))

    # --- Grand Total ---
    if section_totals:
        grand_total = sum(section_totals.values())
        st.success(f"🏆 Grand Total Price (MYR): {grand_total:,.2f}")

    # --- Excel Download ---
    if st.button("Download Excel"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for hole_size, used_special_cases, df_tools_section, special_cases_section in all_calc_dfs_for_excel:
                sheet_name = f'{hole_size}" Hole'
                df_tools_section.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        st.download_button("Download Excel", data=output, file_name="Cost_Estimate.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
