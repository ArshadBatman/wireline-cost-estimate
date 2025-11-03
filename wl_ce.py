import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill

st.title("SMARTLog: Wireline Cost Estimator")

# --- Reference Wells Data ---
reference_wells = {
    "Reference Well A": {
        "Package": "Package A",
        "Well Type": "Standard Well",
        "Hole Sections": {
            "12.25": {
                "Quantity of Tools": 2,
                "Total Months": 1,
                "Depth": 5500,
                "Tools": [
                    "PEX-AIT (150DegC Maximum)",
                    "DSI-Dual OBMI (150DegC Maximum)",
                    "MDT Pretest and Sampling (MDT-LFA-QS-XLD-MIFA-2MS)",
                    "XL Rock (150 DegC Maximum)"
                ]
            },
            "8.5": {
                "Quantity of Tools": 2,
                "Total Months": 1,
                "Depth": 8000,
                "Tools": [
                    "PEX-AIT (150DegC Maximum)",
                    "DSI-Dual OBMI (150DegC Maximum)",
                    "MDT Pretest and Sampling (MDT-LFA-QS-XLD-MIFA-2MS)",
                    "XL Rock (150 DegC Maximum)"
                ]
            }
        },
        "Additional Items": [
            "Pipe Conveyed Logging",
            "FPIT & Back-off services / Drilling contingent Support Services",
            "Unit, Cables & Conveyance",
            "Personnel"
        ]
    }
}

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

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

    # --- Reference Well Selection ---
    st.sidebar.subheader("Reference Well Selection")
    ref_well_options = ["None"] + list(reference_wells.keys())
    selected_ref_well = st.sidebar.selectbox("Choose Reference Well", ref_well_options)

    # --- Dynamic Hole Section Setup ---
    st.sidebar.header("Hole Sections Setup")
    if selected_ref_well == "None":
        num_sections = st.sidebar.number_input("Number of Hole Sections", min_value=1, max_value=5, value=2, step=1)
        hole_sizes = []
        for i in range(num_sections):
            hole_size = st.sidebar.text_input(f"Hole Section {i+1} Size (inches)", value=f"{12.25 - i*3.75:.2f}")
            hole_sizes.append(hole_size)
    else:
        ref_data = reference_wells[selected_ref_well]
        hole_sizes = list(ref_data["Hole Sections"].keys())
        hole_sizes = [str(hs) for hs in hole_sizes]  # ensure string for consistent usage

    # Create dynamic tabs
    tabs = st.tabs([f'{hs}" Hole Section' for hs in hole_sizes])
    section_totals = {}
    all_calc_dfs_for_excel = []  # store data for Excel download

    # --- Loop for each hole section ---
    for tab, hole_size in zip(tabs, hole_sizes):
        with tab:
            st.header(f'{hole_size}" Hole Section')

            # Sidebar inputs per section
            st.sidebar.subheader(f"Inputs for {hole_size}\" Section")
            quantity_tools = st.sidebar.number_input(
                f"Quantity of Tools ({hole_size})",
                min_value=1,
                value=ref_data["Hole Sections"][hole_size]["Quantity of Tools"] if selected_ref_well != "None" else 2,
                key=f"qty_{hole_size}"
            )
            total_days = st.sidebar.number_input(f"Total Days ({hole_size})", min_value=0, value=0, key=f"days_{hole_size}")
            total_months = st.sidebar.number_input(
                f"Total Months ({hole_size})",
                min_value=0,
                value=ref_data["Hole Sections"][hole_size]["Total Months"] if selected_ref_well != "None" else 1,
                key=f"months_{hole_size}"
            )
            total_depth = st.sidebar.number_input(
                f"Total Depth (ft) ({hole_size})",
                min_value=0,
                value=ref_data["Hole Sections"][hole_size]["Depth"] if selected_ref_well != "None" else 5500,
                key=f"depth_{hole_size}"
            )
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
                "STANDARD WELLS": {
                    "PEX-AIT (150DegC Max)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU"],
                    "DSI-Dual OBMI (150DegC Max)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU",
                                                 "AU3:AUX_INCL", "AU2: AUX_PCAL", "AU2: AUX_PCAL", "AC3: ACOU_3", "PP7: PROC_PETR7", "PA7: PROC_ACOU6",
                                                 "PA11: PROC_ACOU13", "PA12: PROC_ACOU14"],
                    "MDT Pretest and Sampling (MDT-LFA-QS-XLD-MIFA-2MS)": ["AU14: AUX_SURELOC","FP25: FPS_SCAR","FP25: FPS_SCAR",
                                                                      "FP18: FPS_SAMP","FP19: FPS_SPHA","FP23: FPS_TRA",
                                                                      "FP24: FPS_TRK","FP28: FPS_FCHA_1","FP33: FPS_FCHA_6",
                                                                      "FP34: FPS_FCHA_7","FP14: FPS_PUMP","FP14: FPS_PUMP",
                                                                      "FP42: FPS_PROB_XLD","FP11: FPS_PROB_FO","FP26: FPS_FCON",
                                                                      "DT3: RTDT_PER","PPT12: PROC_PT12","FP7: FPS_SPPT_2"],
                    "XL Rock (150DegC Max)": ["AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2"]
                },
                "HT WELLS": {}
            }

            special_cases = special_cases_map.get(selected_service, {})
            code_list_with_special = list(special_cases.keys()) + code_list

            # Pre-select tools if reference well is chosen
            preselected_tools = []
            if selected_ref_well != "None":
                preselected_tools = ref_data["Hole Sections"][hole_size]["Tools"]

            selected_codes = st.multiselect(
                "Select Tools (by Specification 1)",
                code_list_with_special,
                default=preselected_tools,
                key=f"tools_{hole_size}"
            )

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

            # --- Row-by-row display ---
            if not df_tools.empty:
                display_rows = []
                inserted_dividers = set()
                for sc in used_special_cases:
                    # Divider row
                    divider = pd.DataFrame({col: "" for col in df_tools.columns}, index=[0])
                    divider["Specification 1"] = f"--- {sc} ---"
                    display_rows.append(divider)
                    for item in special_cases[sc]:
                        item_rows = df_tools[df_tools["Specification 1"] == item]
                        display_rows.extend([row.to_frame().T for _, row in item_rows.iterrows()])
                # Non-special tools
                for item in df_tools["Specification 1"]:
                    if item not in sum(special_cases.values(), []):
                        item_rows = df_tools[df_tools["Specification 1"] == item]
                        display_rows.extend([row.to_frame().T for _, row in item_rows.iterrows()])
                display_df = pd.concat(display_rows, ignore_index=True)
                def highlight_divider(row):
                    if str(row["Specification 1"]).startswith("---"):
                        return ["background-color: red; color: white"] * len(row)
                    return [""] * len(row)
                st.subheader(f"Selected Data - Package {selected_package}, Service {selected_service}")
                st.dataframe(display_df.style.apply(highlight_divider, axis=1))

            # --- Calculation ---
            if not df_tools.empty:
                calc_df = pd.DataFrame()
                calc_df["Source"] = df_tools["Source"]
                calc_df["Ref Item"] = df_tools["Reference"]
                calc_df["Code"] = df_tools["Specification 1"].astype(str).str.strip()
                calc_df["Items"] = df_tools["Specification 2"]
                calc_df["Daily Rate"] = pd.to_numeric(df_tools["Daily Rate"], errors="coerce").fillna(0)
                calc_df["Monthly Rate"] = pd.to_numeric(df_tools["Monthly Rate"], errors="coerce").fillna(0)
                calc_df["Depth Charge (per ft)"] = pd.to_numeric(df_tools["Depth Charge (per ft)"], errors="coerce").fillna(0)
                calc_df["Flat Rate"] = pd.to_numeric(df_tools["Flat Charge"], errors="coerce").fillna(0)
                calc_df["Survey Charge (per ft)"] = 0
                calc_df["Hourly Charge"] = 0
                calc_df["User Flat Charge"] = calc_df["Flat Rate"].apply(lambda x: 1 if x > 0 else 0)
                calc_df = st.data_editor(calc_df, num_rows="dynamic", key=f"editor_{hole_size}")
                calc_df["Quantity of Tools"] = quantity_tools
                calc_df["Total Days"] = total_days
                calc_df["Total Months"] = total_months
                calc_df["Total Depth"] = total_depth
                calc_df["Total Survey"] = total_survey
                calc_df["Total Hours"] = total_hours
                calc_df["Discount"] = discount
                section_totals[hole_size] = calc_df
                all_calc_dfs_for_excel.append((hole_size, used_special_cases, calc_df, special_cases))

    # --- Excel Export ---
    def export_to_excel(all_calc_dfs_for_excel, selected_ref_well):
        output = BytesIO()
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Wireline Estimation"
        current_row = 1

        # Column headers
        headers = ["Ref", "Specification 1", "Specification 2", "Daily Rate", "Monthly Rate", "Depth Charge (per ft)", "Flat Charge"]
        for idx, h in enumerate(headers, start=1):
            ws.cell(row=current_row, column=idx, value=h)
            ws.cell(row=current_row, column=idx).fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            ws.cell(row=current_row, column=idx).alignment = Alignment(horizontal="center")
        current_row += 1

        for hole_size, used_special_cases, df_tools_section, special_cases_section in all_calc_dfs_for_excel:
            # --- Insert Reference Well header ---
            if selected_ref_well != "None":
                ws[f"B{current_row}"] = selected_ref_well
                ws[f"B{current_row}"].fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")  # light yellow
                ws[f"B{current_row}"].alignment = Alignment(horizontal="center", vertical="center")
                current_row += 1

                # Insert Additional Items
                additional_items = reference_wells[selected_ref_well].get("Additional Items", [])
                for item in additional_items:
                    ws[f"B{current_row}"] = item
                    ws[f"B{current_row}"].fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # pale yellow
                    ws[f"B{current_row}"].alignment = Alignment(horizontal="left", vertical="center")
                    current_row += 1

            # Hole Section Header
            ws[f"B{current_row}"] = f'{hole_size}" Hole Section'
            ws[f"B{current_row}"].fill = PatternFill(start_color="CCCCFF", end_color="CCCCFF", fill_type="solid")  # light blue
            ws[f"B{current_row}"].alignment = Alignment(horizontal="center")
            current_row += 1

            for idx, row in df_tools_section.iterrows():
                for col_idx, col in enumerate(headers, start=1):
                    ws.cell(row=current_row, column=col_idx, value=row.get(col, ""))
                current_row += 1

        wb.save(output)
        return output

    st.subheader("Export to Excel")
    if st.button("Download Excel"):
        excel_data = export_to_excel(all_calc_dfs_for_excel, selected_ref_well)
        st.download_button(
            label="Download Excel File",
            data=excel_data,
            file_name="Wireline_Estimation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
