import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill

st.title("SMARTLog: Wireline Cost Estimator")

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

    # --- Dynamic Hole Section Setup ---
    st.sidebar.header("Hole Sections Setup")
    num_sections = st.sidebar.number_input("Number of Hole Sections", min_value=1, max_value=5, value=2, step=1)
    hole_sizes = []
    for i in range(num_sections):
        hole_size = st.sidebar.text_input(f"Hole Section {i+1} Size (inches)", value=f"{12.25 - i*3.75:.2f}")
        hole_sizes.append(hole_size)

    # Create dynamic tabs
    tabs = st.tabs([f'{hs}" Hole Section' for hs in hole_sizes])
    section_totals = {}
    all_calc_dfs_for_excel = []  # store data for Excel download

    # --- Reference Well preset tools ---
    reference_well_tools = {
        "Reference Well A": {
            '12.25" Hole': [
                "PEX-AIT (150DegC Max)",
                "DSI-Dual OBMI (150DegC Max)",
                "MDT: LFA-QS-XLD-MIFA-Saturn-2MS (150DegC Max)",
                "XL Rock (150DegC Max)"
            ],
            '8.5" Hole': [
                "PEX-AIT (150DegC Max)",
                "DSI-Dual OBMI (150DegC Max)",
                "MDT: LFA-QS-XLD-MIFA-Saturn-2MS (150DegC Max)",
                "XL Rock (150DegC Max)"
            ]
        }
    }

    # --- Loop for each hole section ---
    for tab, hole_size in zip(tabs, hole_sizes):
        with tab:
            st.header(f'{hole_size}" Hole Section')

            # Sidebar inputs per section
            st.sidebar.subheader(f"Inputs for {hole_size}\" Section")
            quantity_tools = st.sidebar.number_input(f"Quantity of Tools ({hole_size})", min_value=1, value=2, key=f"qty_{hole_size}")
            total_days = st.sidebar.number_input(f"Total Days ({hole_size})", min_value=0, value=0, key=f"days_{hole_size}")
            total_months = st.sidebar.number_input(f"Total Months ({hole_size})", min_value=0, value=1, key=f"months_{hole_size}")
            total_depth = st.sidebar.number_input(f"Total Depth (ft) ({hole_size})", min_value=0, value=5500, key=f"depth_{hole_size}")
            total_survey = st.sidebar.number_input(f"Total Survey (ft) ({hole_size})", min_value=0, value=0, key=f"survey_{hole_size}")
            total_hours = st.sidebar.number_input(f"Total Hours ({hole_size})", min_value=0, value=0, key=f"hours_{hole_size}")
            discount = st.sidebar.number_input(f"Discount (%) ({hole_size})", min_value=0.0, max_value=100.0, value=0.0, key=f"disc_{hole_size}") / 100.0

            # --- Package & Service ---
            st.subheader("Select Package")
            package_options = df["Package"].dropna().unique().tolist()
            selected_package = st.selectbox(
                "Choose Package",
                package_options,
                index=0,
                key=f"pkg_{hole_size}"
            )
            package_df = df[df["Package"] == selected_package]

            st.subheader("Select Service Name")
            service_options = package_df["Service Name"].dropna().unique().tolist()
            selected_service = st.selectbox("Choose Service Name", service_options, key=f"svc_{hole_size}")
            df_service = package_df[package_df["Service Name"] == selected_service]

            # Filter: include selected service + blank cells for Specification 1
            df_service = package_df[
                (package_df["Service Name"] == selected_service) |
                (package_df["Service Name"].isna()) |
                (package_df["Service Name"] == "")
            ]

            # --- Reference Well selection ---
            reference_wells = ["", "Reference Well A"]
            if any(st.session_state.get(f"ref_well_{hs}") == "Reference Well A" for hs in hole_sizes):
                default_ref = "Reference Well A"
            else:
                default_ref = ""
            selected_ref_well = st.selectbox(
                "Reference Well",
                reference_wells,
                index=reference_wells.index(default_ref),
                key=f"ref_well_{hole_size}"
            )

            # --- Tool selection with special cases ---
            code_list = df_service["Specification 1"].dropna().unique().tolist()
            special_cases_map = {
                "STANDARD WELLS": {
                    "PEX-AIT (150DegC Max)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU"],
                    "PEX-AIT-DSI (150DegC Max)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU",
                                                 "AU3:AUX_INCL", "AU2: AUX_PCAL", "AU2: AUX_PCAL", "AC3: ACOU_3", "PP7: PROC_PETR7", "PA7: PROC_ACOU6",
                                                 "PA11: PROC_ACOU13", "PA12: PROC_ACOU14"],
                    "DOBMI (150DegC Max)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3",
                                            "AU2: AUX_PCAL","AU2: AUX_PCAL","PP7: PROC_PETR7","PA7: PROC_ACOU6",
                                            "PA11: PROC_ACOU13","PA12: PROC_ACOU14","IM3: IMAG_SOBM","PI1: PROC_IMAG1",
                                            "PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8","PI9: PROC_IMAG9",
                                            "PI12: PROC_IMAG12","PI13: PROC_IMAG13"],
                    "MDT: LFA-QS-XLD-MIFA-Saturn-2MS (150DegC Max)": ["AU14: AUX_SURELOC","FP25: FPS_SCAR","FP25: FPS_SCAR",
                                                                      "FP18: FPS_SAMP","FP19: FPS_SPHA","FP23: FPS_TRA",
                                                                      "FP24: FPS_TRK","FP28: FPS_FCHA_1","FP33: FPS_FCHA_6",
                                                                      "FP34: FPS_FCHA_7","FP14: FPS_PUMP","FP14: FPS_PUMP",
                                                                      "FP42: FPS_PROB_LD","FP11: FPS_PROB_FO","FP26: FPS_FCON",
                                                                      "DT3: RTDT_PER","PPT12: PROC_PT12","FP7: FPS_SPPT_2"],
                    "XL Rock (150DegC Max)": ["AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2"]
                },
                "HT WELLS": {}
            }

            special_cases = special_cases_map.get(selected_service, {})
            code_list_with_special = list(special_cases.keys()) + code_list

            # Auto-select tools if Reference Well is chosen
            if selected_ref_well == "Reference Well A":
                preset_tools = reference_well_tools[selected_ref_well].get(f'{hole_size}" Hole', [])
                selected_codes = st.multiselect(
                    "Select Tools (by Specification 1)",
                    code_list_with_special,
                    default=preset_tools,
                    key=f"tools_{hole_size}"
                )
                # propagate Reference Well A to all holes
                for hs in hole_sizes:
                    st.session_state[f"ref_well_{hs}"] = "Reference Well A"
            else:
                selected_codes = st.multiselect(
                    "Select Tools (by Specification 1)",
                    code_list_with_special,
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

            # --- Row-by-row display with dividers ---
            display_rows = []

            # Add Reference Well divider first if selected
            if selected_ref_well == "Reference Well A":
                divider = pd.DataFrame({col: "" for col in df_tools.columns}, index=[0])
                divider["Specification 1"] = f"--- {selected_ref_well} ---"
                display_rows.append(divider)

            # Special cases dividers
            for sc in used_special_cases:
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

            # --- Calculate costs ---
            display_df["Flat Rate"] = display_df["Flat Rate"].fillna(0)
            display_df["Depth Charge (per ft)"] = display_df["Depth Charge (per ft)"].fillna(0)
            display_df["Total Flat"] = display_df["Flat Rate"] * quantity_tools
            display_df["Total Depth"] = display_df["Depth Charge (per ft)"] * total_depth
            display_df["Total Rental"] = display_df["Total Flat"] + display_df["Total Depth"]
            display_df["Total Rental"] = display_df["Total Rental"] * (1 - discount)
            section_totals[hole_size] = display_df["Total Rental"].sum()

            all_calc_dfs_for_excel.append(display_df)

    # --- Excel Export ---
    if st.button("Download Excel"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for hs, df_calc in zip(hole_sizes, all_calc_dfs_for_excel):
                sheet_name = f'{hs}" Hole'
                df_calc.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                # Header formatting
                for col_idx, col in enumerate(df_calc.columns, start=1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                # Divider formatting
                for row_idx, val in enumerate(df_calc["Specification 1"], start=2):
                    if str(val).startswith("---"):
                        for col_idx in range(1, len(df_calc.columns)+1):
                            cell = ws.cell(row=row_idx, column=col_idx)
                            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                            cell.alignment = Alignment(horizontal="center")

        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="Wireline_Cost_Estimate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

