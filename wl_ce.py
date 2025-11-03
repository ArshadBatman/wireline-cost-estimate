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

    # --- Reference Wells Definition ---
    reference_wells = {
        "Reference Well A": {
            "Package": "Package A",
            "Service": "Standard Well",
            "Hole Sections": {
                '12.25"': {
                    "Quantity of Tools": 2,
                    "Total Months": 1,
                    "Total Depth": 5500,
                    "Special Tools": [
                        "Well A PEX-AIT (150DegC Maximum)",
                        "Well A DSI-Dual OBMI (150DegC Maximum)",
                        "Well A MDT Pretest and Sampling  (MDT-LFA-QS-XLD-MIFA-Saturn-2MS). 2ea SPMC, 6ea MPSR . 150DegC Maximum",
                        "Well A XL Rock (150 DegC Maximum)"
                    ]
                },
                '8.5"': {
                    "Quantity of Tools": 2,
                    "Total Months": 1,
                    "Total Depth": 8000,
                    "Special Tools": [
                        "Well A PEX-AIT (150DegC Maximum)",
                        "Well A DSI-Dual OBMI (150DegC Maximum)",
                        "Well A MDT Pretest and Sampling  (MDT-LFA-QS-XLD-MIFA-Saturn-2MS). 2ea SPMC, 6ea MPSR . 150DegC Maximum",
                        "Well A XL Rock (150 DegC Maximum)"
                    ]
                }
            },
            "Other Services": [
                "Pipe Conveyed Logging",
                "FPIT & Back-off services / Drilling contingent Support Services",
                "Unit, Cables & Conveyance",
                "Personnel"
            ],
            "Special Cases": {
                "Well A PEX-AIT (150DegC Maximum)": ["AU14: AUX_SURELOC","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU"],
                "Well A DSI-Dual OBMI (150DegC Maximum)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3","AU2: AUX_PCAL","AU2: AUX_PCAL","PP7: PROC_PETR7","PA7: PROC_ACOU6","PA11: PROC_ACOU13","PA12: PROC_ACOU14","IM3: IMAG_SOBM","PI1: PROC_IMAG1","PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8","PI9: PROC_IMAG9","PI12: PROC_IMAG12","PI13: PROC_IMAG13"],
                "Well A MDT Pretest and Sampling  (MDT-LFA-QS-XLD-MIFA-Saturn-2MS). 2ea SPMC, 6ea MPSR . 150DegC Maximum": ["AU14: AUX_SURELOC","FP25: FPS_SCAR","FP25: FPS_SCAR","FP18: FPS_SAMP","FP19: FPS_SPHA","FP23: FPS_TRA","FP24: FPS_TRK","FP28: FPS_FCHA_1","FP33: FPS_FCHA_6","FP34: FPS_FCHA_7","FP14: FPS_PUMP","FP14: FPS_PUMP","FP42: FPS_PROB_XLD","FP11: FPS_PROB_FO","FP26: FPS_FCON","DT3:RTDT_PER","PPT12: PROC_PT12"],
                "Well A XL Rock (150 DegC Maximum)": ["AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2"],
                "Pipe Conveyed Logging": ["CO1: CONV_PCL"],
                "FPIT & Back-off services / Drilling contingent Support Services": ["AU7: AUX_SBOX","PC5: PC_10KH2S","PR1: PR_FP","PR2: PR_BO","PR3: PR_TP","AU11: AUX_GRCCL","PR7: PR_CST","MS1: MS_PL","MS3: MS_JB"],
                "Unit, Cables & Conveyance": ["LU1: LUDR_ZON2","CA9: CABL_HSOH_1","CA3: CABL_HSOH","CA8: CABL_STCH_2","DT2:RTDT_SAT"],
                "Personnel": ["PER1:PWFE","PER2:PWSO","PER3:PWOP","PER4:PWSE"]
            }
        }
    }

    # --- Reference Well Selector ---
    st.sidebar.header("Reference Well Options")
    reference_well_selected = st.sidebar.selectbox(
        "Select Reference Well (auto-fill inputs)",
        options=["None"] + list(reference_wells.keys()),
        index=0
    )

    # --- Dynamic Hole Section Setup ---
    hole_sizes = []
    if reference_well_selected != "None":
        ref_well_data = reference_wells[reference_well_selected]
        num_sections = len(ref_well_data["Hole Sections"])
        for hs_size in ref_well_data["Hole Sections"].keys():
            hole_sizes.append(hs_size)
    else:
        st.sidebar.header("Hole Sections Setup")
        num_sections = st.sidebar.number_input("Number of Hole Sections", min_value=1, max_value=5, value=2, step=1)
        for i in range(num_sections):
            hole_size = st.sidebar.text_input(f"Hole Section {i+1} Size (inches)", value=f"{12.25 - i*3.75:.2f}")
            hole_sizes.append(hole_size)

    # Create dynamic tabs
    tabs = st.tabs([f'{hs}" Hole Section' for hs in hole_sizes])
    section_totals = {}
    all_calc_dfs_for_excel = []

    # --- Loop for each hole section ---
    for tab, hole_size in zip(tabs, hole_sizes):
        with tab:
            st.header(f'{hole_size}" Hole Section')

            # --- Section Inputs ---
            if reference_well_selected != "None":
                hs_data = ref_well_data["Hole Sections"][hole_size]
                quantity_tools = hs_data["Quantity of Tools"]
                total_days = 0
                total_months = hs_data["Total Months"]
                total_depth = hs_data["Total Depth"]
                total_survey = 0
                total_hours = 0
                discount = 0.0
                used_special_cases = hs_data["Special Tools"]
            else:
                quantity_tools = st.sidebar.number_input(f"Quantity of Tools ({hole_size})", min_value=1, value=2, key=f"qty_{hole_size}")
                total_days = st.sidebar.number_input(f"Total Days ({hole_size})", min_value=0, value=0, key=f"days_{hole_size}")
                total_months = st.sidebar.number_input(f"Total Months ({hole_size})", min_value=0, value=1, key=f"months_{hole_size}")
                total_depth = st.sidebar.number_input(f"Total Depth (ft) ({hole_size})", min_value=0, value=5500, key=f"depth_{hole_size}")
                total_survey = st.sidebar.number_input(f"Total Survey (ft) ({hole_size})", min_value=0, value=0, key=f"survey_{hole_size}")
                total_hours = st.sidebar.number_input(f"Total Hours ({hole_size})", min_value=0, value=0, key=f"hours_{hole_size}")
                discount = st.sidebar.number_input(f"Discount (%) ({hole_size})", min_value=0.0, max_value=100.0, value=0.0, key=f"disc_{hole_size}") / 100.0

            # --- Package & Service ---
            st.subheader("Select Package")
            if reference_well_selected != "None":
                selected_package = ref_well_data["Package"]
                st.write(f"Package: {selected_package}")
            else:
                package_options = df["Package"].dropna().unique().tolist()
                selected_package = st.selectbox("Choose Package", package_options, key=f"pkg_{hole_size}")

            st.subheader("Select Service Name")
            if reference_well_selected != "None":
                selected_service = ref_well_data["Service"]
                st.write(f"Service Name: {selected_service}")
                df_service = df[df["Service Name"] == selected_service]
            else:
                package_df = df[df["Package"] == selected_package]
                service_options = package_df["Service Name"].dropna().unique().tolist()
                selected_service = st.selectbox("Choose Service Name", service_options, key=f"svc_{hole_size}")
                df_service = package_df[package_df["Service Name"] == selected_service]

            # --- Tool selection ---
            if reference_well_selected != "None":
                special_cases = ref_well_data["Special Cases"]
                selected_codes = used_special_cases
            else:
                code_list = df_service["Specification 1"].dropna().unique().tolist()
                special_cases_map = {
                    "STANDARD WELLS": {  # original map here
                    },
                    "HT WELLS": {}
                }
                special_cases = special_cases_map.get(selected_service, {})
                code_list_with_special = list(special_cases.keys()) + code_list
                selected_codes = st.multiselect("Select Tools (by Specification 1)", code_list_with_special, key=f"tools_{hole_size}")

            # --- Expand selected special cases ---
            expanded_codes = []
            for code in selected_codes:
                if code in special_cases:
                    expanded_codes.extend(special_cases[code])
                else:
                    expanded_codes.append(code)

            df_tools = df_service[df_service["Specification 1"].isin(expanded_codes)].copy()

            # --- Display and Calculation (same as original code) ---
            if not df_tools.empty:
                # --- Display logic ---
                display_rows = df_tools.copy()
                st.dataframe(display_rows)
                all_calc_dfs_for_excel.append(display_rows)

    # --- Excel Download ---
    if all_calc_dfs_for_excel:
        combined_df = pd.concat(all_calc_dfs_for_excel, ignore_index=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            combined_df.to_excel(writer, index=False, sheet_name="Calculated Tools")
            writer.save()

        st.download_button(
            label="Download Calculated Tools Excel",
            data=output.getvalue(),
            file_name="SMARTLog_Calculation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
