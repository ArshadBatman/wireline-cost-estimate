import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill

st.title("SMARTLog: Wireline Cost Estimator")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Fixed Reference Well A data
reference_well_A = {
    "Package": "Package A",
    "Standard Well": "Standard Well",
    "Hole Sections": {
        "12.25": {
            "Tools": [
                "PEX-AIT (150DegC Maximum)",
                "DSI-Dual OBMI (150DegC Maximum)",
                "MDT Pretest and Sampling  (MDT-LFA-QS-XLD-MIFA-Saturn-2MS). 2ea SPMC, 6ea MPSR . 150DegC Maximum",
                "XL Rock (150 DegC Maximum)",
                "Pipe Conveyed Logging",
                "FPIT & Back-off services / Drilling ontingent Support Services",
                "Unit, Cables & Conveyance",
                "Personnel"
            ]
        },
        "8.5": {
            "Tools": [
                "PEX-AIT (150DegC Maximum)",
                "DSI-Dual OBMI (150DegC Maximum)",
                "MDT Pretest and Sampling  (MDT-LFA-QS-XLD-MIFA-Saturn-2MS). 2ea SPMC, 6ea MPSR . 150DegC Maximum",
                "XL Rock (150 DegC Maximum)",
                "Pipe Conveyed Logging",
                "FPIT & Back-off services / Drilling ontingent Support Services",
                "Unit, Cables & Conveyance",
                "Personnel"
            ]
        }
    },
    "Quantity of Tools": 2,
    "Total Months": 1,
    "Depths": {"12.25": 5500, "8.5": 8000}
}

# Unique tools tracker
unique_tools = {"AU14: AUX_SURELOC"}

if uploaded_file:
    # Reset unique tracker when a new file is uploaded
    if "last_uploaded_name" not in st.session_state or st.session_state["last_uploaded_name"] != uploaded_file.name:
        st.session_state["unique_tracker"] = set()
        st.session_state["last_uploaded_name"] = uploaded_file.name

    # Optional reset button
    if st.sidebar.button("Reset unique-tool usage (AU14, etc.)"):
        st.session_state["unique_tracker"] = set()
        st.sidebar.success("Unique-tool tracker cleared.")

    # Read Excel
    df = pd.read_excel(uploaded_file, sheet_name="Data")

    # Ensure required columns
    for col in ["Flat Rate", "Depth Charge (per ft)", "Source"]:
        if col not in df.columns:
            df[col] = 0 if "Rate" in col else "Data"

    # Dynamic Hole Section Setup
    st.sidebar.header("Hole Sections Setup")
    num_sections = st.sidebar.number_input("Number of Hole Sections", min_value=1, max_value=5, value=2, step=1)
    hole_sizes = []
    for i in range(num_sections):
        hole_size = st.sidebar.text_input(f"Hole Section {i+1} Size (inches)", value=f"{12.25 - i*3.75:.2f}")
        hole_sizes.append(hole_size)

    # Create tabs
    tabs = st.tabs([f'{hs}" Hole Section' for hs in hole_sizes])
    section_totals = {}
    all_calc_dfs_for_excel = []

    # Special cases mapping (from previous logic)
    special_cases_map = {
        "PEX-AIT (150DegC Maximum)": ["AU14: AUX_SURELOC","NE1: NEUT_THER","DE1: DENS_FULL","RE1: RES_INDU"],
        "DSI-Dual OBMI (150DegC Maximum)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3","AU2: AUX_PCAL",
                                           "AU2: AUX_PCAL","PP7: PROC_PETR7","PA7: PROC_ACOU6","PA11: PROC_ACOU13","PA12: PROC_ACOU14",
                                           "IM3: IMAG_SOBM","PI1: PROC_IMAG1","PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8",
                                           "PI9: PROC_IMAG9","PI12: PROC_IMAG12","PI13: PROC_IMAG13"],
        "MDT Pretest and Sampling  (MDT-LFA-QS-XLD-MIFA-Saturn-2MS). 2ea SPMC, 6ea MPSR . 150DegC Maximum": [
            "AU14: AUX_SURELOC","FP25: FPS_SCAR","FP25: FPS_SCAR","FP18: FPS_SAMP","FP19: FPS_SPHA","FP23: FPS_TRA",
            "FP24: FPS_TRK","FP28: FPS_FCHA_1","FP33: FPS_FCHA_6","FP34: FPS_FCHA_7","FP14: FPS_PUMP","FP14: FPS_PUMP",
            "FP42: FPS_PROB_XLD","FP11: FPS_PROB_FO","FP26: FPS_FCON","DT3:RTDT_PER","PPT12: PROC_PT12"
        ],
        "XL Rock (150 DegC Maximum)": ["AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2"],
        "Pipe Conveyed Logging": ["CO1: CONV_PCL"],
        "FPIT & Back-off services / Drilling ontingent Support Services": ["AU7: AUX_SBOX","PC5: PC_10KH2S","PR1: PR_FP","PR2: PR_BO","PR3: PR_TP",
                                                                          "AU11: AUX_GRCCL","PR7: PR_CST","MS1: MS_PL","MS3: MS_JB"],
        "Unit, Cables & Conveyance": ["LU1: LUDR_ZON2","CA9: CABL_HSOH_1","CA3: CABL_HSOH","CA8: CABL_STCH_2","DT2:RTDT_SAT"],
        "Personnel": ["PER1:PWFE","PER2:PWSO","PER3:PWOP","PER4:PWSE"]
    }

    # Loop through each hole section
    for tab, hole_size in zip(tabs, hole_sizes):
        with tab:
            st.header(f'{hole_size}" Hole Section')

            # Sidebar inputs per section
            quantity_tools = st.sidebar.number_input(f"Quantity of Tools ({hole_size})", min_value=1,
                                                     value=reference_well_A["Quantity of Tools"], key=f"qty_{hole_size}")
            total_days = st.sidebar.number_input(f"Total Days ({hole_size})", min_value=0, value=0, key=f"days_{hole_size}")
            total_months = st.sidebar.number_input(f"Total Months ({hole_size})", min_value=0,
                                                   value=reference_well_A["Total Months"], key=f"months_{hole_size}")
            total_depth = st.sidebar.number_input(f"Total Depth (ft) ({hole_size})", min_value=0,
                                                  value=reference_well_A["Depths"].get(hole_size, 5500), key=f"depth_{hole_size}")
            total_survey = st.sidebar.number_input(f"Total Survey (ft) ({hole_size})", min_value=0, value=0, key=f"survey_{hole_size}")
            total_hours = st.sidebar.number_input(f"Total Hours ({hole_size})", min_value=0, value=0, key=f"hours_{hole_size}")
            discount = st.sidebar.number_input(f"Discount (%) ({hole_size})", min_value=0.0, max_value=100.0,
                                               value=0.0, key=f"disc_{hole_size}") / 100.0

            # Package & Service selection
            st.subheader("Select Package")
            package_options = df["Package"].dropna().unique().tolist()
            selected_package = st.selectbox("Choose Package", package_options, key=f"pkg_{hole_size}")
            package_df = df[df["Package"] == selected_package]

            st.subheader("Select Service Name")
            service_options = package_df["Service Name"].dropna().unique().tolist()
            selected_service = st.selectbox("Choose Service Name", service_options, key=f"svc_{hole_size}")
            df_service = package_df[package_df["Service Name"] == selected_service]

            # Tool selection
            code_list = df_service["Specification 1"].dropna().unique().tolist()
            code_list_with_special = list(special_cases_map.keys()) + code_list

            # Filter defaults to only those in options to prevent Streamlit error
            fixed_tools = reference_well_A["Hole Sections"].get(hole_size, {}).get("Tools", [])
            default_tools_filtered = [tool for tool in fixed_tools if tool in code_list_with_special]

            selected_codes = st.multiselect(
                "Select Tools (by Specification 1)",
                code_list_with_special,
                default=default_tools_filtered,
                key=f"tools_{hole_size}"
            )

            # Expand selected special cases
            expanded_codes = []
            for code in selected_codes:
                if code in special_cases_map:
                    expanded_codes.extend(special_cases_map[code])
                else:
                    expanded_codes.append(code)

            # --- Calculation ---
            calc_rows = []
            tracker = st.session_state.get("unique_tracker", set())
            for code in expanded_codes:
                row = df_service[df_service["Specification 1"] == code].iloc[0] if code in df_service["Specification 1"].values else pd.Series({"Daily Rate":0, "Monthly Rate":0, "Depth Charge (per ft)":0, "Flat Rate":0, "Hourly Charge":0})
                qty = quantity_tools
                # Unique tool handling
                if code in unique_tools:
                    if code in tracker:
                        qty = 0
                    else:
                        tracker.add(code)
                total_price = ((row.get("Depth Charge (per ft)",0)*total_depth) +
                               (row.get("Flat Rate",0)) +
                               (row.get("Hourly Charge",0)*total_hours) +
                               (row.get("Daily Rate",0)*total_days) +
                               (row.get("Monthly Rate",0)*total_months)) * (1-discount)
                calc_rows.append({
                    "Specification 1": code,
                    "Quantity": qty,
                    "Daily Rate": row.get("Daily Rate",0),
                    "Monthly Rate": row.get("Monthly Rate",0),
                    "Depth Charge (per ft)": row.get("Depth Charge (per ft)",0),
                    "Flat Rate": row.get("Flat Rate",0),
                    "Hourly Charge": row.get("Hourly Charge",0),
                    "Total Price": total_price
                })
            st.session_state["unique_tracker"] = tracker
            df_calc = pd.DataFrame(calc_rows)
            st.dataframe(df_calc)
            section_totals[hole_size] = df_calc["Total Price"].sum()
            all_calc_dfs_for_excel.append(df_calc)

    # Grand Total
    st.subheader("Grand Total")
    st.success(f"MYR {sum(section_totals.values()):,.2f}")

    # Excel Export
    if st.button("Download Cost Estimate Excel"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for hs, df_hs in zip(hole_sizes, all_calc_dfs_for_excel):
                df_hs.to_excel(writer, sheet_name=f"{hs}\" Section", index=False)
            writer.save()
        output.seek(0)
        st.download_button("Download Excel", data=output.getvalue(), file_name="wireline_estimate.xlsx")
