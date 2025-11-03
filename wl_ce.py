import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill

st.title("SMARTLog: Wireline Cost Estimator")

# -------------------- Reference Well A Fixed Tools --------------------
reference_well_A = {
    "Package": "Package A",
    "Service Name": "STANDARD WELLS",
    "Hole Sections": {
        "12.25": {
            "Quantity of Tools": 2,
            "Total Months": 1,
            "Total Depth": 5500,
            "Tools": [
                # 1. PEX-AIT
                "PEX-AIT (150DegC Maximum)",
                "AU14: AUX_SURELOC", "NE1: NEUT_THER", "DE1: DENS_FULL", "RE1: RES_INDU",
                # 2. DSI-Dual OBMI
                "DSI-Dual OBMI (150DegC Maximum)",
                "AU14: AUX_SURELOC","GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3",
                "AU2: AUX_PCAL","AU2: AUX_PCAL","PP7: PROC_PETR7","PA7: PROC_ACOU6",
                "PA11: PROC_ACOU13","PA12: PROC_ACOU14","IM3: IMAG_SOBM","PI1: PROC_IMAG1",
                "PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8","PI9: PROC_IMAG9",
                "PI12: PROC_IMAG12","PI13: PROC_IMAG13",
                # 3. MDT Pretest & Sampling
                "MDT Pretest and Sampling (MDT-LFA-QS-XLD-MIFA-2MS)",
                "AU14: AUX_SURELOC","FP25: FPS_SCAR","FP25: FPS_SCAR","FP18: FPS_SAMP",
                "FP19: FPS_SPHA","FP23: FPS_TRA","FP24: FPS_TRK","FP28: FPS_FCHA_1",
                "FP33: FPS_FCHA_6","FP34: FPS_FCHA_7","FP14: FPS_PUMP","FP14: FPS_PUMP",
                "FP42: FPS_PROB_XLD","FP11: FPS_PROB_FO","FP26: FPS_FCON","DT3:RTDT_PER",
                "PPT12: PROC_PT12",
                # 4. XL Rock
                "XL Rock (150 DegC Maximum)","AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2",
                # Pipe Conveyed Logging
                "CO1: CONV_PCL",
                # FPIT & Back-off / Drilling contingent support
                "AU7: AUX_SBOX","PC5: PC_10KH2S","PR1: PR_FP","PR2: PR_BO","PR3: PR_TP",
                "AU11: AUX_GRCCL","PR7: PR_CST","MS1: MS_PL","MS3: MS_JB",
                # Unit, Cables & Conveyance
                "LU1: LUDR_ZON2","CA9: CABL_HSOH_1","CA3: CABL_HSOH","CA8: CABL_STCH_2","DT2:RTDT_SAT",
                # Personnel
                "PER1:PWFE","PER2:PWSO","PER3:PWOP","PER4:PWSE"
            ]
        },
        "8.5": {
            "Quantity of Tools": 2,
            "Total Months": 1,
            "Total Depth": 8000,
            "Tools": [  # Same as 12.25"
                "PEX-AIT (150DegC Maximum)",
                "AU14: AUX_SURELOC", "NE1: NEUT_THER", "DE1: DENS_FULL", "RE1: RES_INDU",
                "DSI-Dual OBMI (150DegC Maximum)",
                "AU14: AUX_SURELOC","GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3",
                "AU2: AUX_PCAL","AU2: AUX_PCAL","PP7: PROC_PETR7","PA7: PROC_ACOU6",
                "PA11: PROC_ACOU13","PA12: PROC_ACOU14","IM3: IMAG_SOBM","PI1: PROC_IMAG1",
                "PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8","PI9: PROC_IMAG9",
                "PI12: PROC_IMAG12","PI13: PROC_IMAG13",
                "MDT Pretest and Sampling (MDT-LFA-QS-XLD-MIFA-2MS)",
                "AU14: AUX_SURELOC","FP25: FPS_SCAR","FP25: FPS_SCAR","FP18: FPS_SAMP",
                "FP19: FPS_SPHA","FP23: FPS_TRA","FP24: FPS_TRK","FP28: FPS_FCHA_1",
                "FP33: FPS_FCHA_6","FP34: FPS_FCHA_7","FP14: FPS_PUMP","FP14: FPS_PUMP",
                "FP42: FPS_PROB_XLD","FP11: FPS_PROB_FO","FP26: FPS_FCON","DT3:RTDT_PER",
                "PPT12: PROC_PT12",
                "XL Rock (150 DegC Maximum)","AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2",
                "CO1: CONV_PCL",
                "AU7: AUX_SBOX","PC5: PC_10KH2S","PR1: PR_FP","PR2: PR_BO","PR3: PR_TP",
                "AU11: AUX_GRCCL","PR7: PR_CST","MS1: MS_PL","MS3: MS_JB",
                "LU1: LUDR_ZON2","CA9: CABL_HSOH_1","CA3: CABL_HSOH","CA8: CABL_STCH_2","DT2:RTDT_SAT",
                "PER1:PWFE","PER2:PWSO","PER3:PWOP","PER4:PWSE"
            ]
        }
    }
}

# -------------------- Upload Excel --------------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    # Reset unique tracker when a new file is uploaded
    if "last_uploaded_name" not in st.session_state or st.session_state["last_uploaded_name"] != uploaded_file.name:
        st.session_state["unique_tracker"] = set()
        st.session_state["last_uploaded_name"] = uploaded_file.name

    if st.sidebar.button("Reset unique-tool usage (AU14, etc.)", key="reset_unique"):
        st.session_state["unique_tracker"] = set()
        st.sidebar.success("Unique-tool tracker cleared.")

    # Read data
    df = pd.read_excel(uploaded_file, sheet_name="Data")
    for col in ["Flat Rate", "Depth Charge (per ft)", "Source"]:
        if col not in df.columns:
            df[col] = 0 if "Rate" in col else "Data"

    unique_tools = {"AU14: AUX_SURELOC"}

    # -------------------- Dynamic Hole Sections --------------------
    st.sidebar.header("Hole Sections Setup")
    num_sections = st.sidebar.number_input("Number of Hole Sections", min_value=1, max_value=5, value=2, step=1)
    hole_sizes = []
    for i in range(num_sections):
        hole_size = st.sidebar.text_input(f"Hole Section {i+1} Size (inches)", value=f"{12.25 - i*3.75:.2f}")
        hole_sizes.append(hole_size)

    tabs = st.tabs([f'{hs}" Hole Section' for hs in hole_sizes])
    section_totals = {}
    all_calc_dfs_for_excel = []

    # -------------------- Loop per Hole Section --------------------
    for tab, hole_size in zip(tabs, hole_sizes):
        with tab:
            st.header(f'{hole_size}" Hole Section')

            # Sidebar inputs
            default_qty = reference_well_A["Hole Sections"].get(hole_size, {}).get("Quantity of Tools", 2)
            default_months = reference_well_A["Hole Sections"].get(hole_size, {}).get("Total Months", 1)
            default_depth = reference_well_A["Hole Sections"].get(hole_size, {}).get("Total Depth", 5500)

            quantity_tools = st.sidebar.number_input(f"Quantity of Tools ({hole_size})", min_value=1, value=default_qty, key=f"qty_{hole_size}")
            total_days = st.sidebar.number_input(f"Total Days ({hole_size})", min_value=0, value=0, key=f"days_{hole_size}")
            total_months = st.sidebar.number_input(f"Total Months ({hole_size})", min_value=0, value=default_months, key=f"months_{hole_size}")
            total_depth = st.sidebar.number_input(f"Total Depth (ft) ({hole_size})", min_value=0, value=default_depth, key=f"depth_{hole_size}")
            total_survey = st.sidebar.number_input(f"Total Survey (ft) ({hole_size})", min_value=0, value=0, key=f"survey_{hole_size}")
            total_hours = st.sidebar.number_input(f"Total Hours ({hole_size})", min_value=0, value=0, key=f"hours_{hole_size}")
            discount = st.sidebar.number_input(f"Discount (%) ({hole_size})", min_value=0.0, max_value=100.0, value=0.0, key=f"disc_{hole_size}") / 100.0

            # Package & Service
            st.subheader("Select Package")
            package_options = df["Package"].dropna().unique().tolist()
            selected_package = st.selectbox("Choose Package", package_options, index=package_options.index(reference_well_A["Package"]) if reference_well_A["Package"] in package_options else 0, key=f"pkg_{hole_size}")

            st.subheader("Select Service Name")
            service_options = df[df["Package"]==selected_package]["Service Name"].dropna().unique().tolist()
            selected_service = st.selectbox("Choose Service Name", service_options, index=service_options.index(reference_well_A["Service Name"]) if reference_well_A["Service Name"] in service_options else 0, key=f"svc_{hole_size}")

            df_service = df[(df["Package"]==selected_package) & (df["Service Name"]==selected_service)]
            code_list = df_service["Specification 1"].dropna().unique().tolist()
            special_cases_map = {}  # Keep your existing special handling if any
            code_list_with_special = code_list + list(special_cases_map.keys())

            # -------------------- Prepopulate Tools --------------------
            fixed_tools = reference_well_A["Hole Sections"].get(hole_size, {}).get("Tools", [])
            selected_codes = st.multiselect(
                "Select Tools (by Specification 1)",
                code_list_with_special,
                default=fixed_tools,
                key=f"tools_{hole_size}"
            )

            # -------------------- Calculation --------------------
            calc_rows = []
            for code in selected_codes:
                row = df_service[df_service["Specification 1"] == code].iloc[0] if code in df_service["Specification 1"].values else pd.Series({"Unit Price": 0, "Daily Rate": 0})
                # Apply special calculations for unique tools (e.g., AU14) once
                qty = quantity_tools
                if code in unique_tools:
                    if code in st.session_state.get("unique_tracker", set()):
                        qty = 0
                    else:
                        st.session_state["unique_tracker"].add(code)
                total_price = row.get("Unit Price", 0) * qty + row.get("Depth Charge (per ft)", 0) * total_depth + row.get("Daily Rate", 0) * total_days + row.get("Monthly Rate", 0) * total_months
                total_price *= (1 - discount)
                calc_rows.append({
                    "Specification 1": code,
                    "Qty": qty,
                    "Unit Price": row.get("Unit Price", 0),
                    "Total Price": total_price
                })
            df_calc = pd.DataFrame(calc_rows)
            st.dataframe(df_calc)
            section_totals[hole_size] = df_calc["Total Price"].sum()
            all_calc_dfs_for_excel.append(df_calc)

    st.subheader("Grand Total")
    st.write(sum(section_totals.values()))

    # -------------------- Excel Export --------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for hs, df_hs in zip(hole_sizes, all_calc_dfs_for_excel):
            df_hs.to_excel(writer, sheet_name=f"{hs}\" Section", index=False)
        writer.save()
    st.download_button("Download Excel", data=output.getvalue(), file_name="wireline_estimate.xlsx")
