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

    # --- Well A Integration ---
    st.sidebar.header("Well Selection")
    well_option = st.sidebar.selectbox("Select Well", ["None", "Well A"], key="well_select")

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

    # Prefill parameters if Well A is selected
    if well_option == "Well A":
        # Well A defaults per hole section
        well_a_defaults_per_section = {
            '12.25': {"Package": "Package A", "Service": "Standard Wells", "qty": 2, "months": 1, "depth": 5500},
            '8.5': {"Package": "Package A", "Service": "Standard Wells", "qty": 2, "months": 1, "depth": 8000}
        }
        
        for hole_size in hole_sizes:
            key_size = str(float(hole_size))
            defaults = well_a_defaults_per_section.get(key_size, {})
            if defaults:
                st.session_state[f"qty_{hole_size}"] = defaults["qty"]
                st.session_state[f"months_{hole_size}"] = defaults["months"]
                st.session_state[f"depth_{hole_size}"] = defaults["depth"]
                # Also prefill Package and Service
                st.session_state[f"pkg_{hole_size}"] = defaults["Package"]
                st.session_state[f"svc_{hole_size}"] = defaults["Service"]
                
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
            quantity_tools = st.sidebar.number_input(f"Quantity of Tools ({hole_size})", min_value=1, value=st.session_state.get(f"qty_{hole_size}",2), key=f"qty_{hole_size}")
            total_days = st.sidebar.number_input(f"Total Days ({hole_size})", min_value=0, value=st.session_state.get(f"days_{hole_size}",0), key=f"days_{hole_size}")
            total_months = st.sidebar.number_input(f"Total Months ({hole_size})", min_value=0, value=st.session_state.get(f"months_{hole_size}",1), key=f"months_{hole_size}")
            total_depth = st.sidebar.number_input(f"Total Depth (ft) ({hole_size})", min_value=0, value=st.session_state.get(f"depth_{hole_size}",5500), key=f"depth_{hole_size}")
            total_survey = st.sidebar.number_input(f"Total Survey (ft) ({hole_size})", min_value=0, value=st.session_state.get(f"survey_{hole_size}",0), key=f"survey_{hole_size}")
            total_hours = st.sidebar.number_input(f"Total Hours ({hole_size})", min_value=0, value=st.session_state.get(f"hours_{hole_size}",0), key=f"hours_{hole_size}")
            discount = st.sidebar.number_input(f"Discount (%) ({hole_size})", min_value=0.0, max_value=100.0, value=st.session_state.get(f"disc_{hole_size}",0.0)*100, key=f"disc_{hole_size}") / 100.0

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
                    "IBC (PowerFlex)-CBL (150DegC Max)": ["CE1:CES_CBL","CE4:CES_CBI_3","CE6:CES_CBI_5", "DT3:RTDT_PER", "PPT13:PROC_PT13", "DT12:USI-DIG-LP-CET3"],
                    "DSI-QuantaGeo-Rt Scanner (150DegC Max)": ["AU14: AUX_SURELOC","GR1: GR_TOTL","AU3: AUX_INCL", "AC4: ACOU_ADD1", "AC3: ACOU_3", "AU2:AUX_PCAL",
                                                              "AU2:AUX_PCAL", "IM4:IMAG_ADD1","PI1:PROC_IMAG1", "DT4:SONIC-WELL-P/S-DIG", "PI2: PROC_IMAG2", "PI7: PROC_IMAG7",
                                                              "PI8:PROC_IMAG8", "PI9:PROC_IMAG9", "PI12: PROC_IMAG12", "PI13: PROC_IMAG13", "RE4: RES_ANIS"],
                    "XL Rock (150DegC Max)": ["AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2"],
                    "XL Rock (150DegC Max) With Core Detection": ["AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2", "SC4: SC_ADD4"]
                },
                "HT WELLS": {}
            }

            special_cases = special_cases_map.get(selected_service, {})
            code_list_with_special = list(special_cases.keys()) + code_list
            selected_codes = st.multiselect("Select Tools (by Specification 1)", code_list_with_special, key=f"tools_{hole_size}")

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
            if not df_tools.empty:
                display_rows = []
                inserted_dividers = set()
                for sc in used_special_cases:
                    divider = pd.DataFrame({col: "" for col in df_tools.columns}, index=[0])
                    divider["Specification 1"] = f"--- {sc} ---"
                    display_rows.append(divider)
                    for item in special_cases[sc]:
                        item_rows = df_tools[df_tools["Specification 1"] == item]
                        display_rows.extend([row.to_frame().T for _, row in item_rows.iterrows()])

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
                calc_df["Total Depth (ft)"] = total_depth
                calc_df["Total Survey (ft)"] = total_survey
                calc_df["Total Hours"] = total_hours
                calc_df["Discount (%)"] = discount * 100
                calc_df["Operating Charge (MYR)"] = (
                    (calc_df["Depth Charge (per ft)"] * total_depth) +
                    (calc_df["Survey Charge (per ft)"] * total_survey) +
                    (calc_df["Flat Rate"] * calc_df["User Flat Charge"]) +
                    (calc_df["Hourly Charge"] * total_hours)
                ) * (1 - discount)
                calc_df["Rental Charge (MYR)"] = (
                    calc_df["Quantity of Tools"] *
                    ((calc_df["Daily Rate"] * total_days) + (calc_df["Monthly Rate"] * total_months))
                ) * (1 - discount)

                # Unique-tool duplication logic
                calc_df["Is Duplicate Unique Tool"] = False
                calc_df["Status"] = "Charged"
                tracker = set(st.session_state.get("unique_tracker", set()))
                for ut in unique_tools:
                    mask = calc_df["Code"] == ut
                    if mask.any():
                        if ut in tracker:
                            calc_df.loc[mask, "Is Duplicate Unique Tool"] = True
                            calc_df.loc[mask, ["Operating Charge (MYR)", "Rental Charge (MYR)"]] = 0
                            calc_df.loc[mask, "Status"] = "Duplicate — Not charged"
                        else:
                            idxs = calc_df[mask].index.tolist()
                            for i in idxs[1:]:
                                calc_df.loc[i, "Is Duplicate Unique Tool"] = True
                                calc_df.loc[i, ["Operating Charge (MYR)", "Rental Charge (MYR)"]] = 0
                                calc_df.loc[i, "Status"] = "Duplicate — Not charged"
                            tracker.add(ut)
                st.session_state["unique_tracker"] = tracker
                calc_df["Total (MYR)"] = calc_df["Operating Charge (MYR)"] + calc_df["Rental Charge (MYR)"]
                cols = list(calc_df.columns)
                if "Status" in cols:
                    cols = ["Status"] + [c for c in cols if c != "Status"]
                st.subheader(f"Calculated Costs - Package {selected_package}, Service {selected_service}")
                st.dataframe(calc_df[cols])
                section_total = calc_df["Total (MYR)"].sum()
                section_totals[hole_size] = section_total
                st.write(f"### 💵 Section Total for {hole_size}\" Hole: {section_total:,.2f}")

                # Store for Excel download
                all_calc_dfs_for_excel.append((hole_size, used_special_cases, df_tools, special_cases))

    # --- Grand Total ---
    if section_totals:
        grand_total = sum(section_totals.values())
        st.success(f"🏆 Grand Total Price (MYR): {grand_total:,.2f}")

# --- Excel Download ---
if st.button("Download Cost Estimate Excel"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for hole_size, used_special_cases, df_tools_section, special_cases_section in all_calc_dfs_for_excel:
            sheet_name = f'{hole_size}" Hole'
            wb = writer.book
            ws = wb.create_sheet(title=sheet_name)
# --- Header rows ---
            ws.merge_cells("B2:B4"); ws["B2"]="Reference"
            ws.merge_cells("C2:C4"); ws["C2"]="Specification 1"
            ws.merge_cells("D2:D4"); ws["D2"]="Specification 2"
            ws.merge_cells("E2:J2"); ws["E2"]="Unit Price"
            ws.merge_cells("E3:F3"); ws["E3"]="Rental Price"
            ws.merge_cells("G3:J3"); ws["G3"]="Operating Charge"
            ws["E4"]="Daily Rate"; ws["F4"]="Monthly Rate"; ws["G4"]="Depth Charge (per ft)"
            ws["H4"]="Survey Charge (per ft)"; ws["I4"]="Flat Charge"; ws["J4"]="Hourly Charge"

            ws.merge_cells("K2:Q2"); ws["K2"] = "Operation Estimated"
            ws.merge_cells("K3:K4"); ws["K3"] = "Quantity of Tools"
            ws.merge_cells("L3:M3"); ws["L3"] = "Rental Parameters"
            ws["L4"] = "Total Days"; ws["M4"] = "Total Months"
            ws.merge_cells("N3:Q3"); ws["N3"] = "Operating Parameters"
            ws["N4"] = "Total Depth (ft)"; ws["O4"] = "Total Survey (ft)"
            ws["P4"] = "Total Flat Charge (ft)"; ws["Q4"] = "Total Hours"
            ws.merge_cells("R2:R4"); ws["R2"] = "Discount (%)"

            ws.merge_cells("S2:S4"); ws["S2"] = "Total (MYR)"
            ws.merge_cells("T2:T4"); ws["T2"] = "Grand Total Price (MYR)"
            ws.merge_cells("U2:U3"); ws["U2"] = "Break Down"; ws["U4"] = "Rental Charge (MYR)"
            ws.merge_cells("V2:V3"); ws["V2"] = "Break Down"; ws["V4"] = "Operating Charge (MYR)"

            # --- Apply colors ---
            white_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            light_green_fill = PatternFill(start_color="CCCC99", end_color="CCCC99", fill_type="solid")
            blue_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

            # White: Reference, Specification 1 & 2
            for cell in ["B2","B3","B4","C2","C3","C4","D2","D3","D4","R2","R3","R4","S2","S3","S4","T2","T3","T4","U2","U3","V2","V3"]:
                ws[cell].fill = white_fill
                ws[cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # Light Green: Unit Price / Rental / Operating columns
            for cell in ["E2","E3","E4","F2","F3","F4","G2","G3","G4","H4","I4","J4"]:
                ws[cell].fill = light_green_fill
                ws[cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # Blue: Operation Estimated headers
            for cell in ["K2","K3","K4","L3","L4","M3","M4","N3","N4","O4","P4","Q4"]:
                ws[cell].fill = blue_fill
                ws[cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # U4 & V4: Rental/Operating charge under Break Down
            ws["U4"].fill = light_green_fill
            ws["V4"].fill = light_green_fill
            ws["U4"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws["V4"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # --- Continue inserting data as before ---
            current_row = 5
            first_data_row = current_row  # Track first row for Grand Total formula

            # Insert special tools
            for sc in used_special_cases:
                ws[f"B{current_row}"] = f"{hole_size}in Section: {sc}"
                ws[f"B{current_row}"].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                ws[f"B{current_row}"].alignment = Alignment(horizontal="center")
                current_row += 1
                for item in special_cases_section[sc]:
                    item_rows = df_tools_section[df_tools_section["Specification 1"] == item]
                    if not item_rows.empty:
                        item_row = item_rows.iloc[0]
                        ws[f"B{current_row}"] = item_row.get("Reference","")
                        ws[f"C{current_row}"] = item_row.get("Specification 1","")
                        ws[f"D{current_row}"] = item_row.get("Specification 2","")
                        ws[f"E{current_row}"] = item_row.get("Daily Rate",0)
                        ws[f"F{current_row}"] = item_row.get("Monthly Rate",0)
                        ws[f"G{current_row}"] = item_row.get("Depth Charge (per ft)",0)
                        ws[f"H{current_row}"] = item_row.get("Survey Charge (per ft)",0)
                        ws[f"I{current_row}"] = item_row.get("Flat Rate",0)
                        ws[f"J{current_row}"] = item_row.get("Hourly Charge",0)

                        # Operation Estimated values
                        qty = st.session_state.get(f"qty_{hole_size}",0)
                        total_days = st.session_state.get(f"days_{hole_size}",0)
                        total_months = st.session_state.get(f"months_{hole_size}",0)
                        total_depth = st.session_state.get(f"depth_{hole_size}",0)
                        total_survey = st.session_state.get(f"survey_{hole_size}",0)
                        total_hours = st.session_state.get(f"hours_{hole_size}",0)
                        discount_pct = st.session_state.get(f"disc_{hole_size}",0)

                        ws[f"K{current_row}"] = qty
                        ws[f"L{current_row}"] = total_days
                        ws[f"M{current_row}"] = total_months
                        ws[f"N{current_row}"] = total_depth
                        ws[f"O{current_row}"] = total_survey
                        ws[f"P{current_row}"] = ws[f"I{current_row}"].value if ws[f"I{current_row}"].value else 0
                        ws[f"Q{current_row}"] = total_hours
                        ws[f"R{current_row}"] = discount_pct * 100

                        rental_charge = qty * ((item_row.get("Daily Rate",0)*total_days) + (item_row.get("Monthly Rate",0)*total_months))*(1-discount_pct)
                        operating_charge = ((item_row.get("Depth Charge (per ft)",0)*total_depth)+
                                            (item_row.get("Survey Charge (per ft)",0)*total_survey)+
                                            (item_row.get("Flat Rate",0))+
                                            (item_row.get("Hourly Charge",0)*total_hours))*(1-discount_pct)
                        total_myr = rental_charge + operating_charge

                        ws[f"S{current_row}"] = total_myr
                        ws[f"U{current_row}"] = rental_charge
                        ws[f"V{current_row}"] = operating_charge

                        current_row += 1

            # Insert non-special tools
            for item in df_tools_section["Specification 1"]:
                if item not in sum(special_cases_section.values(), []):
                    item_rows = df_tools_section[df_tools_section["Specification 1"] == item]
                    if not item_rows.empty:
                        item_row = item_rows.iloc[0]
                        ws[f"B{current_row}"] = item_row.get("Reference","")
                        ws[f"C{current_row}"] = item_row.get("Specification 1","")
                        ws[f"D{current_row}"] = item_row.get("Specification 2","")
                        ws[f"E{current_row}"] = item_row.get("Daily Rate",0)
                        ws[f"F{current_row}"] = item_row.get("Monthly Rate",0)
                        ws[f"G{current_row}"] = item_row.get("Depth Charge (per ft)",0)
                        ws[f"H{current_row}"] = item_row.get("Survey Charge (per ft)",0)
                        ws[f"I{current_row}"] = item_row.get("Flat Rate",0)
                        ws[f"J{current_row}"] = item_row.get("Hourly Charge",0)

                        qty = st.session_state.get(f"qty_{hole_size}",0)
                        total_days = st.session_state.get(f"days_{hole_size}",0)
                        total_months = st.session_state.get(f"months_{hole_size}",0)
                        total_depth = st.session_state.get(f"depth_{hole_size}",0)
                        total_survey = st.session_state.get(f"survey_{hole_size}",0)
                        total_hours = st.session_state.get(f"hours_{hole_size}",0)
                        discount_pct = st.session_state.get(f"disc_{hole_size}",0)

                        ws[f"K{current_row}"] = qty
                        ws[f"L{current_row}"] = total_days
                        ws[f"M{current_row}"] = total_months
                        ws[f"N{current_row}"] = total_depth
                        ws[f"O{current_row}"] = total_survey
                        ws[f"P{current_row}"] = ws[f"I{current_row}"].value if ws[f"I{current_row}"].value else 0
                        ws[f"Q{current_row}"] = total_hours
                        ws[f"R{current_row}"] = discount_pct * 100

                        rental_charge = qty * ((item_row.get("Daily Rate",0)*total_days) + (item_row.get("Monthly Rate",0)*total_months))*(1-discount_pct)
                        operating_charge = ((item_row.get("Depth Charge (per ft)",0)*total_depth)+
                                            (item_row.get("Survey Charge (per ft)",0)*total_survey)+
                                            (item_row.get("Flat Rate",0))+
                                            (item_row.get("Hourly Charge",0)*total_hours))*(1-discount_pct)
                        total_myr = rental_charge + operating_charge

                        ws[f"S{current_row}"] = total_myr
                        ws[f"U{current_row}"] = rental_charge
                        ws[f"V{current_row}"] = operating_charge

                        current_row += 1

            # Grand Total formula
            ws[f"T{first_data_row}"] = f"=SUM(S{first_data_row}:S{current_row-1})"
            ws[f"T{first_data_row}"].alignment = Alignment(horizontal="center")

    output.seek(0)
    st.download_button(
        "Download Cost Estimate Excel",
        data=output,
        file_name="Cost_Estimate.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



