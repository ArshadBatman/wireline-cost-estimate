import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

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

    all_calc_dfs_for_excel = []  # store all sections for Excel download

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
                    "XL Rock (150DegC Max)": ["AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2"],
                    "XL Rock (150DegC Max) With Core Detection": ["AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2", "SC4: SC_ADD4"]
                },
                "HT WELLS": {}
            }

            special_cases = special_cases_map.get(selected_service, {})
            code_list_with_special = list(special_cases.keys()) + code_list

            selected_codes = st.multiselect("Select Tools (by Specification 1)", code_list_with_special, key=f"tools_{hole_size}")

            # --- Expand special cases ---
            expanded_codes = []
            used_special_cases = []
            for code in selected_codes:
                if code in special_cases:
                    expanded_codes.extend(special_cases[code])
                    used_special_cases.append(code)
                else:
                    expanded_codes.append(code)

            df_tools = df_service[df_service["Specification 1"].isin(expanded_codes)].copy()

            # --- Row-by-row approach to insert dividers (for display) ---
            if not df_tools.empty:
                display_rows = []
                inserted_dividers = set()
                for sc in used_special_cases:
                    # insert divider
                    divider = pd.DataFrame({col: "" for col in df_tools.columns}, index=[0])
                    divider["Specification 1"] = f"--- {sc} ---"
                    display_rows.append(divider)
                    # append all line items for this special case
                    for item in special_cases[sc]:
                        row = df_tools[df_tools["Specification 1"] == item]
                        display_rows.append(row)
                    inserted_dividers.add(sc)

                # append non-special tools
                for item in df_tools["Specification 1"]:
                    if item not in sum(special_cases.values(), []):
                        display_rows.append(df_tools[df_tools["Specification 1"] == item])

                display_df = pd.concat(display_rows, ignore_index=True)

                # Style divider row
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

                # Charges
                calc_df["Daily Rate"] = pd.to_numeric(df_tools["Daily Rate"], errors="coerce").fillna(0)
                calc_df["Monthly Rate"] = pd.to_numeric(df_tools["Monthly Rate"], errors="coerce").fillna(0)
                calc_df["Depth Charge (per ft)"] = pd.to_numeric(df_tools["Depth Charge (per ft)"], errors="coerce").fillna(0)
                calc_df["Flat Rate"] = pd.to_numeric(df_tools["Flat Charge"], errors="coerce").fillna(0)
                calc_df["Survey Charge (per ft)"] = 0
                calc_df["Hourly Charge"] = 0

                calc_df["User Flat Charge"] = calc_df["Flat Rate"].apply(lambda x: 1 if x > 0 else 0)
                calc_df = st.data_editor(calc_df, num_rows="dynamic", key=f"editor_{hole_size}")

                # Operation params
                calc_df["Quantity of Tools"] = quantity_tools
                calc_df["Total Days"] = total_days
                calc_df["Total Months"] = total_months
                calc_df["Total Depth (ft)"] = total_depth
                calc_df["Total Survey (ft)"] = total_survey
                calc_df["Total Hours"] = total_hours
                calc_df["Discount (%)"] = discount * 100

                # Charges before duplicate handling
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
                section_total = calc_df["Total (MYR)"].sum()
                section_totals[hole_size] = section_total
                st.write(f"### 💵 Section Total for {hole_size}\" Hole: {section_total:,.2f}")

                all_calc_dfs_for_excel.append((hole_size, used_special_cases, df_tools.copy(), special_cases))

    # --- Grand Total ---
    if section_totals:
        grand_total = sum(section_totals.values())
        st.success(f"🏆 Grand Total Price (MYR): {grand_total:,.2f}")

    # --- Download Excel ---
    if st.button("Download Cost Estimate Excel"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for hole_size, used_special_cases, df_tools_section, special_cases_section in all_calc_dfs_for_excel:
                start_row = 0
                sheet_name = f'{hole_size}" Hole'
                ws_df = pd.DataFrame()  # empty to start sheet
                # add headers (merged manually in Excel if needed)
                df_to_write = pd.DataFrame(columns=["Reference","Specification 1","Specification 2",
                                                    "Daily Rate","Monthly Rate","Depth Charge (per ft)",
                                                    "Survey Charge (per ft)","Flat Charge","Hourly Charge"])
                row_idx = 0
                for sc in used_special_cases:
                    # special case divider as row
                    df_to_write.loc[row_idx] = [ "", f"--- {sc} ---", "", "", "", "", "", "", ""]
                    row_idx += 1
                    for item in special_cases_section[sc]:
                        item_row = df_tools_section[df_tools_section["Specification 1"] == item].iloc[0]
                        df_to_write.loc[row_idx] = [
                            item_row["Reference"], item_row["Specification 1"], item_row["Specification 2"],
                            item_row["Daily Rate"], item_row["Monthly Rate"], item_row["Depth Charge (per ft)"],
                            item_row["Survey Charge (per ft)"], item_row["Flat Rate"], item_row.get("Hourly Charge",0)
                        ]
                        row_idx += 1
                # non-special tools
                for item in df_tools_section["Specification 1"]:
                    if item not in sum(special_cases_section.values(), []):
                        item_row = df_tools_section[df_tools_section["Specification 1"] == item].iloc[0]
                        df_to_write.loc[row_idx] = [
                            item_row["Reference"], item_row["Specification 1"], item_row["Specification 2"],
                            item_row["Daily Rate"], item_row["Monthly Rate"], item_row["Depth Charge (per ft)"],
                            item_row["Survey Charge (per ft)"], item_row["Flat Rate"], item_row.get("Hourly Charge",0)
                        ]
                        row_idx += 1

                df_to_write.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

            writer.save()
        output.seek(0)
        st.download_button("Download Excel", data=output, file_name="Cost_Estimate.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
