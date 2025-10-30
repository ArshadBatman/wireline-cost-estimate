import streamlit as st
import pandas as pd

st.title("Cost Estimate Calculator")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    # --- Reset unique tracker when a new file is uploaded ---
    if "last_uploaded_name" not in st.session_state or st.session_state["last_uploaded_name"] != uploaded_file.name:
        st.session_state["unique_tracker"] = set()
        st.session_state["last_uploaded_name"] = uploaded_file.name

    # optional manual reset button
    if st.sidebar.button("Reset unique-tool usage (AU14, etc.)", key="reset_unique"):
        st.session_state["unique_tracker"] = set()
        st.sidebar.success("Unique-tool tracker cleared.")

    # read data
    df = pd.read_excel(uploaded_file, sheet_name="Data")

    # ensure columns
    if "Flat Rate" not in df.columns:
        df["Flat Rate"] = 0
    if "Depth Charge (per ft)" not in df.columns:
        df["Depth Charge (per ft)"] = 0
    if "Source" not in df.columns:
        df["Source"] = "Data"

    # Define unique tools that must be charged only once across all sections
    unique_tools = {"AU14: AUX_SURELOC"}  # set for O(1) membership checks

    # Tabs
    tab1, tab2 = st.tabs(['12.25" Hole Section', '8.5" Hole Section'])
    section_totals = {}

    # iterate sections in defined order — first tab processed first
    for tab, hole_size in zip([tab1, tab2], ["12.25", "8.5"]):
        with tab:
            st.header(f'{hole_size}" Hole Section')

            # --- Sidebar inputs (unique keys) ---
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
                    "PEX-AIT (150DegC Max)": [
                        "AU14: AUX_SURELOC",
                        "NE1: NEUT_THER",
                        "DE1: DENS_FULL",
                        "RE1: RES_INDU"
                    ],
                    "DOBMI (150DegC Max)": [
                        "AU14: AUX_SURELOC","GR1: GR_TOTL","AU3: AUX_INCL","AC3: ACOU_3",
                        "AU2: AUX_PCAL","AU2: AUX_PCAL","PP7: PROC_PETR7","PA7: PROC_ACOU6",
                        "PA11: PROC_ACOU13","PA12: PROC_ACOU14","IM3: IMAG_SOBM","PI1: PROC_IMAG1",
                        "PI2: PROC_IMAG2","PI7: PROC_IMAG7","PI8: PROC_IMAG8","PI9: PROC_IMAG9",
                        "PI12: PROC_IMAG12","PI13: PROC_IMAG13"
                    ],
                    "MDT: LFA-QS-XLD-MIFA-Saturn-2MS (150DegC Max)": [
                        "AU14: AUX_SURELOC","FP25: FPS_SCAR","FP25: FPS_SCAR","FP18: FPS_SAMP",
                        "FP19: FPS_SPHA","FP23: FPS_TRA","FP24: FPS_TRK","FP28: FPS_FCHA_1",
                        "FP33: FPS_FCHA_6","FP34: FPS_FCHA_7","FP14: FPS_PUMP","FP14: FPS_PUMP",
                        "FP42: FPS_PROB_LD","FP11: FPS_PROB_FO","FP26: FPS_FCON","DT3: RTDT_PER",
                        "PPT12: PROC_PT12", "FP7: FPS_SPPT_2"
                    ],
                    "XL Rock (150DegC Max)": [
                        "AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2"
                    ],
                    "XL Rock (150DegC Max) With Core Detection": [
                        "AU14: AUX_SURELOC","SC2: SC_ADD1","SC2: SC_ADD2", "SC4: SC_ADD4"
                    ], 
                },
                "HT WELLS": {}
            }

            special_cases = special_cases_map.get(selected_service, {})
            code_list_with_special = list(special_cases.keys()) + code_list

            selected_codes = st.multiselect("Select Tools (by Specification 1)", code_list_with_special, key=f"tools_{hole_size}")

            # expand selections
            expanded_codes = []
            for code in selected_codes:
                if code in special_cases:
                    expanded_codes.extend(special_cases[code])
                else:
                    expanded_codes.append(code)

            df_tools = df_service[df_service["Specification 1"].isin(expanded_codes)].copy()

            st.subheader(f"Selected Data - Package {selected_package}, Service {selected_service}")
            st.dataframe(df_tools)

            # --- Calculation ---
            if not df_tools.empty:
                calc_df = pd.DataFrame()
                calc_df["Source"] = df_tools["Source"]
                calc_df["Ref Item"] = df_tools["Reference"]
                calc_df["Code"] = df_tools["Specification 1"].astype(str).str.strip()   # normalize
                calc_df["Items"] = df_tools["Specification 2"]

                # Charges (numeric)
                calc_df["Daily Rate"] = pd.to_numeric(df_tools["Daily Rate"], errors="coerce").fillna(0)
                calc_df["Monthly Rate"] = pd.to_numeric(df_tools["Monthly Rate"], errors="coerce").fillna(0)
                calc_df["Depth Charge (per ft)"] = pd.to_numeric(df_tools["Depth Charge (per ft)"], errors="coerce").fillna(0)
                calc_df["Flat Rate"] = pd.to_numeric(df_tools["Flat Charge"], errors="coerce").fillna(0)

                # Other charges (defaults)
                calc_df["Survey Charge (per ft)"] = 0
                calc_df["Hourly Charge"] = 0

                # User-editable Flat Charge flag (so user can toggle later)
                calc_df["User Flat Charge"] = calc_df["Flat Rate"].apply(lambda x: 1 if x > 0 else 0)
                calc_df = st.data_editor(calc_df, num_rows="dynamic", key=f"editor_{hole_size}")

                # operation params
                calc_df["Quantity of Tools"] = quantity_tools
                calc_df["Total Days"] = total_days
                calc_df["Total Months"] = total_months
                calc_df["Total Depth (ft)"] = total_depth
                calc_df["Total Survey (ft)"] = total_survey
                calc_df["Total Hours"] = total_hours
                calc_df["Discount (%)"] = discount * 100

                # --- Compute operating & rental charges BEFORE duplicate zeroing ---
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

                # --- Detect and zero duplicates for unique_tools ---
                calc_df["Is Duplicate Unique Tool"] = False
                calc_df["Status"] = "Charged"

                tracker = set(st.session_state.get("unique_tracker", set()))

                for ut in unique_tools:
                    mask = calc_df["Code"] == ut
                    if mask.any():
                        if ut in tracker:
                            # already charged in prior section => all current occurrences are duplicates
                            calc_df.loc[mask, "Is Duplicate Unique Tool"] = True
                            calc_df.loc[mask, ["Operating Charge (MYR)", "Rental Charge (MYR)"]] = 0
                            calc_df.loc[mask, "Status"] = "Duplicate — Not charged"
                        else:
                            # present in this section and NOT seen before: charge only first occurrence here
                            idxs = calc_df[mask].index.tolist()
                            # mark subsequent occurrences in the same section as duplicates
                            for i in idxs[1:]:
                                calc_df.loc[i, "Is Duplicate Unique Tool"] = True
                                calc_df.loc[i, ["Operating Charge (MYR)", "Rental Charge (MYR)"]] = 0
                                calc_df.loc[i, "Status"] = "Duplicate — Not charged"
                            # record that this unique tool has now been charged (first occurrence)
                            tracker.add(ut)

                # persist tracker back to session_state
                st.session_state["unique_tracker"] = tracker

                # --- Final total per row and section ---
                calc_df["Total (MYR)"] = calc_df["Operating Charge (MYR)"] + calc_df["Rental Charge (MYR)"]

                # Show the calculated table with status so users can see why AU14 became free
                # reorder columns to show Status early
                cols = list(calc_df.columns)
                if "Status" in cols:
                    cols = ["Status"] + [c for c in cols if c != "Status"]
                st.subheader(f"Calculated Costs - Package {selected_package}, Service {selected_service}")
                st.dataframe(calc_df[cols])

                section_total = calc_df["Total (MYR)"].sum()
                section_totals[hole_size] = section_total

                st.write(f"### 💵 Section Total for {hole_size}\" Hole: {section_total:,.2f}")

    # --- Grand Total ---
    if section_totals:
        grand_total = sum(section_totals.values())
        st.success(f"🏆 Grand Total Price (MYR): {grand_total:,.2f}")




