import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill

st.title("SMARTLog: Wireline Cost Estimator")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# --- Define Reference Well A ---
reference_wells = {
    "Well A": {
        "Package": "Package A",
        "Service": "Standard Well",
        "Hole Sections": [
            {"size": 12.25, "quantity": 2, "total_months": 1, "total_depth": 5500, "special_tools": [
                "Well A PEX-AIT (150DegC Maximum)",
                "Well A DSI-Dual OBMI (150DegC Maximum)",
                "Well A MDT Pretest and Sampling",
                "Well A XL Rock (150DegC Maximum)"
            ]},
            {"size": 8.5, "quantity": 2, "total_months": 1, "total_depth": 8000, "special_tools": [
                "Well A PEX-AIT (150DegC Maximum)",
                "Well A DSI-Dual OBMI (150DegC Maximum)",
                "Well A MDT Pretest and Sampling",
                "Well A XL Rock (150DegC Maximum)"
            ]}
        ]
    }
}

# --- Reference Well select ---
ref_well_option = st.sidebar.selectbox("Reference Well", ["None"] + list(reference_wells.keys()))

if uploaded_file:
    # Reset unique tracker if new file
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

    # --- Determine number of sections ---
    if ref_well_option != "None":
        num_sections = len(reference_wells[ref_well_option]["Hole Sections"])
        hole_sizes = [str(s["size"]) for s in reference_wells[ref_well_option]["Hole Sections"]]
    else:
        st.sidebar.header("Hole Sections Setup")
        num_sections = st.sidebar.number_input("Number of Hole Sections", min_value=1, max_value=5, value=2, step=1)
        hole_sizes = []
        for i in range(num_sections):
            hole_size = st.sidebar.text_input(f"Hole Section {i+1} Size (inches)", value=f"{12.25 - i*3.75:.2f}")
            hole_sizes.append(hole_size)

    # Dynamic tabs
    tabs = st.tabs([f'{hs}" Hole Section' for hs in hole_sizes])
    section_totals = {}
    all_calc_dfs_for_excel = []

    # --- Loop for each hole section ---
    for i, (tab, hole_size) in enumerate(zip(tabs, hole_sizes)):
        with tab:
            st.header(f'{hole_size}" Hole Section')

            # --- Assign inputs ---
            if ref_well_option != "None":
                well_data = reference_wells[ref_well_option]["Hole Sections"][i]
                quantity_tools = well_data["quantity"]
                total_months = well_data["total_months"]
                total_depth = well_data["total_depth"]
                total_days = 0
                total_survey = 0
                total_hours = 0
                discount = 0.0
                selected_package = reference_wells[ref_well_option]["Package"]
                selected_service = reference_wells[ref_well_option]["Service"]
                selected_codes = well_data["special_tools"]
            else:
                # Sidebar inputs
                st.sidebar.subheader(f"Inputs for {hole_size}\" Section")
                quantity_tools = st.sidebar.number_input(f"Quantity of Tools ({hole_size})", min_value=1, value=2, key=f"qty_{hole_size}")
                total_days = st.sidebar.number_input(f"Total Days ({hole_size})", min_value=0, value=0, key=f"days_{hole_size}")
                total_months = st.sidebar.number_input(f"Total Months ({hole_size})", min_value=0, value=1, key=f"months_{hole_size}")
                total_depth = st.sidebar.number_input(f"Total Depth (ft) ({hole_size})", min_value=0, value=5500, key=f"depth_{hole_size}")
                total_survey = st.sidebar.number_input(f"Total Survey (ft) ({hole_size})", min_value=0, value=0, key=f"survey_{hole_size}")
                total_hours = st.sidebar.number_input(f"Total Hours ({hole_size})", min_value=0, value=0, key=f"hours_{hole_size}")
                discount = st.sidebar.number_input(f"Discount (%) ({hole_size})", min_value=0.0, max_value=100.0, value=0.0, key=f"disc_{hole_size}") / 100.0
                # Package & Service selection
                st.subheader("Select Package")
                package_options = df["Package"].dropna().unique().tolist()
                selected_package = st.selectbox("Choose Package", package_options, key=f"pkg_{hole_size}")

                st.subheader("Select Service Name")
                package_df = df[df["Package"] == selected_package]
                service_options = package_df["Service Name"].dropna().unique().tolist()
                selected_service = st.selectbox("Choose Service Name", service_options, key=f"svc_{hole_size}")

                # Tools selection
                df_service = package_df[package_df["Service Name"] == selected_service]
                code_list = df_service["Specification 1"].dropna().unique().tolist()
                selected_codes = st.multiselect("Select Tools (by Specification 1)", code_list, key=f"tools_{hole_size}")

            # --- Filter df_service for selected codes ---
            df_tools = df[df["Specification 1"].isin(selected_codes)].copy()

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

                all_calc_dfs_for_excel.append(calc_df)

    # --- Grand Total ---
    if section_totals:
        grand_total = sum(section_totals.values())
        st.success(f"🏆 Grand Total Price (MYR): {grand_total:,.2f}")

    # --- Excel Download ---
    if all_calc_dfs_for_excel:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for idx, df_sec in enumerate(all_calc_dfs_for_excel):
                df_sec.to_excel(writer, index=False, sheet_name=f"Section_{idx+1}")
        output.seek(0)
        st.download_button(
            "Download Cost Estimate Excel",
            data=output.getvalue(),
            file_name="SMARTLog_Calculation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
