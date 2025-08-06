import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="Spice Pesticide Processor")

st.title("Spice Pesticide Report Generator (Organic + Loose/Normal)")

uploaded_file = st.file_uploader("Upload Excel file with 13 spice sheets", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    st.success(f"Found {len(sheet_names)} sheets: {', '.join(sheet_names)}")

    with st.expander("Set Column Indices & Section Ranges"):
        selected_sheet = st.selectbox("Choose a sample sheet for preview", sheet_names)
        df_preview = xls.parse(selected_sheet, nrows=5)
        st.dataframe(df_preview)

        col_names = df_preview.columns.tolist()
        commodity_col = st.selectbox("Select 'Commodity' column", col_names)
        variant_col = st.selectbox("Select 'Variant' column", col_names)
        separator_col = st.selectbox("Select column that indicates start of Banned Pesticides", col_names)

        col_indices = {col: df_preview.columns.get_loc(col) for col in df_preview.columns}

        separator_index = col_indices[separator_col]
        offlabel_start = st.number_input("Off-label Start Column Index (0-based)", value=29)
        offlabel_end = separator_index - 1
        banned_start = separator_index + 1
        banned_end = len(col_names) - 1

    output_organic = pd.ExcelWriter("organic_output.xlsx", engine='xlsxwriter')
    output_loose = pd.ExcelWriter("loose_output.xlsx", engine='xlsxwriter')

    def process_data(df, spice_name, variant_filter, start_col, end_col, type_name):
        result_rows = []
        headers = list(df.columns)
        pesticide_indexes = {}
        
        # Make sure the end_col doesn't go beyond the available columns
        end_col = min(end_col, len(headers))
        
        for i in range(start_col, end_col, 3):
            pesticide_name = headers[i]
            if i + 1 >= len(headers):
                continue  # Skip if compliance column doesn't exist
            pesticide_indexes[pesticide_name] = {
                "valueIndex": i,
                "complianceIndex": i + 1
            }
        print("Headers:", headers)
        print("Start column:", start_col, "End column:", end_col)

        pesticide_data = {}
        for idx, row in df.iterrows():
            commodity = row[commodity_col]
            variant = row[variant_col]
            if isinstance(variant_filter, list):
                if variant not in variant_filter:
                    continue
            else:
                if variant != variant_filter:
                    continue

            for pest, idxes in pesticide_indexes.items():
                value = row[idxes["valueIndex"]]
                compliance = str(row[idxes["complianceIndex"]]).strip().lower()

                if pd.notna(value) and value != "":
                    try:
                        value = float(str(value).strip())
                    except (ValueError, TypeError):
                        continue  # skip if value is not numeric

                    if pest not in pesticide_data:
                        pesticide_data[pest] = {}
                    if commodity not in pesticide_data[pest]:
                        pesticide_data[pest][commodity] = {
                            "min": None, "max": None, "total": 0, "unsafe": 0
                        }

                    rec = pesticide_data[pest][commodity]
                    rec["total"] += 1
                    if compliance == "unsafe":
                        rec["unsafe"] += 1
                        if rec["min"] is None or value < rec["min"]:
                            rec["min"] = value
                        if rec["max"] is None or value > rec["max"]:
                            rec["max"] = value

        results = [["S. No", type_name + " Pesticide Residues", "Name of Spice",
                    "Min Amount (mg/kg)", "Max Amount (mg/kg)", "No. of unsafe", "Total Samples", "% Unsafe"]]

        sn = 1
        for pest, commodities in pesticide_data.items():
            for commodity, rec in commodities.items():
                percent = (rec["unsafe"] / rec["total"] * 100) if rec["total"] > 0 else 0
                results.append([
                    sn, pest, commodity,
                    rec["min"] if rec["min"] is not None else "No Residue",
                    rec["max"] if rec["max"] is not None else "No Residue",
                    rec["unsafe"], rec["total"], f"{percent:.2f}%"
                ])
                sn += 1
        return pd.DataFrame(results[1:], columns=results[0])

    st.write("### Click below to process:")
    if st.button("Generate Reports"):
        for sheet in sheet_names:
            df = xls.parse(sheet)
            # ORGANIC
            organic_off = process_data(df, sheet, "Organic", offlabel_start, offlabel_end, "Off-label Organic")
            organic_off.to_excel(output_organic, sheet_name=f"Off-Label {sheet}", index=False)

            organic_ban = process_data(df, sheet, "Organic", banned_start, banned_end, "Banned Organic")
            organic_ban.to_excel(output_organic, sheet_name=f"Banned {sheet}", index=False)

            # NORMAL+LOOSE
            normal_off = process_data(df, sheet, ["Normal", "Loose"], offlabel_start, offlabel_end, "Off-label")
            normal_off.to_excel(output_loose, sheet_name=f"Off-Label {sheet}", index=False)

            normal_ban = process_data(df, sheet, ["Normal", "Loose"], banned_start, banned_end, "Banned")
            normal_ban.to_excel(output_loose, sheet_name=f"Banned {sheet}", index=False)

        output_organic.close()
        output_loose.close()

        with open("organic_output.xlsx", "rb") as f1, open("loose_output.xlsx", "rb") as f2:
            st.success("✅ Processing Completed!")
            st.download_button("📥 Download Organic Report", f1, "organic_output.xlsx")
            st.download_button("📥 Download Loose + Normal Report", f2, "loose_output.xlsx")
