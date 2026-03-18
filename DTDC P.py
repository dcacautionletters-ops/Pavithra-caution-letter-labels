import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Caution Letter Generator", layout="wide")

st.title("Caution Letter Data Generator ✉️")
st.markdown("One row per student | Clean Phone Numbers | **Fixed A4 Printing**")

def clean_phone(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""
    phone_str = str(value).replace('.0', '').strip()
    return phone_str

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Attendance Report")
    shortage_file = st.file_uploader("Upload Attendance", type=['xlsx', 'csv'], key="att")
    skip_rows = st.number_input("Starts on row:", min_value=1, value=4)

with col2:
    st.subheader("2. Master Database")
    master_file = st.file_uploader("Upload Master", type=['xlsx', 'csv'], key="mast")

if shortage_file and master_file:
    try:
        df_att_raw = pd.read_excel(shortage_file, skiprows=skip_rows-1) if shortage_file.name.endswith('.xlsx') else pd.read_csv(shortage_file, skiprows=skip_rows-1)
        df_mast_raw = pd.read_excel(master_file) if master_file.name.endswith('.xlsx') else pd.read_csv(master_file)

        if st.button("Generate Final Report"):
            # 1. Extraction & De-duplication
            att_subset = df_att_raw.iloc[:, [1, 2, 6]].copy()
            att_subset.columns = ['Roll_No', 'Student_Name', 'Batch']
            att_subset = att_subset.drop_duplicates(subset=['Roll_No'], keep='first')

            # 2. Extraction Master
            mast_subset = df_mast_raw.iloc[:, [1, 29, 18, 45, 44]].copy()
            mast_subset.columns = ['Roll_No', 'Father_Name', 'Address', 'Father_Phone', 'Student_Phone']

            att_subset['Roll_No'] = att_subset['Roll_No'].astype(str).str.strip()
            mast_subset['Roll_No'] = mast_subset['Roll_No'].astype(str).str.strip()

            # 3. Merge
            merged_df = pd.merge(att_subset, mast_subset, on='Roll_No', how='left')

            # 4. Final Formatting
            final_report = pd.DataFrame()
            final_report['Sl No'] = range(1, len(merged_df) + 1)
            final_report['Batch'] = merged_df['Batch']
            final_report['To name'] = merged_df['Father_Name']
            final_report['Tracking ID'] = "" 
            final_report['Roll no'] = merged_df['Roll_No']
            final_report['Student name'] = merged_df['Student_Name']
            final_report['Complete Address'] = merged_df['Address']
            final_report['Father Number'] = merged_df['Father_Phone'].apply(clean_phone)
            final_report['Student Number'] = merged_df['Student_Phone'].apply(clean_phone)

            # 5. EXCEL EXPORT WITH PRINT SETTINGS
            buffer = io.BytesIO()
            # We use xlsxwriter to control the page layout
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                final_report.to_excel(writer, index=False, sheet_name='Caution_List')
                
                workbook  = writer.book
                worksheet = writer.sheets['Caution_List']

                # --- FORCING PRINTER SETTINGS ---
                worksheet.set_paper(9)  # 9 is the code for A4 (21x29.7cm)
                worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)
                worksheet.set_print_scale(100) # Prevents "Scale to Fit" distortion
                
            excel_data = buffer.getvalue()

            st.success(f"Success! {len(final_report)} students processed.")
            st.dataframe(final_report, hide_index=True)

            st.download_button(
                label="Download Final Excel Report",
                data=excel_data,
                file_name="Caution_Letters_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload files to generate the report.")