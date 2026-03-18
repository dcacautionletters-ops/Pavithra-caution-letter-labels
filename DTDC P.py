import streamlit as st
import pandas as pd
import io

# --- Helper Function: Clean Phone Numbers & IDs ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    # Remove .0 if it's a number stored as a float/decimal
    text = str(val).replace('.0', '').strip()
    return text

st.set_page_config(page_title="Student Label Generator", layout="wide")
st.title("🏷️ Student Label Generator (Excel Output)")

st.markdown("""
### Instructions:
1. **Caution File:** Upload the attendance/shortage report. (Code reads Roll No from **Column B**).
2. **Master Database:** Upload the student details file.
3. **Print:** The downloaded Excel is pre-set for **A4 paper** at **100% scale**.
""")

# --- FILE UPLOAD SECTION ---
col1, col2 = st.columns(2)

with col1:
    file_caution = st.file_uploader("Upload Caution Letter File", type=['xlsx', 'csv'], key="caution")
    skip_rows = st.number_input("Attendance data starts on row (usually 4):", min_value=1, value=4)

with col2:
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'], key="master")

if file_caution and file_master:
    try:
        # Load Caution/Attendance File
        if file_caution.name.endswith('csv'):
            df_c = pd.read_csv(file_caution, skiprows=skip_rows-1)
        else:
            df_c = pd.read_excel(file_caution, skiprows=skip_rows-1)
            
        # Load Master File
        if file_master.name.endswith('csv'):
            df_m = pd.read_csv(file_master)
        else:
            df_m = pd.read_excel(file_master)

        # 1. Match Roll No (Caution Col B [Index 1] -> Master 'ROLL NUMBER')
        # We clean them to ensure '25CG012.0' matches '25CG012'
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.replace('.0', '', regex=False).str.strip().unique()
        
        # Ensure Master column matches the naming in your file
        # If your master uses a different name, the code handles it via .get() below
        if 'ROLL NUMBER' in df_m.columns:
            df_m['ROLL NUMBER'] = df_m['ROLL NUMBER'].astype(str).str.replace('.0', '', regex=False).str.strip()
            df_matched = df_m[df_m['ROLL NUMBER'].isin(caution_rolls)]
        else:
            # Fallback if the column header is slightly different
            st.warning("Could not find 'ROLL NUMBER' column in Master. Please check headers.")
            df_matched = pd.DataFrame()

        if st.button("Generate Excel Label Sheet"):
            if not df_matched.empty:
                # Create an In-Memory Excel File
                output = io.BytesIO()
                workbook_writer = pd.ExcelWriter(output, engine='xlsxwriter')
                ws = workbook_writer.book.add_worksheet('PRINT_LABELS')

                # --- 2. FORMATTING SETUP (Strict A4 Alignment) ---
                # Column Widths (Left Label, Spacing, Right Label)
                ws.set_column('A:A', 46.5)
                ws.set_column('B:B', 2.5)
                ws.set_column('C:C', 46.5)

                # Label Cell Format (Calibri, Size 10, Bordered)
                label_format = workbook_writer.book.add_format({
                    'font_name': 'Calibri',
                    'font_size': 10,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'border': 1,
                    'border_color': '#C8C8C8' 
                })

                # --- 3. LOOP THROUGH DATA (2 per row) ---
                r_num = 0 
                # Converting to list of dicts for easier looping
                data_list = df_matched.to_dict('records')
                num_records = len(data_list)
                
                for i in range(0, num_records, 2):
                    # Set Row Height (Standard label height)
                    ws.set_row(r_num, 122)

                    # --- Left Label (Col A) ---
                    row_data = data_list[i]
                    # Clean the phone numbers to remove .0
                    s_phone = clean_val(row_data.get('STUDENT PHONE'))
                    p_phone = clean_val(row_data.get('PARENT PHONE'))
                    contact = f"{s_phone} / {p_phone}".strip(" / ")
                    
                    txt_left = (f"From: Presidency College Bangalore (AUTONOMOUS)\n"
                                f"To, Shri/Smt. {clean_val(row_data.get('FATHER'))}\n"
                                f"c/o: {clean_val(row_data.get('NAME'))}\n"
                                f"Address: {clean_val(row_data.get('ADDRESS'))}\n"
                                f"Contact: {contact}   ID: {clean_val(row_data.get('ROLL NUMBER'))}")
                    
                    ws.write(r_num, 0, txt_left, label_format)

                    # --- Right Label (Col C) ---
                    if i + 1 < num_records:
                        row_data_r = data_list[i+1]
                        s_phone_r = clean_val(row_data_r.get('STUDENT PHONE'))
                        p_phone_r = clean_val(row_data_r.get('PARENT PHONE'))
                        contact_r = f"{s_phone_r} / {p_phone_r}".strip(" / ")
                        
                        txt_right = (f"From: Presidency College Bangalore (AUTONOMOUS)\n"
                                     f"To, Shri/Smt. {clean_val(row_data_r.get('FATHER'))}\n"
                                     f"c/o: {clean_val(row_data_r.get('NAME'))}\n"
                                     f"Address: {clean_val(row_data_r.get('ADDRESS'))}\n"
                                     f"Contact: {contact_r}   ID: {clean_val(row_data_r.get('ROLL NUMBER'))}")
                        
                        ws.write(r_num, 2, txt_right, label_format)

                    # --- Spacing Row (Small gap between labels) ---
                    r_num += 1
                    ws.set_row(r_num, 8)
                    r_num += 1

                # --- 4. HARDCODED PRINT SETUP (For A4 Fit) ---
                ws.set_paper(9) # 9 = A4 (21 x 29.7 cm)
                ws.set_margins(left=0.2, right=0.2, top=0.3, bottom=0.3)
                ws.set_print_scale(100) # Force 100% scale, prevents shrinking
                ws.center_horizontally()

                workbook_writer.close()
                
                st.success(f"Generated {num_records} unique student labels.")
                st.download_button(
                    label="📥 Download Excel Labels",
                    data=output.getvalue(),
                    file_name="Student_Mailing_Labels.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No matches found. Check if Roll Numbers match in both files.")
                
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload both files to generate the label sheet.")
