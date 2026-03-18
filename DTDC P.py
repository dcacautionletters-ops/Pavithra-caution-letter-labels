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

# --- Custom Sorting Function for your specific Series ---
def get_sort_rank(roll):
    roll = str(roll).upper()
    if roll.startswith('25CG'): return 1
    if roll.startswith('25CAI'): return 2
    if roll.startswith('25CDS'): return 3
    if roll.startswith('24C'): return 4
    if roll.startswith('23C'): return 5
    return 6 # Everything else

st.set_page_config(page_title="Student Label Generator", layout="wide")
st.title("🏷️ Student Label Generator (Final Version)")

# --- SIDEBAR SETTINGS ---
st.sidebar.header("Global Settings")
from_address = st.sidebar.text_area(
    "Edit 'From' Address:", 
    value="Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024"
)
st.sidebar.info("The labels will be sorted: 25CG > 25CAI > 25CDS > 24 > 23")

# --- FILE UPLOAD SECTION ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Attendance Report")
    file_caution = st.file_uploader("Upload Attendance/Shortage File", type=['xlsx', 'csv'], key="caution")
    skip_rows = st.number_input("Attendance data starts on row:", min_value=1, value=4)

with col2:
    st.subheader("2. Master Database")
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'], key="master")

if file_caution and file_master:
    try:
        # Load Attendance File (Reads Col B, C, G)
        if file_caution.name.endswith('csv'):
            df_c = pd.read_csv(file_caution, skiprows=skip_rows-1)
        else:
            df_c = pd.read_excel(file_caution, skiprows=skip_rows-1)
            
        # Load Master File (Reads Col B, F, AD, S, AT, AS)
        if file_master.name.endswith('csv'):
            df_m = pd.read_csv(file_master)
        else:
            df_m = pd.read_excel(file_master)

        # 1. Match Roll No from Attendance (Column B - Index 1)
        # We clean and get unique rolls so 1 student = 1 label
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.replace('.0', '', regex=False).str.strip().unique()
        
        # 2. Extract Data from Master based on Roll No (Column B - Index 1)
        # B=1, F=5 (Name), AD=29 (Father), S=18 (Address), AT=45 (Father Ph), AS=44 (Student Ph)
        # We use iloc to ensure we hit the right columns regardless of header names
        mast_data = df_m.iloc[:, [1, 5, 29, 18, 45, 44]].copy()
        mast_data.columns = ['Roll_No', 'Name', 'Father', 'Address', 'Father_Phone', 'Student_Phone']
        
        # Clean Roll numbers for the merge
        mast_data['Roll_No'] = mast_data['Roll_No'].astype(str).str.replace('.0', '', regex=False).str.strip()
        
        # Filter master data to only include students in the caution list
        df_matched = mast_data[mast_data['Roll_No'].isin(caution_rolls)].copy()

        if st.button("Generate Sorted A4 Label Sheet"):
            if not df_matched.empty:
                
                # --- SORTING LOGIC ---
                df_matched['sort_rank'] = df_matched['Roll_No'].apply(get_sort_rank)
                df_matched = df_matched.sort_values(by=['sort_rank', 'Roll_No'])

                # Create Excel Workbook
                output = io.BytesIO()
                workbook_writer = pd.ExcelWriter(output, engine='xlsxwriter')
                ws = workbook_writer.book.add_worksheet('PRINT_LABELS')

                # --- 3. FORMATTING (A4 Alignment) ---
                ws.set_column('A:A', 46.5)
                ws.set_column('B:B', 2.5) # The gap column
                ws.set_column('C:C', 46.5)

                label_format = workbook_writer.book.add_format({
                    'font_name': 'Calibri',
                    'font_size': 10,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'border': 1,
                    'border_color': '#C8C8C8' 
                })

                # --- 4. LOOP THROUGH DATA (2 per row) ---
                r_num = 0 
                data_list = df_matched.to_dict('records')
                num_records = len(data_list)
                
                for i in range(0, num_records, 2):
                    ws.set_row(r_num, 122) # Height of label

                    # --- Left Label (Col A) ---
                    d = data_list[i]
                    f_ph = clean_val(d.get('Father_Phone'))
                    s_ph = clean_val(d.get('Student_Phone'))
                    contact = f"{f_ph} / {s_ph}".strip(" / ")
                    
                    txt_left = (f"From: {from_address}\n"
                                f"To, Shri/Smt. {clean_val(d.get('Father'))}\n"
                                f"c/o: {clean_val(d.get('Name'))}\n"
                                f"Address: {clean_val(d.get('Address'))}\n"
                                f"Contact: {contact}   ID: {clean_val(d.get('Roll_No'))}")
                    ws.write(r_num, 0, txt_left, label_format)

                    # --- Right Label (Col C) ---
                    if i + 1 < num_records:
                        dr = data_list[i+1]
                        f_ph_r = clean_val(dr.get('Father_Phone'))
                        s_ph_r = clean_val(dr.get('Student_Phone'))
                        contact_r = f"{f_ph_r} / {s_ph_r}".strip(" / ")
                        
                        txt_right = (f"From: {from_address}\n"
                                     f"To, Shri/Smt. {clean_val(dr.get('Father'))}\n"
                                     f"c/o: {clean_val(dr.get('Name'))}\n"
                                     f"Address: {clean_val(dr.get('Address'))}\n"
                                     f"Contact: {contact_r}   ID: {clean_val(dr.get('Roll_No'))}")
                        ws.write(r_num, 2, txt_right, label_format)

                    # --- Vertical Gap Row ---
                    r_num += 1
                    ws.set_row(r_num, 8)
                    r_num += 1

                # --- 5. PAGE SETUP (A4 FIX) ---
                ws.set_paper(9) # Force A4
                ws.set_margins(left=0.2, right=0.2, top=0.3, bottom=0.3)
                ws.set_print_scale(100) # Force 100% (No shrinking)
                ws.center_horizontally()

                workbook_writer.close()
                
                st.success(f"Generated {num_records} labels sorted by Roll No series.")
                st.download_button(
                    label="📥 Download Sorted Label Sheet",
                    data=output.getvalue(),
                    file_name="Student_Sorted_Labels.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No matches found between Attendance and Master Database.")
                
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload both files to start.")
