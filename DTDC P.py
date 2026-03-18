import streamlit as st
import pandas as pd
import io

# --- Helper Function: Clean Phone Numbers & IDs ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    return str(val).replace('.0', '').strip()

# --- Custom Sort Function ---
def get_sort_rank(roll):
    roll = str(roll).upper()
    if roll.startswith('25CG'): return 1
    if roll.startswith('25CAI'): return 2
    if roll.startswith('25CDS'): return 3
    if roll.startswith('24C'): return 4
    if roll.startswith('23C'): return 5
    return 6 # Everything else at the end

st.set_page_config(page_title="Student Label Generator", layout="wide")
st.title("🏷️ Student Label Generator (Chronological Sort)")

col1, col2 = st.columns(2)

with col1:
    file_caution = st.file_uploader("Upload Caution File", type=['xlsx', 'csv'])
    skip_rows = st.number_input("Attendance data starts on row:", min_value=1, value=4)

with col2:
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_caution and file_master:
    try:
        # Load Files
        df_c = pd.read_csv(file_caution, skiprows=skip_rows-1) if file_caution.name.endswith('csv') else pd.read_excel(file_caution, skiprows=skip_rows-1)
        df_m = pd.read_csv(file_master) if file_master.name.endswith('csv') else pd.read_excel(file_master)

        # 1. Match Roll No (Column B)
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.replace('.0', '', regex=False).str.strip().unique()
        df_m['ROLL NUMBER'] = df_m['ROLL NUMBER'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_matched = df_m[df_m['ROLL NUMBER'].isin(caution_rolls)].copy()

        if st.button("Generate Sorted Label Sheet"):
            if not df_matched.empty:
                # --- NEW: SORTING LOGIC ---
                # Create a temporary column for the priority rank
                df_matched['sort_rank'] = df_matched['ROLL NUMBER'].apply(get_sort_rank)
                # Sort by rank first, then by the Roll Number itself
                df_matched = df_matched.sort_values(by=['sort_rank', 'ROLL NUMBER'])

                # Create Excel
                output = io.BytesIO()
                workbook_writer = pd.ExcelWriter(output, engine='xlsxwriter')
                ws = workbook_writer.book.add_worksheet('PRINT_LABELS')

                # Column Widths & Format
                ws.set_column('A:A', 46.5)
                ws.set_column('B:B', 2.5)
                ws.set_column('C:C', 46.5)
                label_format = workbook_writer.book.add_format({
                    'font_name': 'Calibri', 'font_size': 10, 'text_wrap': True,
                    'valign': 'vcenter', 'border': 1, 'border_color': '#C8C8C8' 
                })

                # Loop Through Sorted Data
                r_num = 0 
                data_list = df_matched.to_dict('records')
                for i in range(0, len(data_list), 2):
                    ws.set_row(r_num, 122)
                    
                    # Left Label
                    row_data = data_list[i]
                    contact = f"{clean_val(row_data.get('STUDENT PHONE'))} / {clean_val(row_data.get('PARENT PHONE'))}".strip(" / ")
                    txt_left = (f"From: Presidency College Bangalore (AUTONOMOUS)\n"
                                f"To, Shri/Smt. {clean_val(row_data.get('FATHER'))}\n"
                                f"c/o: {clean_val(row_data.get('NAME'))}\n"
                                f"Address: {clean_val(row_data.get('ADDRESS'))}\n"
                                f"Contact: {contact}   ID: {clean_val(row_data.get('ROLL NUMBER'))}")
                    ws.write(r_num, 0, txt_left, label_format)

                    # Right Label
                    if i + 1 < len(data_list):
                        row_r = data_list[i+1]
                        contact_r = f"{clean_val(row_r.get('STUDENT PHONE'))} / {clean_val(row_r.get('PARENT PHONE'))}".strip(" / ")
                        txt_right = (f"From: Presidency College Bangalore (AUTONOMOUS)\n"
                                     f"To, Shri/Smt. {clean_val(row_r.get('FATHER'))}\n"
                                     f"c/o: {clean_val(row_r.get('NAME'))}\n"
                                     f"Address: {clean_val(row_r.get('ADDRESS'))}\n"
                                     f"Contact: {contact_r}   ID: {clean_val(row_r.get('ROLL NUMBER'))}")
                        ws.write(r_num, 2, txt_right, label_format)

                    r_num += 1
                    ws.set_row(r_num, 8)
                    r_num += 1

                # Page Setup
                ws.set_paper(9) # A4
                ws.set_margins(0.2, 0.2, 0.3, 0.3)
                ws.set_print_scale(100)
                ws.center_horizontally()
                workbook_writer.close()
                
                st.success(f"Labels sorted by series: 25CG > 25CAI > 25CDS > 24 > 23.")
                st.download_button(label="📥 Download Sorted Labels", data=output.getvalue(), file_name="Sorted_Labels.xlsx")
            else:
                st.error("No matches found.")
    except Exception as e:
        st.error(f"Error: {e}")
