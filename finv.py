import streamlit as st
import pandas as pd
import numpy as np
import time
import re
import io

# ==========================================
# 🛠️ FIXED EXCEL COL-WIDTH FORMATTING FUNCTION
# ==========================================
def convert_to_styled_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
        workbook = writer.book
        worksheet = writer.sheets['Report']
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#5E239D', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        cell_format = workbook.add_format({'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            col_series = df[value].fillna("").astype(str).map(str)
            max_len = max(col_series.map(len).max(), len(str(value)))
            column_len = max_len + 2
            worksheet.set_column(col_num, col_num, min(column_len, 50), cell_format)
    return output.getvalue()

# --- PAGE CONFIGURATION & PREMIUM CSS ---
st.set_page_config(page_title="PhonePe FnF Pro", page_icon="📱", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #F4F6F9; }
    @keyframes gradientBG { 0% {background-position: 0% 50%;} 50% {background-position: 100% 50%;} 100% {background-position: 0% 50%;} }
    .main-header {
        background: linear-gradient(-45deg, #5E239D, #9D4EDD, #FF007F, #5E239D);
        background-size: 300% 300%; animation: gradientBG 8s ease infinite;
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        text-align: center; font-size: 3.5rem; font-weight: 900; padding-top: 10px;
    }
    .sub-header { text-align: center; color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom: 30px; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E5E7EB; box-shadow: 2px 0 15px rgba(0,0,0,0.05); }
    div.stButton > button { border-radius: 10px !important; font-weight: 700 !important; background-color: #5E239D !important; color: white !important; transition: transform 0.2s; }
    div.stButton > button:hover { transform: scale(1.02); }
    .feature-card { background: white; padding: 25px; border-radius: 12px; border-top: 4px solid #5E239D; box-shadow: 0 4px 6px rgba(0,0,0,0.05); text-align: center; height: 100%; transition: 0.3s; }
    .feature-card:hover { box-shadow: 0 8px 15px rgba(0,0,0,0.1); }
    .alert-popup { background: linear-gradient(135deg, #FFF1F2 0%, #FFE4E6 100%); border: 2px solid #F43F5E; border-radius: 12px; padding: 25px; margin: 20px 0; text-align: center; animation: pulse-red 2s infinite; }
    @keyframes pulse-red { 0% { box-shadow: 0 0 0 0 rgba(244, 63, 94, 0.4); } 70% { box-shadow: 0 0 0 15px rgba(244, 63, 94, 0); } 100% { box-shadow: 0 0 0 0 rgba(244, 63, 94, 0); } }
    </style>
""", unsafe_allow_html=True)

# --- SESSION STATE ---
if 'current_page' not in st.session_state: st.session_state.current_page = "Master"
if 'absconding_decision' not in st.session_state: st.session_state.absconding_decision = 'pending'

def nav_home(): st.session_state.current_page = "Home"
def nav_master(): st.session_state.current_page = "Master"

# --- HELPERS ---
def load_file(uploaded_file, sheet_name=None):
    uploaded_file.seek(0)
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file, encoding='latin1', on_bad_lines='skip')
    if sheet_name:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name)
    return pd.read_excel(uploaded_file)

def standardize_id(df, possible_names, report_name):
    def scrub(val): return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower()
    clean_possible = [scrub(p) for p in possible_names]
    id_indices = [i for i, h in enumerate(df.columns) if scrub(h) in clean_possible]
    if not id_indices:
        for i, row in df.head(50).iterrows():
            found_indices = [idx for idx, val in enumerate(row.values) if scrub(val) in clean_possible]
            if found_indices:
                df.columns = [str(v).strip() for v in row.values]
                df = df.iloc[i+1:].reset_index(drop=True)
                id_indices = found_indices
                break
    if id_indices:
        df = df.copy()
        df['Base_ID'] = df.iloc[:, id_indices[0]].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        return df.dropna(subset=['Base_ID']).reset_index(drop=True)
    else:
        st.error(f"🚨 Could not locate ID in {report_name}"); st.stop()

def lookup_column_index(columns_list, include_tokens, exclude_tokens=None):
    for idx, col in enumerate(columns_list):
        if all(t.lower() in col for t in include_tokens):
            if exclude_tokens and any(e.lower() in col for e in exclude_tokens):
                continue
            return idx
    return None

def parse_numeric_cell(df_row, columns_lower, tokens, skip_tokens=None):
    if df_row.empty: return 0.0
    idx = lookup_column_index(columns_lower, tokens, skip_tokens)
    if idx is not None:
        cell_val = df_row.iloc[0, idx]
        return pd.to_numeric(cell_val, errors='coerce') if pd.notna(cell_val) else 0.0
    return 0.0

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h2 style='color:#5E239D; text-align:center;'>🧭 Workspace</h2>", unsafe_allow_html=True)
    st.button("🏠 Home", on_click=nav_home, use_container_width=True)
    st.button("📂 Master Builder", on_click=nav_master, use_container_width=True)

st.markdown("<div class='main-header'>PhonePe FnF Pro</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>For Non piece rate</div>", unsafe_allow_html=True)

# ==========================================
# HOME PAGE MODULE
# ==========================================
if st.session_state.current_page == "Home":
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='feature-card' style='max-width: 600px; margin: 0 auto;'><h3>📂 Consolidator</h3><p>Build Master FnF Reports from HC, Exit, FFS, and Relocation metrics.</p></div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.button("Launch Consolidator", on_click=nav_master, use_container_width=True, key="home_master")

# ==========================================
# MODULE 2: MASTER BUILDER (NON PIECE RATE)
# ==========================================
elif st.session_state.current_page == "Master":
    st.markdown("### 📂 Step 2: Master Report Builder")
    
    f_emp_list = st.file_uploader("📥 Step A: Upload Master Employee List File (Finds EMP ID / emp id / Employee ID)", type=["xlsx", "csv"])
    
    if f_emp_list is not None:
        df_target_list = standardize_id(load_file(f_emp_list), ['EMP ID', 'emp id', 'Employee ID'], "Employee Filter List")
        allowed_ids = df_target_list['Base_ID'].astype(str).str.strip().tolist()
        
        st.success(f"🎉 Verification Complete! Identified {len(allowed_ids)} unique employees considered for the present cut.")
        st.markdown("---")
        
        st.markdown("#### 📊 Phase 2: Upload all the reports")
        col1, col2 = st.columns(2)
        with col1:
            f_hc = st.file_uploader("1. HC Report", type=["xlsx", "csv"])
            f_exit = st.file_uploader("2. Exit Report", type=["xlsx", "csv"])
            f_ffs = st.file_uploader("3. FFS Input Report", type=["xlsx", "csv"])
        with col2:
            f_rel_nj = st.file_uploader("4. Relocation Expenses NJ Report", type=["xlsx", "csv"])
            f_rel_ijp = st.file_uploader("5. IJP Relocation of goods Report", type=["xlsx", "csv"])

        if all([f_hc, f_exit, f_ffs, f_rel_nj, f_rel_ijp]):
            if st.button("🚀 Run Master Consolidation", use_container_width=True):
                with st.spinner("Consolidating data elements..."):
                    
                    df_hc_raw = standardize_id(load_file(f_hc), ['User/Employee ID', 'user/employee id'], "HC Report")
                    df_exit_raw = standardize_id(load_file(f_exit), ["Employee' User ID", "employee' user id"], "Exit Report")
                    df_ffs_raw = standardize_id(load_file(f_ffs), ['Employee ID', 'employee id'], "FFS Input Report")
                    df_nj_raw = standardize_id(load_file(f_rel_nj), ['Employee ID', 'emp id', 'id'], "Relocation NJ Report")
                    df_ijp_raw = standardize_id(load_file(f_rel_ijp), ['Employee ID', 'emp id', 'id'], "IJP Relocation Report")

                    hc_cols_lower = [str(c).strip().lower() for c in df_hc_raw.columns]
                    exit_cols_lower = [str(c).strip().lower() for c in df_exit_raw.columns]
                    ffs_cols_lower = [str(c).strip().lower() for c in df_ffs_raw.columns]
                    nj_cols_lower = [str(c).strip().lower() for c in df_nj_raw.columns]
                    ijp_cols_lower = [str(c).strip().lower() for c in df_ijp_raw.columns]

                    master_rows = []

                    for emp_id in allowed_ids:
                        row_data = {'Employee ID': emp_id}
                        
                        row_hc = df_hc_raw[df_hc_raw['Base_ID'].astype(str).str.strip() == emp_id]
                        row_exit = df_exit_raw[df_exit_raw['Base_ID'].astype(str).str.strip() == emp_id]
                        row_ffs = df_ffs_raw[df_ffs_raw['Base_ID'].astype(str).str.strip() == emp_id]
                        row_nj = df_nj_raw[df_nj_raw['Base_ID'].astype(str).str.strip() == emp_id]
                        row_ijp = df_ijp_raw[df_ijp_raw['Base_ID'].astype(str).str.strip() == emp_id]

                        d_min_doj = pd.NaT
                        d_res = pd.NaT
                        d_lwd_sf = pd.NaT

                        if not row_hc.empty:
                            idx_first = lookup_column_index(hc_cols_lower, ['first', 'name'])
                            idx_mid = lookup_column_index(hc_cols_lower, ['middle', 'name'])
                            idx_last = lookup_column_index(hc_cols_lower, ['last', 'name'])
                            idx_emptype = lookup_column_index(hc_cols_lower, ['employee', 'type'])
                            idx_reason = lookup_column_index(hc_cols_lower, ['event', 'reason'])
                            if idx_reason is None: idx_reason = lookup_column_index(hc_cols_lower, ['exit', 'reason'])
                            idx_entity = lookup_column_index(hc_cols_lower, ['company'])
                            
                            idx_pos = lookup_column_index(hc_cols_lower, ['position', 'title'])
                            if idx_pos is None: idx_pos = lookup_column_index(hc_cols_lower, ['job', 'title'])
                            
                            idx_dept = lookup_column_index(hc_cols_lower, ['department'])
                            idx_loc = lookup_column_index(hc_cols_lower, ['statutory', 'location'])
                            if idx_loc is None: idx_loc = lookup_column_index(hc_cols_lower, ['location'])

                            row_data['First Name'] = str(row_hc.iloc[0, idx_first]).strip() if idx_first is not None else ""
                            
                            if idx_mid is not None:
                                m_val = str(row_hc.iloc[0, idx_mid]).strip()
                                row_data['Middle Name'] = "" if m_val.lower() == "nan" or pd.isna(row_hc.iloc[0, idx_mid]) else m_val
                            else:
                                row_data['Middle Name'] = ""

                            row_data['Last Name'] = str(row_hc.iloc[0, idx_last]).strip() if idx_last is not None else ""
                            row_data['Employee Type'] = str(row_hc.iloc[0, idx_emptype]).strip() if idx_emptype is not None else ""
                            row_data['Reason'] = str(row_hc.iloc[0, idx_reason]).strip() if idx_reason is not None else ""
                            row_data['Entity'] = str(row_hc.iloc[0, idx_entity]).strip() if idx_entity is not None else ""
                            row_data['Position'] = str(row_hc.iloc[0, idx_pos]).strip() if idx_pos is not None else ""
                            row_data['Department'] = str(row_hc.iloc[0, idx_dept]).strip() if idx_dept is not None else ""
                            row_data['Location'] = str(row_hc.iloc[0, idx_loc]).strip() if idx_loc is not None else ""

                            idx_doj1 = lookup_column_index(hc_cols_lower, ['date', 'joining'], ['legal', 'group'])
                            idx_doj2 = lookup_column_index(hc_cols_lower, ['legal', 'entity', 'joining'])
                            idx_doj3 = lookup_column_index(hc_cols_lower, ['group', 'joining'])

                            d1 = pd.to_datetime(row_hc.iloc[0, idx_doj1], dayfirst=True, errors='coerce') if idx_doj1 is not None else pd.NaT
                            d2 = pd.to_datetime(row_hc.iloc[0, idx_doj2], dayfirst=True, errors='coerce') if idx_doj2 is not None else pd.NaT
                            d3 = pd.to_datetime(row_hc.iloc[0, idx_doj3], dayfirst=True, errors='coerce') if idx_doj3 is not None else pd.NaT

                            # --- UPDATED DOJ LOGIC: Fallback to Group Date (d3) if Legal Entity Date (d2) is blank ---
                            if pd.isna(d2):
                                d2_final = d3
                            else:
                                d2_final = d2

                            row_data['Employment Details Date of Joining'] = d1.strftime('%d-%m-%Y') if pd.notna(d1) else ""
                            row_data['Employment Details Legal Entity Date of Joining'] = d2_final.strftime('%d-%m-%Y') if pd.notna(d2_final) else ""
                            row_data['Employment Details Group Date of Joining'] = d3.strftime('%d-%m-%Y') if pd.notna(d3) else ""
                            
                            valid_dates = [d for d in [d1, d2_final, d3] if pd.notna(d)]
                            if valid_dates:
                                d_min_doj = min(valid_dates)
                                row_data['Min DOJ'] = d_min_doj.strftime('%d-%m-%Y')
                            else:
                                d_min_doj = pd.NaT
                                row_data['Min DOJ'] = ""
                            
                            idx_res_hc = lookup_column_index(hc_cols_lower, ['employment', 'details', 'date', 'resignation'])
                            if idx_res_hc is None: idx_res_hc = lookup_column_index(hc_cols_lower, ['date', 'resignation'])
                            
                            d_res = pd.to_datetime(row_hc.iloc[0, idx_res_hc], dayfirst=True, errors='coerce') if idx_res_hc is not None else pd.NaT
                            row_data['Resignation Date'] = d_res.strftime('%d-%m-%Y') if pd.notna(d_res) else ""
                        else:
                            for c in ['First Name', 'Middle Name', 'Last Name', 'Employee Type', 'Reason', 'Entity', 'Position', 'Department', 'Location', 'Employment Details Date of Joining', 'Employment Details Legal Entity Date of Joining', 'Employment Details Group Date of Joining', 'Min DOJ', 'Resignation Date']:
                                row_data[c] = ""

                        # --- PREVIOUSLY UPDATED LOGIC: FETCH LWD_SF FROM HC REPORT ---
                        if not row_hc.empty:
                            idx_lwd_hc = lookup_column_index(hc_cols_lower, ['employment', 'details', 'actual', 'exit', 'date'])
                            if idx_lwd_hc is None:
                                idx_lwd_hc = lookup_column_index(hc_cols_lower, ['actual', 'exit', 'date'])

                            if idx_lwd_hc is not None:
                                d_lwd_sf = pd.to_datetime(row_hc.iloc[0, idx_lwd_hc], dayfirst=True, errors='coerce')
                                row_data['LWD_SF'] = d_lwd_sf.strftime('%d-%m-%Y') if pd.notna(d_lwd_sf) else ""
                            else:
                                row_data['LWD_SF'] = ""
                        else:
                            row_data['LWD_SF'] = ""

                        # Tenure Logic modification: fallback to LWD_SF if d_res is empty
                        target_end_date = d_res if pd.notna(d_res) else d_lwd_sf

                        if pd.notna(target_end_date) and pd.notna(d_min_doj):
                            days_diff = (target_end_date - d_min_doj).days
                            row_data['Tenure'] = round((days_diff + 1) / 365.0, 2)
                        else:
                            row_data['Tenure'] = 0.0

                        # Core calculations from FFS report
                        row_data['NP recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['notice', 'days', 'recovered'])
                        if row_data['NP recovery'] == 0.0: row_data['NP recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['notice', 'period', 'recover'])

                        row_data['NP payable'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['payment', 'lieu', 'notice'])
                        if row_data['NP payable'] == 0.0: row_data['NP payable'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['lieu', 'notice'])

                        row_data['Onfield Allowance'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['field', 'allowance', 'payable'])
                        if row_data['Onfield Allowance'] == 0.0: row_data['Onfield Allowance'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['field', 'allowance'])

                        row_data['Mobile Allowance Recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['handset', 'allowance'])
                        if row_data['Mobile Allowance Recovery'] == 0.0: row_data['Mobile Allowance Recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['mobile', 'allowance'])

                        # Fetch from Relocation Expenses NJ Report
                        row_data['Relocation of Goods'] = parse_numeric_cell(row_nj, nj_cols_lower, ['goods'])
                        row_data['Relocation Travel Expense'] = parse_numeric_cell(row_nj, nj_cols_lower, ['flight'])
                        row_data['Relocation Hotel Expenses'] = parse_numeric_cell(row_nj, nj_cols_lower, ['hotel'])
                        row_data['Relocation conveyance'] = parse_numeric_cell(row_nj, nj_cols_lower, ['local', 'conveyance'])
                        if row_data['Relocation conveyance'] == 0.0: row_data['Relocation conveyance'] = parse_numeric_cell(row_nj, nj_cols_lower, ['conveyance'])

                        # Fetch from IJP Relocation of goods Report
                        row_data['IJP Relocation of Goods'] = parse_numeric_cell(row_ijp, ijp_cols_lower, ['goods'])
                        row_data['IJP Relocation Travel Expense'] = parse_numeric_cell(row_ijp, ijp_cols_lower, ['flight'])
                        row_data['IJP Relocation Hotel Expenses'] = parse_numeric_cell(row_ijp, ijp_cols_lower, ['hotel'])
                        row_data['IJP Relocation conveyance'] = parse_numeric_cell(row_ijp, ijp_cols_lower, ['local', 'conveyance'])
                        if row_data['IJP Relocation conveyance'] == 0.0: row_data['IJP Relocation conveyance'] = parse_numeric_cell(row_ijp, ijp_cols_lower, ['conveyance'])

                        row_data['Joining Bonus Rec. (SF)'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['joining', 'bonus', 'recovery'])
                        if row_data['Joining Bonus Rec. (SF)'] == 0.0: row_data['Joining Bonus Rec. (SF)'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['joining', 'bonus'])

                        row_data['Notice Period Buyout Recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['notice', 'period', 'buyout'])
                        if row_data['Notice Period Buyout Recovery'] == 0.0: row_data['Notice Period Buyout Recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['buyout'])

                        row_data['Retention bonus recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['retention', 'bonus', 'recovery'])
                        if row_data['Retention bonus recovery'] == 0.0: row_data['Retention bonus recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['retention', 'bonus'])

                        idx_itac = lookup_column_index(ffs_cols_lower, ['itac', 'final', 'recovery'])
                        row_data['IT Asset Recovery'] = pd.to_numeric(row_ffs.iloc[0, idx_itac], errors='coerce') if (idx_itac is not None and not row_ffs.empty) else 0.0

                        # Fetch Info Sec from FC : Finance Clearance column inside FFS report
                        idx_infosec = lookup_column_index(ffs_cols_lower, ['fc', 'finance', 'clearance'])
                        if idx_infosec is None: idx_infosec = lookup_column_index(ffs_cols_lower, ['finance', 'clearance'])
                        if idx_infosec is not None and not row_ffs.empty:
                            row_data['Info Sec'] = str(row_ffs.iloc[0, idx_infosec]).strip()
                        else:
                            row_data['Info Sec'] = ""

                        idx_facility = lookup_column_index(ffs_cols_lower, ['sc', 'financial', 'recovery', 'amount'])
                        row_data['Facility Recovery'] = pd.to_numeric(row_ffs.iloc[0, idx_facility], errors='coerce') if (idx_facility is not None and not row_ffs.empty) else 0.0

                        row_data['Finance Recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['fc', 'financial', 'recovery', 'amount'])
                        if row_data['Finance Recovery'] == 0.0: row_data['Finance Recovery'] = parse_numeric_cell(row_ffs, ffs_cols_lower, ['financial', 'recovery', 'amount'])
                        
                        idx_fin_reason = lookup_column_index(ffs_cols_lower, ['fc', 'financial', 'recovery', 'reason'])
                        if idx_fin_reason is None: idx_fin_reason = lookup_column_index(ffs_cols_lower, ['financial', 'recovery', 'reason'])
                        if not row_ffs.empty and idx_fin_reason is not None:
                            raw_reason = str(row_ffs.iloc[0, idx_fin_reason]).strip()
                            if pd.isna(row_ffs.iloc[0, idx_fin_reason]) or raw_reason == "":
                                row_data['Finance Recovery Reason'] = 0
                            elif raw_reason.lower() == "nil":
                                row_data['Finance Recovery Reason'] = "Nil"
                            else:
                                row_data['Finance Recovery Reason'] = raw_reason
                        else:
                            row_data['Finance Recovery Reason'] = 0

                        master_rows.append(row_data)

                    df_final_master = pd.DataFrame(master_rows)

                    # Exact structural schema ordered map requested by user
                    final_req_cols = [
                        'Employee ID', 'First Name', 'Middle Name', 'Last Name', 'Employee Type', 'Reason', 'Entity', 'Position', 'Department', 
                        'Employment Details Date of Joining', 'Employment Details Legal Entity Date of Joining', 'Employment Details Group Date of Joining', 
                        'Min DOJ', 'Resignation Date', 'LWD_SF', 'Tenure', 'Location', 'NP recovery', 'NP payable', 'Leave Encashment', 'Onfield Allowance', 
                        'Mobile Allowance Recovery', 'Relocation of Goods', 'Relocation Travel Expense', 'Relocation Hotel Expenses', 'Relocation conveyance', 
                        'IJP Relocation of Goods', 'IJP Relocation Travel Expense', 'IJP Relocation Hotel Expenses', 'IJP Relocation conveyance', 
                        'Joining Bonus Rec. (SF)', 'Notice Period Buyout Recovery', 'Retention bonus recovery', 'IT Asset Recovery', 'Info Sec', 
                        'Facility Recovery', 'Finance Recovery', 'Finance Recovery Reason', 'Food Allowance', 'Inventory Recovery', 'Car Lease', 
                        'NPS', 'Vital claim reimbursement', 'Remarks'
                    ]

                    # Dynamically fill entirely new/unmatched target profiles with raw default placeholders
                    for col in final_req_cols:
                        if col not in df_final_master.columns:
                            df_final_master[col] = ""

                    st.session_state.final_master_df = df_final_master[final_req_cols].set_index('Employee ID')
                    st.session_state.absconding_decision = 'pending'
                    st.rerun()

    # --- ENHANCED OFFBOARDING PROCESS INTERFACE MODULE ---
    if st.session_state.get('final_master_df') is not None:
        curr_df = st.session_state.final_master_df.reset_index()
        reason_col = 'Reason' if 'Reason' in curr_df.columns else 'Event Reason'
        
        abs_mask = curr_df[reason_col].astype(str).str.contains('33|absconding|termination|resignation', case=False, na=False)
        total_absconding_cases = abs_mask.sum()
        
        edit_mask = abs_mask & (pd.to_numeric(curr_df['NP recovery'], errors='coerce').fillna(0) == 0)
        review_cases_data = curr_df[edit_mask]
        
        if total_absconding_cases > 0 and st.session_state.absconding_decision == 'pending':
            st.markdown(f"<div class='alert-popup'><h2>⚠️ Total absconding cases: {total_absconding_cases}</h2></div>", unsafe_allow_html=True)
            c1, r_btn = st.columns(2)
            with c1: 
                if st.button("✂️ Exclude all from Master", use_container_width=True):
                    st.session_state.final_master_df = curr_df[~abs_mask].set_index('Employee ID')
                    st.session_state.absconding_decision = 'done'; st.rerun()
            with r_btn: 
                if st.button("📝 Review & Edit", use_container_width=True):
                    st.session_state.absconding_decision = 'review'; st.rerun()
        
        elif st.session_state.absconding_decision == 'review':
            st.markdown("### 📝 Edit Separation Cases")
            edited = st.data_editor(review_cases_data, use_container_width=True, key="abs_editor")
            if st.button("✔️ Save Edits & Proceed to Actual DOE", use_container_width=True):
                if 'Employee ID' in edited.columns:
                    edited_cols = edited.copy()
                else:
                    edited_cols = edited.reset_index()

                original_review = review_cases_data.set_index('Employee ID')
                comp_edited = edited_cols.set_index('Employee ID')
                
                cell_tracking = {}
                for emp_id in comp_edited.index:
                    cell_tracking[str(emp_id)] = []
                    for col in comp_edited.columns:
                        if col in original_review.columns:
                            if str(comp_edited.loc[emp_id, col]) != str(original_review.loc[emp_id, col]):
                                cell_tracking[str(emp_id)].append(col)
                
                st.session_state.edited_cells = cell_tracking
                curr_df.set_index('Employee ID', inplace=True)
                curr_df.update(comp_edited)
                st.session_state.final_master_df = curr_df
                st.session_state.absconding_decision = 'actual_doe_popup'
                st.rerun()

        elif st.session_state.absconding_decision == 'actual_doe_popup':
            st.markdown("### 📅 Enter Actual DOE Details")
            st.info("Please fill in the Actual Date of End (DOE) values for the processed employees below:")
            
            curr_df_with_idx = st.session_state.final_master_df.reset_index()
            target_mask = curr_df_with_idx[reason_col].astype(str).str.contains('33|absconding|termination|resignation', case=False, na=False) & (pd.to_numeric(curr_df_with_idx['NP recovery'], errors='coerce').fillna(0) == 0)
            target_employees = curr_df_with_idx[target_mask]
            
            form_dict = {}
            for _, row in target_employees.iterrows():
                emp_id = str(row['Employee ID'])
                form_dict[emp_id] = st.date_input(f"Employee ID: {emp_id} | Enter Actual DOE", value=None, key=f"doe_in_{emp_id}")

            if st.button("✔️ Save DOE & Finalize Master Report", use_container_width=True):
                np_rows = []
                for _, row in target_employees.iterrows():
                    emp_id = str(row['Employee ID'])
                    actual_doe_val = form_dict.get(emp_id)
                    lwd_sf_str = str(row['LWD_SF']).strip()
                    
                    actual_doe_dt = pd.to_datetime(actual_doe_val, errors='coerce')
                    res_doe_dt = pd.to_datetime(lwd_sf_str, dayfirst=True, errors='coerce')
                    
                    if pd.notna(actual_doe_dt) and pd.notna(res_doe_dt):
                        shortfall_days = (actual_doe_dt - res_doe_dt).days + 1
                    else:
                        shortfall_days = ""

                    np_rows.append({
                        'Employee ID': emp_id,
                        'DOE': lwd_sf_str if lwd_sf_str != "nan" else "",
                        'DOE as per resignation': actual_doe_dt.strftime('%d-%m-%Y') if pd.notna(actual_doe_dt) else "",
                        'Shortfall': shortfall_days
                    })
                
                st.session_state.np_absconding_df = pd.DataFrame(np_rows).set_index('Employee ID')
                st.session_state.absconding_decision = 'done'
                st.rerun()
                
        else:
            final = st.session_state.final_master_df.reset_index()
            st.success("✅ Consolidation Finished Successfully!")
            st.dataframe(final, use_container_width=True)
            
            # Show secondary shortfall sub-report context if available
            if st.session_state.get('np_absconding_df') is not None:
                st.markdown("### 📊 Calculated Shortfall References")
                st.dataframe(st.session_state.np_absconding_df.reset_index(), use_container_width=True)
                
            st.download_button("📥 Download Final Master Report", convert_to_styled_excel(final), "Master_FnF_Report.xlsx", use_container_width=True)
