import os
import tempfile
import pandas as pd
import streamlit as st

import fitz
from PIL import Image, ImageDraw
from streamlit_image_coordinates import streamlit_image_coordinates


def interactive_column_separator(uploaded_pdf, sample_page_num, session_key):
    """Renders a PDF page and allows the user to click to define column boundaries as percentage fractions."""
    if session_key not in st.session_state:
        st.session_state[session_key] = []

    # 1. Version tracking key to force component remounting when lines are cleared/deleted
    version_key = f"{session_key}_version"
    if version_key not in st.session_state:
        st.session_state[version_key] = 0

    # 2. State tracking key for click deduplication across Streamlit reruns
    last_click_key = f"{session_key}_last_click"
    if last_click_key not in st.session_state:
        st.session_state[last_click_key] = None

    st.markdown("**Visual Column Separator:** Click on the image below to add vertical column boundaries. These replace the need for 'Col Tolerance'.")

    # 3. Status Indicator Badges
    if st.session_state[session_key]:
        st.success(f"🟢 Active: Using {len(st.session_state[session_key])} Visual Column Separator(s)")
    else:
        st.info("ℹ️ No visual lines set — Falling back to standard Column Tolerance")

    col1, col2 = st.columns([3, 1])
    current_version = st.session_state[version_key]
    click_key = f"{session_key}_click_v{current_version}"

    with col1:
        # Load the specific page
        doc = fitz.open(stream=uploaded_pdf.getvalue(), filetype="pdf")
        page = doc[max(0, sample_page_num - 1)]

        # Render image at standard 72 DPI
        pix = page.get_pixmap(dpi=72)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Scale down for the UI view
        DISPLAY_WIDTH = 900 
        scale_ratio = DISPLAY_WIDTH / float(pix.width)
        display_height = int(pix.height * scale_ratio)
        img = img.resize((DISPLAY_WIDTH, display_height))

        # Draw existing separator lines using percentage fractions
        draw = ImageDraw.Draw(img)
        for pct in st.session_state[session_key]:
            # Normalize legacy point values if present
            pct_val = pct / page.rect.width if pct > 1.0 else pct
            display_x = pct_val * DISPLAY_WIDTH
            draw.line([(display_x, 0), (display_x, img.height)], fill="red", width=3)

        # Display clickable image using a versioned key
        coords = streamlit_image_coordinates(img, key=click_key)

        # 4. Deduplication Check: Only process if coords exist AND differ from previous click state
        if coords and coords != st.session_state[last_click_key]:
            st.session_state[last_click_key] = coords  # Store click to prevent rerun ghosting
            click_pct = coords['x'] / float(DISPLAY_WIDTH)

            # Use 1% (0.01) threshold to eliminate duplicate clicks nearby
            if not any(abs(click_pct - ex) < 0.01 for ex in st.session_state[session_key]):
                st.session_state[session_key].append(click_pct)
                st.session_state[session_key].sort()
                st.rerun()

    with col2:
        st.write("### Active Boundaries")
        
        # 5. Individual Line Removal Buttons
        lines_to_remove = []
        for idx, pct in enumerate(st.session_state[session_key]):
            pct_val = pct / page.rect.width if pct > 1.0 else pct
            lbl_col, btn_col = st.columns([3, 1])
            lbl_col.write(f"Line {idx + 1}: **{pct_val * 100:.1f}%**")
            if btn_col.button("❌", key=f"del_{session_key}_{idx}_v{current_version}"):
                lines_to_remove.append(pct)

        # Apply single line removal and increment version counter
        if lines_to_remove:
            for pct_to_rem in lines_to_remove:
                st.session_state[session_key].remove(pct_to_rem)
            st.session_state[version_key] += 1
            st.session_state[last_click_key] = None
            st.rerun()

        st.markdown("---")

        # Global Clear All and increment version counter
        if st.button("🗑️ Clear All Lines", key=f"clear_{session_key}_v{current_version}"):
            st.session_state[session_key] = []
            st.session_state[version_key] += 1
            st.session_state[last_click_key] = None
            st.rerun()

    return st.session_state[session_key]


# Custom backend helper import
from qc_helpers import (
    process_and_compare_pdfs,
    process_and_compare_ditch_layout,
    process_and_compare_top_elevations,
    parse_page_ranges,
    parse_annotation_loc
)

try:
    from dashboard_generator import generate_high_res_dashboard
except ImportError:
    generate_high_res_dashboard = None

st.set_page_config(page_title="Engineering QC/QA & Dashboard Tool",page_icon="🧐", layout="wide")
st.title("🏗️ Engineering Plan & Schedule QC/QA Audit Tool")



# --- ADD THIS HELPER FUNCTION ---
def get_safe_temp_path(suffix):
    """Creates a secure temporary file, closes it to prevent permission locks, and returns the path."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.close()
    return tmp.name
# --------------------------------


# Initialize Session State
if "da_run_successful" not in st.session_state: st.session_state.da_run_successful = False
if "da_excel_data" not in st.session_state: st.session_state.da_excel_data = None
if "da_pdf_data" not in st.session_state: st.session_state.da_pdf_data = None

if "plan_success" not in st.session_state: st.session_state.plan_success = False
if "plan_excel" not in st.session_state: st.session_state.plan_excel = None
if "plan_pdf" not in st.session_state: st.session_state.plan_pdf = None

if "prof_success" not in st.session_state: st.session_state.prof_success = False
if "prof_excel" not in st.session_state: st.session_state.prof_excel = None
if "prof_pdf" not in st.session_state: st.session_state.prof_pdf = None

if "top_success" not in st.session_state: st.session_state.top_success = False
if "top_excel" not in st.session_state: st.session_state.top_excel = None
if "top_pdf" not in st.session_state: st.session_state.top_pdf = None

if "ditch_success" not in st.session_state: st.session_state.ditch_success = False
if "ditch_excel" not in st.session_state: st.session_state.ditch_excel = None
if "ditch_pdf" not in st.session_state: st.session_state.ditch_pdf = None

if "da_config_df" not in st.session_state:
    st.session_state.da_config_df = pd.DataFrame({
        "Variable": ["Name", "DrainageArea"],
        "Table_Column": ["Col_1", "Col_2"],
        "Annotation_Location": ["Target", "1 Above"]
    })

if "config_df" not in st.session_state:
    st.session_state.config_df = pd.DataFrame({
        "Variable": ["Name", "Station", "Offset", "Baseline"],
        "Table_Column": ["Col_1", "Col_2", "Col_3", "Col_4"],
        "Annotation_Location": ["Target", "2 Above", "1 Above", "3 Above"]
    })

# --- ADDED: Session state for Plan Baseline Aliases ---
if "plan_aliases_df" not in st.session_state:
    st.session_state.plan_aliases_df = pd.DataFrame({
        "Plan Format": ["PFENNIGLN", "HWY 35"],
        "Table Format": ["PFENNIG LN", "I-35"]
    })

if "prof_config_df" not in st.session_state:
    st.session_state.prof_config_df = pd.DataFrame({
        "Variable": ["Name", "Slope", "Size_Type", "Length"],
        "Table_Column": ["Col_1", "Col_11", "Col_10", "Col_12"],
        "Annotation_Location": ["Target", "1 Above", "2 Above", "3 Above"]
    })
    
# --- UPDATED: Top Elevation now includes Annotation Location by default ---
if "top_config_df" not in st.session_state:
    st.session_state.top_config_df = pd.DataFrame({
        "Variable": ["Name", "Top_Elevation"],
        "Table_Column": ["Col_1", "Col_5"],
        "Annotation_Location": ["Target", "1 Below"]
    })

if "ditch_col_map_df" not in st.session_state:
    st.session_state.ditch_col_map_df = pd.DataFrame({
        "Variable": ["Name", "Baseline", "Station", "Offset"],
        "Table_Column": ["Col_1", "Col_4", "Col_2", "Col_3"],
        "Annotation_Location": ["Target", "1 Below", "2 Below", "3 Below"]
    })

if "ditch_aliases_df" not in st.session_state:
    st.session_state.ditch_aliases_df = pd.DataFrame({
        "Plan Format": ["PFENNIGLN", "HWY 35"],
        "Table Format": ["PFENNIG LN", "I-35"]
    })

tab_da, tab_plan, tab_profile, tab_top, tab_ditch, tab_vis = st.tabs([
    "Drainage Area", "Plan", "Profile", "Top Elevation", "Ditch", "📊 Visualization Dashboard"
])

# =============================================================================
# TAB 1: DRAINAGE AREA
# =============================================================================
with tab_da:
    st.header("Dynamic Variables & Column Mapping")
    st.markdown("Define the variables to extract for Drainage Area. **One variable must have 'Target' as its Annotation_Location**.")
    
    da_edited_config = st.data_editor(
        st.session_state.da_config_df,
        num_rows="dynamic", use_container_width=True, hide_index=True, key="da_data_editor",
        column_config={
            "Annotation_Location": st.column_config.SelectboxColumn(
                "Annotation Location",
                options=["4 Above", "3 Above", "2 Above", "1 Above", "Target", "1 Below", "2 Below", "3 Below", "4 Below", "None"],
                required=True
            )
        }
    )
    
    st.divider()
    st.header("Plan Comparison Setup")
    uploaded_da_pdf = st.file_uploader("Upload Drainage Plan PDF", type=['pdf'], key="da_pdf_uploader")
    
    # --- ADD THE VISUAL SEPARATOR HERE ---
    da_visual_splits = []
    if uploaded_da_pdf:
        # Let user pick which table page to preview
        preview_page = st.number_input("Sample Table Page for Column Setup", value=22, min_value=1, key="da_preview_page")
        da_visual_splits = interactive_column_separator(uploaded_da_pdf, preview_page, "da_x_splits")
    # ------------------------------------    

    da_structure_pattern = st.text_input("Structure Regex Pattern", value=r"\b[A-Za-z]+-[A-Za-z]+-\d+\b", key="da_structure_pattern")
    
    da_col1, da_col2, da_col3, da_col4 = st.columns(4)
    with da_col1: da_full_pages = st.text_input("Plan Full Pages (e.g., 2-5)", value="2-5", key="da_full_pages")
    with da_col2: da_top_pages = st.text_input("Plan Top Half Pages", value="", key="da_top_pages")
    with da_col3: da_bottom_pages = st.text_input("Plan Bottom Half Pages", value="", key="da_bottom_pages")
    with da_col4: da_table_pages_input = st.text_input("Table Pages", value="9", key="da_table_pages")

    with st.expander("Advanced Tolerances & Page Margins", expanded=False):
        da_t_col1, da_t_col2 = st.columns(2)
        with da_t_col1:
            da_row_tol = st.number_input("Row Tolerance", value=5, key="da_row_tol")
            da_col_tol = st.number_input("Col Tolerance", value=20, key="da_col_tol")
            da_skip_header_rows = st.number_input("Skip Header Rows (Table)", value=3, step=1, key="da_skip_header")
        with da_t_col2:
            da_search_w = st.number_input("Search Width Mult", value=1.0, key="da_search_w")
            da_search_h = st.number_input("Search Height Mult", value=2.0, key="da_search_h")
            
        st.markdown("---")
        da_m_col1, da_m_col2 = st.columns(2)
        with da_m_col1:
            st.markdown("**Table Margins (Inches)**")
            da_margin_left = st.number_input("Table Left Margin", value=1.5, step=0.1, key="da_t_left")
            da_margin_right = st.number_input("Table Right Margin", value=2.5, step=0.1, key="da_t_right")
            da_margin_top = st.number_input("Table Top Margin", value=0.5, step=0.1, key="da_t_top")
            da_margin_bottom = st.number_input("Table Bottom Margin", value=0.5, step=0.1, key="da_t_bottom")
        with da_m_col2:
            st.markdown("**Plan Margins (Inches)**")
            da_plan_margin_left = st.number_input("Plan Left Margin", value=0.0, step=0.1, key="da_p_left")
            da_plan_margin_right = st.number_input("Plan Right Margin", value=0.0, step=0.1, key="da_p_right")
            da_plan_margin_top = st.number_input("Plan Top Margin", value=0.0, step=0.1, key="da_p_top")
            da_plan_margin_bottom = st.number_input("Plan Bottom Margin", value=0.0, step=0.1, key="da_p_bottom")
        
    st.divider()

    if st.button("Run Drainage Area Comparison", type="primary", key="btn_run_da"):
        da_config_list = da_edited_config.to_dict('records')
        da_target_var = next((row['Variable'] for row in da_config_list if str(row['Annotation_Location']).strip().lower() == 'target'), None)
        
        if not uploaded_da_pdf:
            st.error("Please upload a Drainage PDF file.")
        elif not da_target_var:
            st.error("Error: You must assign exactly one variable's Annotation_Location to 'Target'.")
        else:
            for c in da_config_list: c['parsed_loc'] = parse_annotation_loc(c['Annotation_Location'])
            da_overrides = {}
            for p in parse_page_ranges(da_full_pages): da_overrides[p] = "full"
            for p in parse_page_ranges(da_top_pages): da_overrides[p] = "top"
            for p in parse_page_ranges(da_bottom_pages): da_overrides[p] = "bottom"
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_input_pdf:
                tmp_input_pdf.write(uploaded_da_pdf.read())
                tmp_input_path = tmp_input_pdf.name
                
            # out_excel_path, out_pdf_path = tempfile.mktemp(suffix=".xlsx"), tempfile.mktemp(suffix=".pdf")
            out_excel_path = get_safe_temp_path(".xlsx")
            out_pdf_path = get_safe_temp_path(".pdf")

            with st.spinner("Processing Drainage PDF..."):
                try:
                    process_and_compare_pdfs(
                        pdf_path=tmp_input_path, table_pages=parse_page_ranges(da_table_pages_input),
                        plan_page_config={"default": "full", "overrides": da_overrides},
                        dynamic_config=da_config_list, target_var=da_target_var,
                        plan_margin_left=da_plan_margin_left, plan_margin_right=da_plan_margin_right,
                        plan_margin_top=da_plan_margin_top, plan_margin_bottom=da_plan_margin_bottom,
                        margin_left_in=da_margin_left, margin_right_in=da_margin_right,
                        margin_top_in=da_margin_top, margin_bottom_in=da_margin_bottom,
                        row_tolerance=da_row_tol, col_tolerance=da_col_tol, search_width_mult=da_search_w, search_height_mult=da_search_h,
                        skip_header_rows=int(da_skip_header_rows), structure_pattern=da_structure_pattern,
                        x_separators=da_visual_splits,
                        output_excel=out_excel_path, marked_pdf_output=out_pdf_path
                    )
                    st.session_state.da_excel_data = open(out_excel_path, "rb").read()
                    st.session_state.da_pdf_data = open(out_pdf_path, "rb").read()
                    st.session_state.da_run_successful = True
                    st.success("Drainage area comparison complete!")
                except Exception as e:
                    st.error(f"An error occurred: {e}")
                    st.session_state.da_run_successful = False
                finally:
                    for p in [tmp_input_path, out_excel_path, out_pdf_path]:
                        if os.path.exists(p): os.remove(p)

    if st.session_state.da_run_successful:
        st.write("### Download Drainage Results")
        d_col1, d_col2 = st.columns(2)
        d_col1.download_button("📥 Download Drainage Excel", st.session_state.da_excel_data, "drainage_comparison_summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_da_xl")
        d_col2.download_button("📥 Download Marked Drainage PDF", st.session_state.da_pdf_data, "marked_drainage_output.pdf", "application/pdf", key="dl_da_pdf")

# =============================================================================
# TAB 2: PLAN
# =============================================================================
with tab_plan:
    st.header("Plan Comparison Setup")
    
    st.subheader("1. Plan Column Mapping")
    edited_config = st.data_editor(
        st.session_state.config_df,
        num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "Annotation_Location": st.column_config.SelectboxColumn(
                "Annotation Location",
                options=["4 Above", "3 Above", "2 Above", "1 Above", "Target", "1 Below", "2 Below", "3 Below", "4 Below", "None"],
                required=True
            )
        }
    )
    
    st.subheader("2. Baseline Aliases")
    st.caption("Map alternate naming conventions (e.g., 'PFENNIGLN' in plan vs 'PFENNIG LN' in table).")
    plan_edited_aliases = st.data_editor(
        st.session_state.plan_aliases_df,
        num_rows="dynamic", use_container_width=True, hide_index=True, key="plan_alias_editor"
    )
    
    st.divider()
    uploaded_pdf = st.file_uploader("Upload Plan PDF", type=['pdf'], key="plan_pdf_uploader")
    
    # --- ADD THE VISUAL SEPARATOR HERE ---
    plan_visual_splits = []
    if uploaded_pdf:
        # Let user pick which table page to preview
        preview_page = st.number_input("Sample Table Page for Column Setup", value=24, min_value=1, key="plan_preview_page")
        plan_visual_splits = interactive_column_separator(uploaded_pdf, preview_page, "plan_x_splits")
    # ------------------------------------    

    structure_pattern = st.text_input("Structure Regex Pattern", value=r"\b[A-Za-z]{2}-[A-Za-z](?![PCp c])[A-Za-z]-\d{1,4}\b", key="plan_regex")

    p_col1, p_col2, p_col3, p_col4 = st.columns(4)
    with p_col1: full_pages = st.text_input("Plan Full Pages", key="plan_full")
    with p_col2: top_pages = st.text_input("Plan Top Half Pages", value="18-26", key="plan_top")
    with p_col3: bottom_pages = st.text_input("Plan Bottom Half Pages", key="plan_bottom")
    with p_col4: table_pages_input = st.text_input("Table Pages", value="14", key="plan_table_p")

    with st.expander("Advanced Tolerances & Page Margins", expanded=False):
        t_col1, t_col2 = st.columns(2)
        with t_col1:
            row_tol = st.number_input("Row Tolerance", value=5, key="plan_row_tol")
            col_tol = st.number_input("Col Tolerance", value=27, key="plan_col_tol")
            skip_header_rows = st.number_input("Skip Header Rows", value=3, step=1, key="plan_skip")
        with t_col2:
            search_w = st.number_input("Search Width Mult", value=2.25, key="plan_sw")
            search_h = st.number_input("Search Height Mult", value=1.25, key="plan_sh")
            
        m_col1, m_col2 = st.columns(2)
        with m_col1:
            margin_left = st.number_input("Table Left Margin", value=1.5, step=0.1, key="t_left")
            margin_right = st.number_input("Table Right Margin", value=2.5, step=0.1, key="t_right")
            margin_top = st.number_input("Table Top Margin", value=0.5, step=0.1, key="t_top")
            margin_bottom = st.number_input("Table Bottom Margin", value=0.5, step=0.1, key="t_bottom")
        with m_col2:
            plan_margin_left = st.number_input("Plan Left Margin", value=0.0, step=0.1, key="p_left")
            plan_margin_right = st.number_input("Plan Right Margin", value=0.0, step=0.1, key="p_right")
            plan_margin_top = st.number_input("Plan Top Margin", value=0.0, step=0.1, key="p_top")
            plan_margin_bottom = st.number_input("Plan Bottom Margin", value=0.0, step=0.1, key="p_bot")

    if st.button("Run Plan Comparison", type="primary", key="btn_run_plan"):
        config_list = edited_config.to_dict('records')
        target_var = next((row['Variable'] for row in config_list if str(row['Annotation_Location']).strip().lower() == 'target'), None)
        
        plan_alias_records = plan_edited_aliases.to_dict('records')
        plan_alias_dict = {str(r['Plan Format']): str(r['Table Format']) for r in plan_alias_records if r.get('Plan Format')}

        if not uploaded_pdf:
            st.error("Please upload a Plan PDF file.")
        elif not target_var:
            st.error("Error: You must assign exactly one variable's Annotation_Location to 'Target'.")
        else:
            for c in config_list: c['parsed_loc'] = parse_annotation_loc(c['Annotation_Location'])
            overrides = {}
            for p in parse_page_ranges(full_pages): overrides[p] = "full"
            for p in parse_page_ranges(top_pages): overrides[p] = "top"
            for p in parse_page_ranges(bottom_pages): overrides[p] = "bottom"
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
                tmp_in.write(uploaded_pdf.read())
                tmp_in_path = tmp_in.name
                
            # out_xl, out_pdf = tempfile.mktemp(suffix=".xlsx"), tempfile.mktemp(suffix=".pdf")
            out_xl = get_safe_temp_path(".xlsx")
            out_pdf = get_safe_temp_path(".pdf")

            with st.spinner("Processing Plan PDF..."):
                try:
                    process_and_compare_pdfs(
                        pdf_path=tmp_in_path, table_pages=parse_page_ranges(table_pages_input),
                        plan_page_config={"default": "full", "overrides": overrides},
                        dynamic_config=config_list, target_var=target_var,
                        baseline_aliases=plan_alias_dict,
                        plan_margin_left=plan_margin_left, plan_margin_right=plan_margin_right,
                        plan_margin_top=plan_margin_top, plan_margin_bottom=plan_margin_bottom,
                        margin_left_in=margin_left, margin_right_in=margin_right,
                        margin_top_in=margin_top, margin_bottom_in=margin_bottom,
                        row_tolerance=row_tol, col_tolerance=col_tol, search_width_mult=search_w, search_height_mult=search_h,
                        skip_header_rows=int(skip_header_rows), structure_pattern=structure_pattern,
                        x_separators=plan_visual_splits,
                        output_excel=out_xl, marked_pdf_output=out_pdf
                    )
                    st.session_state.plan_excel = open(out_xl, "rb").read()
                    st.session_state.plan_pdf = open(out_pdf, "rb").read()
                    st.session_state.plan_success = True
                    st.success("Plan comparison completed successfully!")
                except Exception as e:
                    st.error(f"Error during plan comparison: {e}")
                    st.session_state.plan_success = False
                finally:
                    for p in [tmp_in_path, out_xl, out_pdf]:
                        if os.path.exists(p): os.remove(p)

    if st.session_state.plan_success:
        st.write("### Download Plan Results")
        d1, d2 = st.columns(2)
        d1.download_button("📥 Download Plan Comparison Excel", st.session_state.plan_excel, "comparison_summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_plan_xl")
        d2.download_button("📥 Download Marked Plan PDF", st.session_state.plan_pdf, "marked_output.pdf", "application/pdf", key="dl_plan_pdf")

# =============================================================================
# TAB 3: PROFILE
# =============================================================================
with tab_profile:
    st.header("Profile Comparison Setup")
    edited_prof_config = st.data_editor(
        st.session_state.prof_config_df,
        num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "Annotation_Location": st.column_config.SelectboxColumn(
                "Annotation Location",
                options=["4 Above", "3 Above", "2 Above", "1 Above", "Target", "1 Below", "2 Below", "3 Below", "4 Below", "None"],
                required=True
            )
        }
    )

    st.divider()
    uploaded_prof_pdf = st.file_uploader("Upload Profile PDF", type=['pdf'], key="prof_pdf_uploader")
    
    # --- ADD THE VISUAL SEPARATOR HERE ---
    profile_visual_splits = []
    if uploaded_prof_pdf:
        # Let user pick which table page to preview
        preview_page = st.number_input("Sample Table Page for Column Setup", value=24, min_value=1, key="profile_preview_page")
        profile_visual_splits = interactive_column_separator(uploaded_prof_pdf, preview_page, "profile_visual_splits")
    # ------------------------------------    


    prof_structure_pattern = st.text_input("Profile Structure Regex Pattern", value=r"\b[A-Za-z]{2}-[A-Za-z]P-[A-Za-z0-9-]+\b", key="prof_regex")

    pr_col1, pr_col2, pr_col3, pr_col4 = st.columns(4)
    with pr_col1: prof_full_pages = st.text_input("Profile Full Pages", key="prof_full")
    with pr_col2: prof_top_pages = st.text_input("Profile Top Half Pages", key="prof_top")
    with pr_col3: prof_bottom_pages = st.text_input("Profile Bottom Half Pages", value="18-26", key="prof_bottom")
    with pr_col4: prof_table_pages_input = st.text_input("Table Pages", value="11", key="prof_table_p")

    with st.expander("Advanced Tolerances & Page Margins (Profile)", expanded=False):
        pt_col1, pt_col2 = st.columns(2)
        with pt_col1:
            prof_row_tol = st.number_input("Row Tolerance", value=5, key="prof_row_tol")
            prof_col_tol = st.number_input("Col Tolerance", value=20, key="prof_col_tol")
            prof_skip_header = st.number_input("Skip Header Rows", value=3, step=1, key="prof_skip")
        with pt_col2:
            prof_search_w = st.number_input("Search Width Mult", value=2.0, key="prof_sw")
            prof_search_h = st.number_input("Search Height Mult", value=1.25, key="prof_sh")

        pm_col1, pm_col2 = st.columns(2)
        with pm_col1:
            prof_margin_left = st.number_input("Table Left Margin", value=1.5, step=0.1, key="prof_t_left")
            prof_margin_right = st.number_input("Table Right Margin", value=2.5, step=0.1, key="prof_t_right")
            prof_margin_top = st.number_input("Table Top Margin", value=0.5, step=0.1, key="prof_t_top")
            prof_margin_bottom = st.number_input("Table Bottom Margin", value=0.5, step=0.1, key="prof_t_bot")
        with pm_col2:
            prof_p_margin_left = st.number_input("Profile Left Margin", value=0.0, step=0.1, key="prof_p_left")
            prof_p_margin_right = st.number_input("Profile Right Margin", value=0.0, step=0.1, key="prof_p_right")
            prof_p_margin_top = st.number_input("Profile Top Margin", value=0.0, step=0.1, key="prof_p_top")
            prof_p_margin_bottom = st.number_input("Profile Bottom Margin", value=0.0, step=0.1, key="prof_p_bot")

    if st.button("Run Profile Comparison", type="primary", key="btn_run_prof"):
        prof_config_list = edited_prof_config.to_dict('records')
        prof_target_var = next((row['Variable'] for row in prof_config_list if str(row['Annotation_Location']).strip().lower() == 'target'), None)

        if not uploaded_prof_pdf:
            st.error("Please upload a Profile PDF file.")
        elif not prof_target_var:
            st.error("Error: You must assign exactly one variable's Annotation_Location to 'Target'.")
        else:
            for c in prof_config_list: c['parsed_loc'] = parse_annotation_loc(c['Annotation_Location'])
            prof_overrides = {}
            for p in parse_page_ranges(prof_full_pages): prof_overrides[p] = "full"
            for p in parse_page_ranges(prof_top_pages): prof_overrides[p] = "top"
            for p in parse_page_ranges(prof_bottom_pages): prof_overrides[p] = "bottom"

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
                tmp_in.write(uploaded_prof_pdf.read())
                tmp_in_path = tmp_in.name

            # out_xl = tempfile.mktemp(suffix=".xlsx")
            # out_pdf = tempfile.mktemp(suffix=".pdf")
            out_xl = get_safe_temp_path(".xlsx")
            out_pdf = get_safe_temp_path(".pdf")

            with st.spinner("Processing Profile PDF..."):
                try:
                    process_and_compare_pdfs(
                        pdf_path=tmp_in_path, table_pages=parse_page_ranges(prof_table_pages_input),
                        plan_page_config={"default": "bottom", "overrides": prof_overrides},
                        dynamic_config=prof_config_list, target_var=prof_target_var,
                        plan_margin_left=prof_p_margin_left, plan_margin_right=prof_p_margin_right,
                        plan_margin_top=prof_p_margin_top, plan_margin_bottom=prof_p_margin_bottom,
                        margin_left_in=prof_margin_left, margin_right_in=prof_margin_right,
                        margin_top_in=prof_margin_top, margin_bottom_in=prof_margin_bottom,
                        row_tolerance=prof_row_tol, col_tolerance=prof_col_tol,
                        search_width_mult=prof_search_w, search_height_mult=prof_search_h,
                        skip_header_rows=int(prof_skip_header), structure_pattern=prof_structure_pattern,
                        x_separators=profile_visual_splits,
                        output_excel=out_xl, marked_pdf_output=out_pdf
                    )
                    st.session_state.prof_excel = open(out_xl, "rb").read()
                    st.session_state.prof_pdf = open(out_pdf, "rb").read()
                    st.session_state.prof_success = True
                    st.success("Profile comparison completed successfully!")
                except Exception as e:
                    st.error(f"Error during profile comparison: {e}")
                    st.session_state.prof_success = False
                finally:
                    for p in [tmp_in_path, out_xl, out_pdf]:
                        if os.path.exists(p): os.remove(p)

    if st.session_state.prof_success:
        st.write("### Download Profile Results")
        pd1, pd2 = st.columns(2)
        pd1.download_button("📥 Download Profile Comparison Excel", st.session_state.prof_excel, "profile_comparison_summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_prof_xl")
        pd2.download_button("📥 Download Marked Profile PDF", st.session_state.prof_pdf, "marked_profile_output.pdf", "application/pdf", key="dl_prof_pdf")

# =============================================================================
# TAB 4: TOP ELEVATION
# =============================================================================
with tab_top:
    st.header("Top Elevation Comparison Setup")
    
    # --- UPDATED: Included specific formatting for dropdown selection ---
    top_edited_config = st.data_editor(
        st.session_state.top_config_df,
        num_rows="dynamic", use_container_width=True, hide_index=True, key="top_data_editor",
        column_config={
            "Annotation_Location": st.column_config.SelectboxColumn(
                "Annotation Location (if text is vertical above means to the right and below means to the left)",
                options=["4 Above", "3 Above", "2 Above", "1 Above", "Target", "1 Below", "2 Below", "3 Below", "4 Below", "None"],
                required=True
            )
        }
    )

    st.divider()
    top_orientation = st.radio(
    "Plan Annotation Orientation", 
    options=["Horizontal", "Vertical (Bottom-to-Top)"], 
    horizontal=True,
    key="top_orient_toggle"
    )
    uploaded_top_pdf = st.file_uploader("Upload Plan/Profile PDF for Elevation", type=['pdf'], key="top_pdf_uploader")
    
    # --- ADD THE VISUAL SEPARATOR HERE ---
    top_ele_visual_splits = []
    if uploaded_top_pdf:
        # Let user pick which table page to preview
        preview_page = st.number_input("Sample Table Page for Column Setup", value=24, min_value=1, key="top_ele_preview_page")
        top_ele_visual_splits = interactive_column_separator(uploaded_top_pdf, preview_page, "top_ele_visual_splits")
    # ------------------------------------    

    top_structure_pattern = st.text_input("Structure Regex Pattern", value=r"\b[A-Za-z]{2}-[A-Za-z](?![PCp c])[A-Za-z]-\d{1,4}\b", key="top_regex")

    tp_col1, tp_col2, tp_col3, tp_col4 = st.columns(4)
    with tp_col1: top_full_pages = st.text_input("Full Pages", value="", key="top_full")
    with tp_col2: top_top_pages = st.text_input("Top Half Pages", value="18-26", key="top_top")
    with tp_col3: top_bottom_pages = st.text_input("Bottom Half Pages", value="", key="top_bottom")
    with tp_col4: top_table_pages_input = st.text_input("Table Pages", value="14", key="top_table_p")

    with st.expander("Advanced Margins & Tolerances (Top Elevation)", expanded=False):
        tt_col1, tt_col2 = st.columns(2)
        with tt_col1:
            top_row_tol = st.number_input("Row Tolerance", value=5, key="top_row_tol")
            top_col_tol = st.number_input("Col Tolerance", value=20, key="top_col_tol")
            top_skip_header = st.number_input("Skip Header Rows", value=3, step=1, key="top_skip")
        with tt_col2:
            top_search_w = st.number_input("Search Width Mult", value=2.0, key="top_sw")
            top_search_h = st.number_input("Search Height Mult", value=1.25, key="top_sh")

        tm_col1, tm_col2 = st.columns(2)
        with tm_col1:
            top_margin_left = st.number_input("Table Left Margin", value=1.5, step=0.1, key="top_t_left")
            top_margin_right = st.number_input("Table Right Margin", value=2.5, step=0.1, key="top_t_right")
            top_margin_top = st.number_input("Table Top Margin", value=0.5, step=0.1, key="top_t_top")
            top_margin_bottom = st.number_input("Table Bottom Margin", value=0.5, step=0.1, key="top_t_bot")
        with tm_col2:
            top_p_margin_left = st.number_input("Plan Left Margin", value=0.0, step=0.1, key="top_p_left")
            top_p_margin_right = st.number_input("Plan Right Margin", value=0.0, step=0.1, key="top_p_right")
            top_p_margin_top = st.number_input("Plan Top Margin", value=0.0, step=0.1, key="top_p_top")
            top_p_margin_bottom = st.number_input("Plan Bottom Margin", value=0.0, step=0.1, key="top_p_bot")

    if st.button("Run Top Elevation Comparison", type="primary", key="btn_run_top"):
        # --- UPDATED: Capturing dynamically chosen config ---
        top_cfg_dict = top_edited_config.to_dict('records')
        top_target_var = next((row['Variable'] for row in top_cfg_dict if str(row.get('Annotation_Location', '')).strip().lower() == 'target'), None)
        
        if not uploaded_top_pdf:
            st.error("Please upload a PDF file.")
        elif not top_target_var:
            st.error("Error: You must assign exactly one variable's Annotation_Location to 'Target'.")
        else:
            for c in top_cfg_dict: 
                c['parsed_loc'] = parse_annotation_loc(c.get('Annotation_Location', 'None'))
            
            top_overrides = {}
            for p in parse_page_ranges(top_full_pages): top_overrides[p] = "full"
            for p in parse_page_ranges(top_top_pages): top_overrides[p] = "top"
            for p in parse_page_ranges(top_bottom_pages): top_overrides[p] = "bottom"

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
                tmp_in.write(uploaded_top_pdf.read())
                tmp_in_path = tmp_in.name

            # out_xl = tempfile.mktemp(suffix=".xlsx")
            # out_pdf = tempfile.mktemp(suffix=".pdf")
            out_xl = get_safe_temp_path(".xlsx")
            out_pdf = get_safe_temp_path(".pdf")

            with st.spinner("Comparing Top Elevations..."):
                try:
                    # process_and_compare_top_elevations(
                    #     pdf_path=tmp_in_path,
                    #     table_pages=parse_page_ranges(top_table_pages_input),
                    #     plan_page_config={"default": "top", "overrides": top_overrides},
                    #     dynamic_config=top_cfg_dict, # --- UPDATED: Used dynamically instead of static config map ---
                    #     target_var=top_target_var,
                    #     margin_left_in=top_margin_left, margin_right_in=top_margin_right,
                    #     margin_top_in=top_margin_top, margin_bottom_in=top_margin_bottom,
                    #     row_tolerance=top_row_tol, col_tolerance=top_col_tol,
                    #     search_width_mult=top_search_w, search_height_mult=top_search_h,
                    #     skip_header_rows=int(top_skip_header),
                    #     structure_pattern=top_structure_pattern,
                    #     output_excel=out_xl, marked_pdf_output=out_pdf
                    # )
                    
                    process_and_compare_top_elevations(
                        pdf_path=tmp_in_path,
                        table_pages=parse_page_ranges(top_table_pages_input),
                        plan_page_config={"default": "top", "overrides": top_overrides},
                        dynamic_config=top_cfg_dict,
                        target_var=top_target_var,
                        margin_left_in=top_margin_left, margin_right_in=top_margin_right,
                        margin_top_in=top_margin_top, margin_bottom_in=top_margin_bottom,
                        row_tolerance=top_row_tol, col_tolerance=top_col_tol,
                        search_width_mult=top_search_w, search_height_mult=top_search_h,
                        skip_header_rows=int(top_skip_header),
                        structure_pattern=top_structure_pattern,
                        orientation="vertical_b2t" if "Vertical" in top_orientation else "horizontal",
                        x_separators=top_ele_visual_splits,
                        output_excel=out_xl, marked_pdf_output=out_pdf
                    )
                    st.session_state.top_excel = open(out_xl, "rb").read()
                    st.session_state.top_pdf = open(out_pdf, "rb").read()
                    st.session_state.top_success = True
                    st.success("Top Elevation comparison completed!")
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.session_state.top_success = False
                finally:
                    for p in [tmp_in_path, out_xl, out_pdf]:
                        if os.path.exists(p): os.remove(p)

    if st.session_state.top_success:
        st.write("### Download Top Elevation Results")
        td1, td2 = st.columns(2)
        td1.download_button("📥 Download Elevation Excel", st.session_state.top_excel, "profile_elevation_comparison.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_top_xl")
        td2.download_button("📥 Download Marked PDF", st.session_state.top_pdf, "marked_top_elevation_output.pdf", "application/pdf", key="dl_top_pdf")

# =============================================================================
# TAB 5: DITCH
# =============================================================================
with tab_ditch:
    st.header("Ditch Layout & Baseline Comparison Setup")
    
    st.subheader("1. Ditch Column Mapping")
    ditch_edited_col_map = st.data_editor(
        st.session_state.ditch_col_map_df,
        num_rows="dynamic", use_container_width=True, hide_index=True, key="ditch_col_editor",
        column_config={
            "Annotation_Location": st.column_config.SelectboxColumn(
                "Annotation Location (if text is vertical above means to the right and below means to the left)",
                options=["4 Above", "3 Above", "2 Above", "1 Above", "Target", "1 Below", "2 Below", "3 Below", "4 Below", "None"],
                required=True
            )
        }
    )

    st.subheader("2. Baseline Aliases")
    st.caption("Map alternate naming conventions (e.g., 'PFENNIGLN' in plan vs 'PFENNIG LN' in table).")
    ditch_edited_aliases = st.data_editor(
        st.session_state.ditch_aliases_df,
        num_rows="dynamic", use_container_width=True, hide_index=True, key="ditch_alias_editor"
    )

    st.divider()
    ditch_orientation = st.radio(
        "Plan Annotation Orientation", 
        options=["Horizontal", "Vertical (Bottom-to-Top)"], 
        horizontal=True,
        key="ditch_orient_toggle"
    )
    uploaded_ditch_pdf = st.file_uploader("Upload Ditch Plan PDF", type=['pdf'], key="ditch_pdf_uploader")
    
    # --- ADD THE VISUAL SEPARATOR HERE ---
    ditch_visual_splits = []
    if uploaded_ditch_pdf:
        # Let user pick which table page to preview
        preview_page = st.number_input("Sample Table Page for Column Setup", value=24, min_value=1, key="ditch_preview_page")
        ditch_visual_splits = interactive_column_separator(uploaded_ditch_pdf, preview_page, "ditch_visual_splits")
    # ------------------------------------    

    ditch_structure_pattern = st.text_input("Structure Regex Pattern", value=r"\b[A-Za-z]{2}-[A-Za-z]{2}-\d{1,4}\b", key="ditch_regex")

    dt_col1, dt_col2, dt_col3, dt_col4 = st.columns(4)
    with dt_col1: ditch_full_pages = st.text_input("Full Pages", value="18-26", key="ditch_full")
    with dt_col2: ditch_top_pages = st.text_input("Top Half Pages", value="", key="ditch_top")
    with dt_col3: ditch_bottom_pages = st.text_input("Bottom Half Pages", value="", key="ditch_bottom")
    with dt_col4: ditch_table_pages_input = st.text_input("Table Pages", value="14", key="ditch_table_p")

    with st.expander("Advanced Margins & Tolerances (Ditch)", expanded=False):
        dt_t1, dt_t2 = st.columns(2)
        with dt_t1:
            ditch_row_tol = st.number_input("Row Tolerance", value=5, key="ditch_row_tol")
            ditch_col_tol = st.number_input("Col Tolerance", value=20, key="ditch_col_tol")
            ditch_skip_header = st.number_input("Skip Header Rows", value=3, step=1, key="ditch_skip")
        with dt_t2:
            ditch_search_w = st.number_input("Search Width Mult", value=2.5, key="ditch_sw")
            ditch_search_h = st.number_input("Search Height Mult", value=1.15, key="ditch_sh")

        dm_t1, dm_t2 = st.columns(2)
        with dm_t1:
            ditch_margin_left = st.number_input("Table Left Margin", value=1.5, step=0.1, key="ditch_t_left")
            ditch_margin_right = st.number_input("Table Right Margin", value=2.5, step=0.1, key="ditch_t_right")
            ditch_margin_top = st.number_input("Table Top Margin", value=0.5, step=0.1, key="ditch_t_top")
            ditch_margin_bottom = st.number_input("Table Bottom Margin", value=0.5, step=0.1, key="ditch_t_bot")
        with dm_t2:
            ditch_p_margin_left = st.number_input("Plan Left Margin", value=0.0, step=0.1, key="ditch_p_left")
            ditch_p_margin_right = st.number_input("Plan Right Margin", value=0.0, step=0.1, key="ditch_p_right")
            ditch_p_margin_top = st.number_input("Plan Top Margin", value=0.0, step=0.1, key="ditch_p_top")
            ditch_p_margin_bottom = st.number_input("Plan Bottom Margin", value=0.0, step=0.1, key="ditch_p_bot")

    if st.button("Run Ditch Comparison", type="primary", key="btn_run_ditch"):
        # --- UPDATED: Capturing dynamically chosen config ---
        col_records = ditch_edited_col_map.to_dict('records')
        ditch_target_var = next((row['Variable'] for row in col_records if str(row.get('Annotation_Location', '')).strip().lower() == 'target'), None)

        if not uploaded_ditch_pdf:
            st.error("Please upload a Ditch Plan PDF.")
        elif not ditch_target_var:
            st.error("Error: You must assign exactly one variable's Annotation_Location to 'Target'.")
        else:
            for c in col_records:
                c['parsed_loc'] = parse_annotation_loc(c.get('Annotation_Location', 'None'))
            
            ditch_overrides = {}
            for p in parse_page_ranges(ditch_full_pages): ditch_overrides[p] = "full"
            for p in parse_page_ranges(ditch_top_pages): ditch_overrides[p] = "top"
            for p in parse_page_ranges(ditch_bottom_pages): ditch_overrides[p] = "bottom"

            alias_records = ditch_edited_aliases.to_dict('records')
            alias_dict = {str(r['Plan Format']): str(r['Table Format']) for r in alias_records if r.get('Plan Format')}

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
                tmp_in.write(uploaded_ditch_pdf.read())
                tmp_in_path = tmp_in.name

            # out_xl = tempfile.mktemp(suffix=".xlsx")
            # out_pdf = tempfile.mktemp(suffix=".pdf")
            out_xl = get_safe_temp_path(".xlsx")
            out_pdf = get_safe_temp_path(".pdf")

            with st.spinner("Processing Ditch Layout..."):
                try:
                    process_and_compare_ditch_layout(
                        pdf_path=tmp_in_path,
                        table_pages=parse_page_ranges(ditch_table_pages_input),
                        plan_page_config={"default": "full", "overrides": ditch_overrides},
                        dynamic_config=col_records, # --- UPDATED: Use dynamic location mapping --- 
                        target_var=ditch_target_var, 
                        baseline_aliases=alias_dict,
                        plan_margin_left=ditch_p_margin_left, plan_margin_right=ditch_p_margin_right,
                        plan_margin_top=ditch_p_margin_top, plan_margin_bottom=ditch_p_margin_bottom,
                        margin_left_in=ditch_margin_left, margin_right_in=ditch_margin_right,
                        margin_top_in=ditch_margin_top, margin_bottom_in=ditch_margin_bottom,
                        row_tolerance=ditch_row_tol, col_tolerance=ditch_col_tol,
                        search_width_mult=ditch_search_w, search_height_mult=ditch_search_h,
                        skip_header_rows=int(ditch_skip_header),
                        structure_pattern=ditch_structure_pattern,
                        orientation="vertical_b2t" if "Vertical" in ditch_orientation else "horizontal",
                        x_separators=ditch_visual_splits,
                        output_excel=out_xl, marked_pdf_output=out_pdf
                    )
                    st.session_state.ditch_excel = open(out_xl, "rb").read()
                    st.session_state.ditch_pdf = open(out_pdf, "rb").read()
                    st.session_state.ditch_success = True
                    st.success("Ditch layout comparison completed!")
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.session_state.ditch_success = False
                finally:
                    for p in [tmp_in_path, out_xl, out_pdf]:
                        if os.path.exists(p): os.remove(p)

    if st.session_state.ditch_success:
        st.write("### Download Ditch Results")
        dtd1, dtd2 = st.columns(2)
        dtd1.download_button("📥 Download Ditch Excel", st.session_state.ditch_excel, "ditch_layout_comparison.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_ditch_xl")
        dtd2.download_button("📥 Download Marked PDF", st.session_state.ditch_pdf, "marked_ditch_output.pdf", "application/pdf", key="dl_ditch_pdf")

# =============================================================================
# TAB 6: VISUALIZATION DASHBOARD
# =============================================================================
import tempfile
import pandas as pd
import streamlit as st
from dashboard_generator_v1 import generate_high_res_dashboard


def parse_page_spec(spec_str):
  pages = set()
  if not spec_str:
    return pages
  parts = spec_str.split(",")
  for part in parts:
    part = part.strip()
    if "-" in part:
      subparts = part.split("-")
      if (
          len(subparts) == 2
          and subparts[0].strip().isdigit()
          and subparts[1].strip().isdigit()
      ):
        start, end = int(subparts[0].strip()), int(subparts[1].strip())
        pages.update(range(start, end + 1))
    elif part.isdigit():
      pages.add(int(part))
  return pages


with tab_vis:
  st.header("📊 High-Resolution Dual-Viewer Dashboard")
  st.markdown(
      "Generate a self-contained interactive audit dashboard strictly matching"
      " the v3 analysis."
  )

  col1, col2 = st.columns(2)
  with col1:
    uploaded_excel = st.file_uploader(
        "Upload Comparison Excel File (.xlsx)", type=["xlsx"], key="vis_excel_up"
    )
  with col2:
    uploaded_pdf = st.file_uploader(
        "Upload Marked PDF File (.pdf)", type=["pdf"], key="vis_pdf_up"
    )

  if uploaded_excel and uploaded_pdf:
    xls = pd.ExcelFile(uploaded_excel)
    sheet_names = xls.sheet_names

    st.subheader("📋 Select Excel Sheets")
    c1, c2, c3 = st.columns(3)
    with c1:
      main_sheet_name = st.selectbox(
          "Main Comparison Sheet",
          sheet_names,
          index=(
              sheet_names.index("Comparison_Summary")
              if "Comparison_Summary" in sheet_names
              else 0
          ),
          key="vis_main_sheet",
      )
    with c2:
      plan_extraction_sheet = st.selectbox(
          "Plan/Profile Extraction Sheet",
          sheet_names,
          index=(
              sheet_names.index("Plan_Extraction")
              if "Plan_Extraction" in sheet_names
              else (
                  sheet_names.index("Profile_Extraction")
                  if "Profile_Extraction" in sheet_names
                  else 0
              )
          ),
          key="vis_plan_sheet",
      )
    with c3:
      table_extraction_sheet = st.selectbox(
          "Table Extraction Sheet",
          sheet_names,
          index=(
              sheet_names.index("Table_Extraction")
              if "Table_Extraction" in sheet_names
              else 0
          ),
          key="vis_table_sheet",
      )

    
    col1, col2 = st.columns(2)
    
    with col1:
        table_name_column = st.text_input("Target Column", value="Col_1")
    
    with col2:
        col_tol_input = st.number_input("Column Tolerance (pts)", value=27, min_value=1, step=1)
    
    
    st.subheader("✂️ Profile Page Configuration (Spatial Bounds)")
    st.markdown(
        "Configure the default search boundaries for a specific range of pages,"
        " then apply exceptions (overrides)."
    )

    cfg_c1, cfg_c2, cfg_c3 = st.columns(3)
    with cfg_c1:
      range_start = st.number_input(
          "Range Start Page", value=18, min_value=1, key="vis_range_start"
      )
    with cfg_c2:
      range_end = st.number_input(
          "Range End Page", value=26, min_value=1, key="vis_range_end"
      )
    with cfg_c3:
      default_mode = st.selectbox(
          "Default Scan Mode (For Range)",
          ["full", "top", "bottom"],
          index=2,
          key="vis_default_mode",
      )

    st.markdown("**Page Overrides (Exceptions to the rule)**")
    o_col1, o_col2, o_col3 = st.columns(3)
    with o_col1:
      full_pages = st.text_input(
          "Force Full Pages (e.g. 27, 28)", value="27, 28", key="vis_plan_full"
      )
    with o_col2:
      top_pages = st.text_input(
          "Force Top Half Pages", value="", key="vis_plan_top"
      )
    with o_col3:
      bottom_pages = st.text_input(
          "Force Bottom Half Pages", value="", key="vis_plan_bottom"
      )

    dashboard_title = st.text_input(
        "Dashboard Title",
        value="Profile Layout Audit Dashboard",
        key="vis_dash_title",
    )

    if st.button(
        "🚀 Generate High-Resolution Dashboard",
        type="primary",
        key="vis_gen_btn",
    ):
      overrides = {}
      for p in parse_page_spec(full_pages):
        overrides[p] = "full"
      for p in parse_page_spec(top_pages):
        overrides[p] = "top"
      for p in parse_page_spec(bottom_pages):
        overrides[p] = "bottom"

      profile_page_config = {
          "range": (int(range_start), int(range_end)),
          "default": default_mode,
          "overrides": overrides,
      }

      # with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_exc:
      #   tmp_exc.write(uploaded_excel.getvalue())
      #   exc_path = tmp_exc.name

      # with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
      #   tmp_pdf.write(uploaded_pdf.getvalue())
      #   pdf_path = tmp_pdf.name

      # output_html_path = "audit_dashboard_output.html"

      with st.spinner(
          "Generating high-resolution dashboard and rendering images..."
      ):
      
          try:
              # Pass raw bytes directly to the updated function
              html_content = generate_high_res_dashboard(
                  excel_bytes=uploaded_excel.getvalue(),
                  pdf_bytes=uploaded_pdf.getvalue(),
                  dashboard_title=dashboard_title,
                  main_sheet_name=main_sheet_name,
                  plan_extraction_sheet=plan_extraction_sheet,
                  table_extraction_sheet=table_extraction_sheet,
                  table_name_column=table_name_column,
                  page_config=profile_page_config,
                  col_tol_input=col_tol_input,
                  render_dpi=300,
              )
    
              st.success("Dashboard generated successfully!")
              # Serve the returned HTML string straight to the download button
              st.download_button(
                  label="📥 Download Audit Dashboard HTML",
                  data=html_content,
                  file_name="audit_dashboard.html",
                  mime="text/html",
                  key="vis_download_btn",
              )
          except Exception as e:
                st.error(e)