import os
import re
import tempfile
import fitz  # PyMuPDF
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill

# =============================================================================
# 1. TEXT SANITIZATION & NORMALIZATION HELPERS
# =============================================================================

ILLEGAL_XML_CHARS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1F\uD800-\uDFFF\uFFFE\uFFFF]')

def sanitize_excel_text(text):
    if not text or pd.isna(text):
        return ""
    text = str(text)
    text = text.replace("℄", "CL ").replace("ℬ", "BL ")
    text = text.replace("¢", "CL ").replace("Q ", "CL ")
    text = re.sub(r'\bC\s*/\s*L\b', 'CL', text, flags=re.IGNORECASE)
    text = re.sub(r'\bB\s*/\s*L\b', 'BL', text, flags=re.IGNORECASE)
    text = ILLEGAL_XML_CHARS_RE.sub('', text).strip()
    if text.startswith(('=', '+', '-', '@')):
        text = "'" + text
    return text

def normalize_baseline_strict(val):
    if not val or pd.isna(val):
        return ""
    s = str(val).lstrip("'").strip().upper()
    s = ILLEGAL_XML_CHARS_RE.sub('', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def prepare_alias_dict(raw_alias_dict):
    if not raw_alias_dict:
        return {}
    bidirectional_map = {}
    for k, v in raw_alias_dict.items():
        k_clean = normalize_baseline_strict(k)
        v_clean = normalize_baseline_strict(v)
        if k_clean and v_clean:
            canonical = v_clean
            bidirectional_map[k_clean] = canonical
            bidirectional_map[v_clean] = canonical
    return bidirectional_map

def get_aliased_baseline(val, alias_map):
    clean_val = normalize_baseline_strict(val)
    if not alias_map:
        return clean_val
    return alias_map.get(clean_val, clean_val)

def clean_and_format_offset(text):
    if not text:
        return ""
    t = str(text).strip().lstrip("'")
    is_negative = t.startswith("-")
    dir_pattern = r"(RT|LT|RIGHT|LEFT|\bR\b|\bL\b|LI|RI|1T|PT|L1|R1|L\s*T|R\s*T|I\s*T|1\s*T)"
    pattern = re.compile(r"(-?\d{1,5}(?:\.\d+)?)\s*(?:FT|FEET)?\s*['’\"`,\.\s]*\(?\s*" + dir_pattern + r"?\s*\)?", re.IGNORECASE)
    match = pattern.search(t)
    if match:
        num_str = match.group(1)
        raw_dir = re.sub(r"\s+", "", (match.group(2) or "")).upper()
        num_val = abs(float(num_str))
        if raw_dir in ["LT", "LEFT", "L", "LI", "1T", "L1", "IT"] or is_negative or num_str.startswith("-"):
            direction = "LT"
        elif raw_dir in ["RT", "RIGHT", "R", "RI", "PT", "R1"]:
            direction = "RT"
        else:
            direction = ""
        return f"{num_val:.2f}'{' ' + direction if direction else ''}".strip()
    return t

def clean_and_format_station(text):
    if not text:
        return ""
    t = str(text).strip().lstrip("'")
    fixed_text = re.sub(r'(\d{1,5})\s*[tT1iI]\s*(\d{2})', r'\1+\2', t)
    match = re.search(r"(?:STA\.?\s*)?(\d{1,5}\+\d{2}(?:\.\d+)?)", fixed_text, re.IGNORECASE)
    if match:
        return f"STA {match.group(1)}"
    return t

def normalize_baseline(baseline_val):
    if not baseline_val:
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(baseline_val).lstrip("'").upper().strip())

def normalize_structure_name(name_val):
    if not name_val:
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(name_val).lstrip("'").upper())

def clean_and_format_slope(text):
    if not text:
        return ""
    t = str(text).strip().lstrip("'")
    match = re.search(r"(@\s*)?(\d{1,4}(?:\.\d+)?)\s*%", t)
    if match:
        return f"@ {float(match.group(2)):.2f}%"
    return t

def clean_and_format_length(text):
    if not text:
        return ""
    t = str(text).strip().lstrip("'")
    match = re.search(r"(\d{1,6}(?:\.\d+)?)\s*(?:LF|FEET)?", t, re.IGNORECASE)
    if match:
        return f"{float(match.group(1)):.2f} LF"
    return t

def clean_and_format_area(text):
    if not text:
        return ""
    t = str(text).strip().lstrip("'")
    match = re.search(r"(\d{1,4}\.\d{2}|\d+\.\d+|\d+)", t)
    if match:
        return f"{float(match.group(1)):.2f}"
    return t

def normalize_size_type(text):
    if not text:
        return ""
    t = str(text).strip().lstrip("'").upper()
    return re.sub(r'\s+', ' ', t)

def clean_and_format_elevation(text):
    if not text: return ""
    t = str(text).strip().lstrip("'")
    match = re.search(r'(?:TOP\s*EL\.?|EL\.?|ELEV\.?)\s*(-?\d{1,6}(?:\.\d+)?)', t, re.IGNORECASE)
    if match: return f"TOP EL {match.group(1)}"
    match = re.search(r'(-?\d{2,6}\.\d{1,2})', t)
    if match: return f"TOP EL {match.group(1)}"
    return t

def parse_station_to_numeric(sta_val):
    if not sta_val or pd.isna(sta_val):
        return None
    s = re.sub(r'(\d+)\s*[tT1iI]\s*(\d+)', r'\1+\2', str(sta_val).strip().lstrip("'").upper())
    m = re.search(r'(\d+)\+(\d+(?:\.\d+)?)', s)
    if m:
        try:
            return (float(m.group(1)) * 100.0) + float(m.group(2))
        except ValueError:
            return None
    num_match = re.search(r'(\d+(?:\.\d+)?)', s)
    return float(num_match.group(1)) if num_match else None

def parse_offset_to_signed_numeric(off_val):
    if not off_val or pd.isna(off_val):
        return None
    val_str = str(off_val).strip().upper()
    is_lt = 'LT' in val_str or 'LEFT' in val_str or 'L' in val_str
    is_rt = 'RT' in val_str or 'RIGHT' in val_str or 'R' in val_str
    m = re.search(r'(-?\d+(?:\.\d+)?)', val_str)
    if not m:
        return None
    try:
        num = float(m.group(1))
    except ValueError:
        return None
    if is_lt:
        return -abs(num)
    elif is_rt:
        return abs(num)
    else:
        return num

def parse_offset_to_numeric(offset_val):
    return parse_offset_to_signed_numeric(offset_val)

def parse_elevation_to_numeric(elevation_val):
    if elevation_val is None or pd.isna(elevation_val):
        return None
    s = str(elevation_val).strip().lstrip("'").upper()
    s = re.sub(r'^(TOP\s*EL\.?|EL\.?|ELEV\.?)\s*', '', s)
    match = re.search(r'(-?\d{1,6}(?:\.\d+)?)', s)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    return None

def parse_slope_to_numeric(val):
    if val is None or pd.isna(val):
        return None
    match = re.search(r"(\d+(?:\.\d+)?)", str(val))
    return float(match.group(1)) if match else None

def parse_length_to_numeric(val):
    if val is None or pd.isna(val):
        return None
    match = re.search(r"(\d+(?:\.\d+)?)", str(val))
    return float(match.group(1)) if match else None

def parse_area_to_numeric(val):
    if val is None or pd.isna(val):
        return None
    match = re.search(r"(\d+(?:\.\d+)?)", str(val))
    return float(match.group(1)) if match else None

def apply_heuristic_formatting(var_name, text):
    v = var_name.lower()
    if "offset" in v:
        return clean_and_format_offset(text) or text
    if "station" in v or "sta" in v:
        return clean_and_format_station(text) or text
    if "baseline" in v:
        return normalize_baseline(text)
    if "slope" in v:
        return clean_and_format_slope(text)
    if "length" in v:
        return clean_and_format_length(text)
    if "size" in v:
        return normalize_size_type(text)
    if "area" in v:
        return clean_and_format_area(text)
    if "elevation" in v or "elev" in v or "el" == v or "top_el" in v:
        return clean_and_format_elevation(text)
    return sanitize_excel_text(text)

# =============================================================================
# 2. SPATIAL & TABLE EXTRACTION ENGINES
# =============================================================================

def group_blocks_into_lines(blocks, y_tolerance=5):
    if not blocks:
        return []
    sorted_blocks = sorted(blocks, key=lambda b: b['y_center'])
    lines, current_row, current_y = [], [], None

    for b in sorted_blocks:
        if current_y is None:
            current_y = b['y_center']
            current_row = [b]
        elif abs(b['y_center'] - current_y) <= y_tolerance:
            current_row.append(b)
            current_y = sum(i['y_center'] for i in current_row) / len(current_row)
        else:
            lines.append(current_row)
            current_row = [b]
            current_y = b['y_center']
    if current_row:
        lines.append(current_row)

    line_objects = []
    for line_blocks in lines:
        line_blocks.sort(key=lambda item: item['x0'])
        line_objects.append({
            'text': " ".join(item['text'] for item in line_blocks),
            'y_center': sum(item['y_center'] for item in line_blocks) / len(line_blocks),
            'blocks': line_blocks
        })
    return line_objects


def group_blocks_into_vertical_lines(blocks, x_tolerance=5):
    """Groups text blocks into vertical columns, reading bottom-to-top."""
    if not blocks:
        return []
    
    # Sort blocks by their X-center to group them into vertical "lines" (columns)
    blocks.sort(key=lambda b: b['x_center'])
    lines, current_col, current_x = [], [], None

    for b in blocks:
        if current_x is None:
            current_x = b['x_center']
            current_col = [b]
        elif abs(b['x_center'] - current_x) <= x_tolerance:
            current_col.append(b)
            current_x = sum(i['x_center'] for i in current_col) / len(current_col)
        else:
            lines.append(current_col)
            current_col = [b]
            current_x = b['x_center']
    if current_col:
        lines.append(current_col)

    line_objects = []
    for col_blocks in lines:
        # For bottom-to-top text, read by sorting from highest Y to lowest Y
        col_blocks.sort(key=lambda item: -item['y_center'])
        line_objects.append({
            'text': " ".join(item['text'] for item in col_blocks),
            'x_center': sum(item['x_center'] for item in col_blocks) / len(col_blocks),
            'blocks': col_blocks
        })
    return line_objects



def group_ocr_into_table(words, row_tolerance=5, col_tolerance=20, skip_header_rows=3, x_separators=None, page_width=None):
    if not words:
        return pd.DataFrame(), []
        
    items = []
    for w in words:
        text = sanitize_excel_text(w[4])
        if text:
            items.append({
                'x0': w[0], 'y0': w[1], 'x1': w[2], 'y1': w[3],
                'x_center': (w[0] + w[2]) / 2.0, 'y_center': (w[1] + w[3]) / 2.0,
                'text': text
            })
    if not items:
        return pd.DataFrame(), []

    # 1. Y-axis Row Grouping
    items.sort(key=lambda item: item['y_center'])
    rows, current_row, current_y = [], [], None

    for item in items:
        if current_y is None:
            current_y, current_row = item['y_center'], [item]
        elif abs(item['y_center'] - current_y) <= row_tolerance:
            current_row.append(item)
            current_y = sum(i['y_center'] for i in current_row) / len(current_row)
        else:
            rows.append(current_row)
            current_row, current_y = [item], item['y_center']
    if current_row:
        rows.append(current_row)

    for row in rows:
        row.sort(key=lambda i: i['x_center'])

    grid_data, row_bboxes_pt = [], []

    # 2. X-axis Column Grouping
    if x_separators and len(x_separators) > 0:
        # Determine native reference page width in PDF points
        ref_width = page_width if page_width else max(item['x1'] for item in items)
        
        # Convert percentages (0.0-1.0 or 0-100%) into native PDF points
        converted_separators = []
        for sep in x_separators:
            if 0.0 <= sep <= 1.0:
                converted_separators.append(sep * ref_width)
            elif 1.0 < sep <= 100.0 and ref_width > 100.0:
                converted_separators.append((sep / 100.0) * ref_width)
            else:
                converted_separators.append(sep)

        boundaries = [0] + sorted(converted_separators) + [float('inf')]
        
        for row in rows:
            if not row: continue
            row_dict = {f"Col_{i+1}": "" for i in range(len(boundaries) - 1)}
            rx0, ry0, rx1, ry1 = min(i['x0'] for i in row), min(i['y0'] for i in row), max(i['x1'] for i in row), max(i['y1'] for i in row)
            
            for item in row:
                for col_idx in range(len(boundaries) - 1):
                    if boundaries[col_idx] <= item['x_center'] < boundaries[col_idx + 1]:
                        col_key = f"Col_{col_idx + 1}"
                        row_dict[col_key] = (row_dict[col_key] + " " + item['text']).strip() if row_dict[col_key] else item['text']
                        break
                        
            grid_data.append(row_dict)
            row_bboxes_pt.append([rx0, ry0, rx1, ry1])

    else:
        # FALLBACK MODE: Dynamic col_tolerance grouping
        col_centers = []
        for x in sorted([item['x_center'] for item in items]):
            if not col_centers:
                col_centers.append(x)
            else:
                matched = False
                for idx, c in enumerate(col_centers):
                    if abs(x - c) <= col_tolerance:
                        col_centers[idx], matched = (c + x) / 2.0, True
                        break
                if not matched:
                    col_centers.append(x)
        col_centers.sort()

        for row in rows:
            if not row: continue
            row_dict = {f"Col_{i+1}": "" for i in range(len(col_centers))}
            rx0, ry0, rx1, ry1 = min(i['x0'] for i in row), min(i['y0'] for i in row), max(i['x1'] for i in row), max(i['y1'] for i in row)
            for item in row:
                best_col_idx = min(range(len(col_centers)), key=lambda i: abs(item['x_center'] - col_centers[i]))
                col_key = f"Col_{best_col_idx + 1}"
                row_dict[col_key] = (row_dict[col_key] + " " + item['text']).strip() if row_dict[col_key] else item['text']
            grid_data.append(row_dict)
            row_bboxes_pt.append([rx0, ry0, rx1, ry1])

    # 3. Final DataFrame Formatting
    df = pd.DataFrame(grid_data)
    if skip_header_rows > 0 and len(df) > skip_header_rows:
        df = df.iloc[skip_header_rows:].reset_index(drop=True)
        row_bboxes_pt = row_bboxes_pt[skip_header_rows:]

    # PRUNE EMPTY BOUNDARY COLUMNS & RE-INDEX POPULATED COLUMNS
    if not df.empty and x_separators and len(x_separators) > 0:
        non_empty_cols = [c for c in df.columns if df[c].astype(str).str.strip().ne("").any()]
        if non_empty_cols:
            df = df[non_empty_cols]
            df.columns = [f"Col_{i+1}" for i in range(len(df.columns))]
        
    return df, row_bboxes_pt


def draw_translucent_highlighter(page, bbox_pts, color_rgb, fill_opacity=0.35):
    rect = fitz.Rect(bbox_pts[0], bbox_pts[1], bbox_pts[2], bbox_pts[3])
    shape = page.new_shape()
    shape.draw_rect(rect)
    shape.finish(color=None, fill=color_rgb, fill_opacity=fill_opacity)
    shape.commit()

def parse_callout_stack_dynamic(words, structure_pattern, dynamic_config, target_var, x_width_mult=3.0, y_height_mult=1.25, orientation="horizontal"):
    if not words:
        return []
    blocks = []
    for w in words:
        clean_txt = sanitize_excel_text(w[4])
        if clean_txt:
            blocks.append({
                'text': clean_txt, 'x_center': (w[0] + w[2]) / 2.0, 'y_center': (w[1] + w[3]) / 2.0,
                'y0': w[1], 'y1': w[3], 'x0': w[0], 'x1': w[2], 'bbox_pts': [w[0], w[1], w[2], w[3]]
            })

    target_pattern_regex = re.compile(structure_pattern, re.IGNORECASE)
    structure_targets = [b for b in blocks if target_pattern_regex.search(b['text'])]
    extracted_records = []

    for target in structure_targets:
        cx, cy = target['x_center'], target['y_center']
        
        if orientation == "horizontal":
            search_w = (target['x1'] - target['x0']) * x_width_mult
            step_h = max(8.0, target['y1'] - target['y0']) * y_height_mult

            corridor_blocks = [
                b for b in blocks if (cx - (search_w / 2.0) <= b['x_center'] <= cx + (search_w / 2.0))
                and b != target and (cy - (6 * step_h) <= b['y_center'] <= cy + (6 * step_h))
            ]

            above_lines = group_blocks_into_lines([b for b in corridor_blocks if b['y_center'] < cy], y_tolerance=4)
            below_lines = group_blocks_into_lines([b for b in corridor_blocks if b['y_center'] > cy], y_tolerance=4)
            above_lines.sort(key=lambda line: cy - line['y_center'])
            below_lines.sort(key=lambda line: line['y_center'] - cy)
            
        elif orientation == "vertical_b2t":
            # For vertical text, Width and Height logic is inverted
            search_h = (target['y1'] - target['y0']) * x_width_mult
            step_w = max(8.0, target['x1'] - target['x0']) * y_height_mult

            corridor_blocks = [
                b for b in blocks if (cy - (search_h / 2.0) <= b['y_center'] <= cy + (search_h / 2.0))
                and b != target and (cx - (6 * step_w) <= b['x_center'] <= cx + (6 * step_w))
            ]
            
            # Bottom-to-top text means "Above" reading order is physically to the RIGHT (higher X)
            # and "Below" reading order is physically to the LEFT (lower X)
            above_lines = group_blocks_into_vertical_lines([b for b in corridor_blocks if b['x_center'] > cx], x_tolerance=4)
            below_lines = group_blocks_into_vertical_lines([b for b in corridor_blocks if b['x_center'] < cx], x_tolerance=4)
            above_lines.sort(key=lambda line: line['x_center'] - cx)
            below_lines.sort(key=lambda line: cx - line['x_center'])

        record = {target_var: target['text'], "BBox_Pts": target['bbox_pts']}

        for config in dynamic_config:
            var_name = config['Variable']
            if var_name == target_var:
                continue

            raw_text = ""
            loc = config['parsed_loc']
            if loc['type'] == 'relative':
                idx = loc['index']
                if loc['dir'] == 'above' and len(above_lines) > idx:
                    raw_text = above_lines[idx]['text']
                elif loc['dir'] == 'below' and len(below_lines) > idx:
                    raw_text = below_lines[idx]['text']

            record[var_name] = apply_heuristic_formatting(var_name, raw_text)

        extracted_records.append(record)
    return extracted_records


# =============================================================================
# 3. PIPELINE comparison ENGINES
# =============================================================================

def process_and_compare_pdfs(
    pdf_path, table_pages, plan_page_config, dynamic_config, target_var,
    baseline_aliases=None,
    plan_margin_left=0.0, plan_margin_right=0.0, plan_margin_top=0.0, plan_margin_bottom=0.0,
    margin_left_in=1.5, margin_right_in=2.5, margin_top_in=0.5, margin_bottom_in=0.5,
    row_tolerance=5, col_tolerance=20, search_width_mult=3.0, search_height_mult=1.25,
    skip_header_rows=3, structure_pattern=r"\b[A-Za-z]{2}-[A-Za-z](?![PCp c])[A-Za-z]-\d{1,4}\b",
    x_separators=None,
    output_excel="comparison_summary.xlsx", marked_pdf_output="marked_output.pdf"
):
    if baseline_aliases is None:
        baseline_aliases = {}
    alias_map = prepare_alias_dict(baseline_aliases)

    doc = fitz.open(pdf_path)
    plan_records, table_records = [], []

    overrides = plan_page_config.get("overrides", {})
    target_pages = plan_page_config.get("pages", [])
    all_pages_to_scan = sorted(list(set(target_pages).union(set(overrides.keys()))))

    for p_num in all_pages_to_scan:
        if p_num - 1 < 0 or p_num - 1 >= len(doc):
            continue
        page = doc[p_num - 1]
        portion = str(overrides.get(p_num, plan_page_config.get("default", "full"))).lower()

        pw, ph = page.rect.width, page.rect.height
        px0, py0 = plan_margin_left * 72.0, plan_margin_top * 72.0
        px1, py1 = max(px0 + 72.0, pw - (plan_margin_right * 72.0)), max(py0 + 72.0, ph - (plan_margin_bottom * 72.0))

        if portion == "top":
            py1 = min(py1, ph / 2.0)
        elif portion == "bottom":
            py0 = max(py0, ph / 2.0)

        crop_rect = fitz.Rect(px0, py0, px1, py1)
        words = page.get_text("words", clip=crop_rect)
        records = parse_callout_stack_dynamic(words, structure_pattern, dynamic_config, target_var, search_width_mult, search_height_mult)
        
        for r in records:
            r["Page"] = p_num
            plan_records.append(r)

    for p_num in table_pages:
        if p_num - 1 < 0 or p_num - 1 >= len(doc):
            continue
        page = doc[p_num - 1]
        crop_rect = fitz.Rect(margin_left_in * 72.0, margin_top_in * 72.0, page.rect.width - (margin_right_in * 72.0), page.rect.height - (margin_bottom_in * 72.0))
        words = page.get_text("words", clip=crop_rect)
        df_table, row_bboxes = group_ocr_into_table(words, row_tolerance, col_tolerance, skip_header_rows, x_separators=x_separators, page_width=page.rect.width)

        for idx, row in df_table.iterrows():
            target_col = next((c['Table_Column'] for c in dynamic_config if c['Variable'] == target_var), None)
            name_val = str(row.get(target_col, "")).strip() if target_col else ""
            if not name_val:
                continue

            t_rec = {"Page": p_num, target_var: apply_heuristic_formatting(target_var, name_val), "BBox_Pts": row_bboxes[idx]}
            for config in dynamic_config:
                if config['Variable'] == target_var:
                    continue
                t_rec[config['Variable']] = apply_heuristic_formatting(config['Variable'], str(row.get(config['Table_Column'], "")))
            table_records.append(t_rec)

    plan_dict = {}
    for r in plan_records:
        plan_dict.setdefault(normalize_structure_name(r[target_var]), []).append(r)
    table_dict = {normalize_structure_name(r[target_var]): r for r in table_records}

    all_keys = sorted(list(set(plan_dict.keys()).union(set(table_dict.keys()))))
    comparison_rows = []
    other_vars = [c['Variable'] for c in dynamic_config if c['Variable'] != target_var]

    for key in all_keys:
        p_items, t_item = plan_dict.get(key, []), table_dict.get(key)
        if p_items and t_item:
            overall_status, pages_list = "MATCH", []
            for p_item in p_items:
                pages_list.append(str(p_item["Page"]))
                item_match = True
                for var in other_vars:
                    pval, tval = p_item.get(var, ""), t_item.get(var, "")
                    if "station" in var.lower():
                        pn, tn = parse_station_to_numeric(pval), parse_station_to_numeric(tval)
                        if pn is not None and tn is not None:
                            if abs(pn - tn) > 0.1: item_match = False
                        elif pval != tval: item_match = False
                    elif "offset" in var.lower():
                        pn, tn = parse_offset_to_numeric(pval), parse_offset_to_numeric(tval)
                        if pn is not None and tn is not None:
                            if abs(pn - tn) > 0.1: item_match = False
                        elif pval != tval: item_match = False
                    elif "slope" in var.lower():
                        pn, tn = parse_slope_to_numeric(pval), parse_slope_to_numeric(tval)
                        if pn is not None and tn is not None:
                            if abs(pn - tn) > 0.01: item_match = False
                        elif pval != tval: item_match = False
                    elif "length" in var.lower():
                        pn, tn = parse_length_to_numeric(pval), parse_length_to_numeric(tval)
                        if pn is not None and tn is not None:
                            if abs(pn - tn) > 0.1: item_match = False
                        elif pval != tval: item_match = False
                    elif "baseline" in var.lower():
                        p_base = get_aliased_baseline(pval, alias_map)
                        t_base = get_aliased_baseline(tval, alias_map)
                        if p_base != t_base: item_match = False
                    else:
                        if normalize_baseline(pval) != normalize_baseline(tval): item_match = False
                p_item["Status"] = "MATCH" if item_match else "MISMATCH"
                if not item_match: overall_status = "MISMATCH"

            t_item["Status"] = overall_status
            row_data = {target_var: t_item[target_var], "Status": overall_status, "Plan_Page(s)": ", ".join(sorted(set(pages_list)))}
            for var in other_vars: row_data[f"Plan_{var}"] = p_items[0].get(var, "")
            row_data["Table_Page"] = t_item["Page"]
            for var in other_vars: row_data[f"Table_{var}"] = t_item.get(var, "")
            comparison_rows.append(row_data)
        elif p_items and not t_item:
            for p_item in p_items: p_item["Status"] = "MISSING_IN_TABLE"
            row_data = {target_var: p_items[0][target_var], "Status": "MISSING_IN_TABLE", "Plan_Page(s)": ", ".join(sorted(set(str(p["Page"]) for p in p_items)))}
            for var in other_vars: row_data[f"Plan_{var}"] = p_items[0].get(var, "")
            row_data["Table_Page"] = "-"
            for var in other_vars: row_data[f"Table_{var}"] = "-"
            comparison_rows.append(row_data)
        elif t_item and not p_items:
            t_item["Status"] = "MISSING_IN_PLAN"
            row_data = {target_var: t_item[target_var], "Status": "MISSING_IN_PLAN", "Plan_Page(s)": "-"}
            for var in other_vars: row_data[f"Plan_{var}"] = "-"
            row_data["Table_Page"] = t_item["Page"]
            for var in other_vars: row_data[f"Table_{var}"] = t_item.get(var, "")
            comparison_rows.append(row_data)

    df_comp = pd.DataFrame(comparison_rows).replace(ILLEGAL_XML_CHARS_RE, '', regex=True)
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_comp.to_excel(writer, sheet_name="Comparison_Summary", index=False)
        pd.DataFrame(plan_records).drop(columns=["BBox_Pts"], errors="ignore").to_excel(writer, sheet_name="Plan_Extraction", index=False)
        pd.DataFrame(table_records).drop(columns=["BBox_Pts"], errors="ignore").to_excel(writer, sheet_name="Table_Extraction", index=False)

    wb = openpyxl.load_workbook(output_excel)
    ws = wb["Comparison_Summary"]
    fills = {"MATCH": PatternFill(start_color="C6EFCE", fill_type="solid"), "MISMATCH": PatternFill(start_color="FFC7CE", fill_type="solid")}
    for row in ws.iter_rows(min_row=2):
        fill = fills.get(row[1].value, PatternFill(start_color="FFEB9C", fill_type="solid"))
        for cell in row: cell.fill = fill
    wb.save(output_excel)

    for r_list, status_key in [(plan_records, "MISSING_IN_TABLE"), (table_records, "MISSING_IN_PLAN")]:
        for rec in r_list:
            color = (0.0, 0.8, 0.2) if rec.get("Status") == "MATCH" else ((1.0, 0.2, 0.2) if rec.get("Status") == "MISMATCH" else (1.0, 0.9, 0.1))
            draw_translucent_highlighter(doc[rec["Page"] - 1], rec["BBox_Pts"], color)
    doc.save(marked_pdf_output)
    doc.close()


def process_and_compare_ditch_layout(
    pdf_path, table_pages, plan_page_config, dynamic_config, target_var, baseline_aliases=None,
    plan_margin_left=0.0, plan_margin_right=0.0, plan_margin_top=0.0, plan_margin_bottom=0.0,
    margin_left_in=1.5, margin_right_in=2.5, margin_top_in=0.5, margin_bottom_in=0.5,
    row_tolerance=5, col_tolerance=20, search_width_mult=2.5, search_height_mult=1.15,
    skip_header_rows=3, structure_pattern=r"\b[A-Za-z]{2}-[A-Za-z]{2}-\d{1,4}\b",
    orientation="horizontal",
    x_separators=None,
    output_excel="ditch_layout_comparison.xlsx", marked_pdf_output="marked_ditch_output.pdf"
):
    if baseline_aliases is None:
        baseline_aliases = {}
    alias_map = prepare_alias_dict(baseline_aliases)

    # dynamically map keys based on config
    var_sta = next((c['Variable'] for c in dynamic_config if 'station' in c['Variable'].lower()), 'Station')
    var_off = next((c['Variable'] for c in dynamic_config if 'offset' in c['Variable'].lower()), 'Offset')
    var_base = next((c['Variable'] for c in dynamic_config if 'baseline' in c['Variable'].lower()), 'Baseline')

    col_name_key = next((c['Table_Column'] for c in dynamic_config if c['Variable'] == target_var), 'Col_1')
    col_sta_key = next((c['Table_Column'] for c in dynamic_config if c['Variable'] == var_sta), 'Col_2')
    col_off_key = next((c['Table_Column'] for c in dynamic_config if c['Variable'] == var_off), 'Col_3')
    col_base_key = next((c['Table_Column'] for c in dynamic_config if c['Variable'] == var_base), 'Col_4')

    doc = fitz.open(pdf_path)
    plan_records, table_records = [], []

    default_portion = plan_page_config.get("default", "full").lower()
    overrides = plan_page_config.get("overrides", {})

    target_pages = plan_page_config.get("pages", [])
    if "range" in plan_page_config:
        start_p, end_p = plan_page_config["range"]
        target_pages.extend(list(range(start_p, end_p + 1)))

    all_pages_to_scan = sorted(list(set(target_pages).union(set(overrides.keys()))))

    for p_num in all_pages_to_scan:
        page_idx = p_num - 1
        if page_idx < 0 or page_idx >= len(doc): continue
        page = doc[page_idx]
        page_w, page_h = page.rect.width, page.rect.height

        portion = str(overrides.get(p_num, default_portion)).lower()
        px0, py0 = plan_margin_left * 72.0, plan_margin_top * 72.0
        px1 = max(px0 + 72.0, page_w - (plan_margin_right * 72.0))
        py1 = max(py0 + 72.0, page_h - (plan_margin_bottom * 72.0))

        if portion == "top": py1 = min(py1, page_h / 2.0)
        elif portion == "bottom": py0 = max(py0, page_h / 2.0)

        crop_rect = fitz.Rect(px0, py0, px1, py1)
        words = page.get_text("words", clip=crop_rect)

        # Uses the dynamic engine now!
        records = parse_callout_stack_dynamic(
            words, structure_pattern=structure_pattern,
            dynamic_config=dynamic_config, target_var=target_var,
            x_width_mult=search_width_mult, y_height_mult=search_height_mult,
            orientation=orientation
        )
        
        for r in records:
            r["Page"] = p_num
            # standardizing keys for comparison engine
            #r["Structure_Name"] = r.get(target_var, "")
            r["Station"] = r.get(var_sta, "")
            r["Offset"] = r.get(var_off, "")
            r["Baseline"] = r.get(var_base, "")
            plan_records.append(r)

    for p_num in table_pages:
        page_idx = p_num - 1
        if page_idx < 0 or page_idx >= len(doc): continue
        page = doc[page_idx]

        crop_x0, crop_y0 = margin_left_in * 72.0, margin_top_in * 72.0
        crop_x1 = max(crop_x0 + 72.0, page.rect.width - (margin_right_in * 72.0))
        crop_y1 = max(crop_y0 + 72.0, page.rect.height - (margin_bottom_in * 72.0))
        crop_rect = fitz.Rect(crop_x0, crop_y0, crop_x1, crop_y1)

        words = page.get_text("words", clip=crop_rect)
        df_table, row_bboxes_pt = group_ocr_into_table(words, row_tolerance=row_tolerance, col_tolerance=col_tolerance, skip_header_rows=skip_header_rows, x_separators=x_separators, page_width=page.rect.width)
        
        for idx, row in df_table.iterrows():
            name_val = row.get(col_name_key, "")
            if not name_val or not str(name_val).strip(): continue
            table_records.append({
                "Page": p_num,
                target_var: sanitize_excel_text(str(name_val)),
                "Station_Raw": sanitize_excel_text(str(row.get(col_sta_key, ""))),
                "Offset_Raw": sanitize_excel_text(str(row.get(col_off_key, ""))),
                "Baseline_Raw": sanitize_excel_text(str(row.get(col_base_key, ""))),
                "BBox_Pts": row_bboxes_pt[idx]
            })

    plan_dict = {}
    for r in plan_records:
        key = normalize_baseline_strict(r[target_var]).replace(" ", "")
        plan_dict.setdefault(key, []).append(r)

    table_dict = {
        normalize_baseline_strict(r[target_var]).replace(" ", ""): r
        for r in table_records
    }

    all_keys = sorted(list(set(plan_dict.keys()).union(set(table_dict.keys()))))
    comparison_rows = []

    for key in all_keys:
        p_items, t_item = plan_dict.get(key, []), table_dict.get(key)
        if p_items and t_item:
            t_sta_num = parse_station_to_numeric(t_item["Station_Raw"])
            t_off_num = parse_offset_to_signed_numeric(t_item["Offset_Raw"])
            t_base_clean = get_aliased_baseline(t_item["Baseline_Raw"], alias_map)

            all_match = True
            plan_pages_str, plan_sta_str, plan_off_str, plan_base_str = [], [], [], []

            for p_item in p_items:
                plan_pages_str.append(str(p_item["Page"]))
                plan_sta_str.append(p_item["Station"])
                plan_off_str.append(p_item["Offset"])
                plan_base_str.append(p_item["Baseline"])

                p_sta_num = parse_station_to_numeric(p_item["Station"])
                p_off_num = parse_offset_to_signed_numeric(p_item["Offset"])
                p_base_clean = get_aliased_baseline(p_item["Baseline"], alias_map)

                match_sta = (p_sta_num is not None and t_sta_num is not None and abs(p_sta_num - t_sta_num) <= 0.05)
                match_off = (p_off_num is not None and t_off_num is not None and abs(p_off_num - t_off_num) <= 0.05)
                match_base = (p_base_clean != "" and t_base_clean != "" and p_base_clean == t_base_clean)

                if match_sta and match_off and match_base:
                    p_item["Status"] = "MATCH"
                else:
                    all_match = False
                    p_item["Status"] = "MISMATCH"

            status = "MATCH" if all_match else "MISMATCH"
            t_item["Status"] = status
            comparison_rows.append({
                target_var: t_item[target_var], "Status": status,
                "Plan_Page": ", ".join(sorted(list(set(plan_pages_str)))),
                "Plan_Baseline": " | ".join(sorted(list(set(plan_base_str)))),
                "Plan_Station": " | ".join(sorted(list(set(plan_sta_str)))),
                "Plan_Offset": " | ".join(sorted(list(set(plan_off_str)))),
                "Table_Page": t_item["Page"], "Table_Baseline": t_item["Baseline_Raw"],
                "Table_Station": t_item["Station_Raw"], "Table_Offset": t_item["Offset_Raw"]
            })

        elif p_items and not t_item:
            plan_pages_str, plan_sta_str, plan_off_str, plan_base_str = [], [], [], []
            for p_item in p_items:
                p_item["Status"] = "MISSING_IN_TABLE"
                plan_pages_str.append(str(p_item["Page"]))
                plan_sta_str.append(p_item["Station"])
                plan_off_str.append(p_item["Offset"])
                plan_base_str.append(p_item["Baseline"])

            comparison_rows.append({
                target_var: p_items[0][target_var], "Status": "MISSING_IN_TABLE",
                "Plan_Page": ", ".join(sorted(list(set(plan_pages_str)))),
                "Plan_Baseline": " | ".join(sorted(list(set(plan_base_str)))),
                "Plan_Station": " | ".join(sorted(list(set(plan_sta_str)))),
                "Plan_Offset": " | ".join(sorted(list(set(plan_off_str)))),
                "Table_Page": "-", "Table_Baseline": "-", "Table_Station": "-", "Table_Offset": "-",
            })

        elif t_item and not p_items:
            t_item["Status"] = "MISSING_IN_PLAN"
            comparison_rows.append({
                target_var: t_item[target_var], "Status": "MISSING_IN_PLAN",
                "Plan_Page": "-", "Plan_Baseline": "-", "Plan_Station": "-", "Plan_Offset": "-",
                "Table_Page": t_item["Page"], "Table_Baseline": t_item["Baseline_Raw"],
                "Table_Station": t_item["Station_Raw"], "Table_Offset": t_item["Offset_Raw"]
            })

    df_comp = pd.DataFrame(comparison_rows).replace(ILLEGAL_XML_CHARS_RE, '', regex=True)
    df_plan = pd.DataFrame(plan_records).drop(columns=["BBox_Pts"], errors="ignore").replace(ILLEGAL_XML_CHARS_RE, '', regex=True)
    df_table_out = pd.DataFrame(table_records).drop(columns=["BBox_Pts"], errors="ignore").replace(ILLEGAL_XML_CHARS_RE, '', regex=True)

    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_comp.to_excel(writer, sheet_name="Layout_Comparison", index=False)
        if not df_plan.empty: df_plan.to_excel(writer, sheet_name="Plan_Extraction", index=False)
        if not df_table_out.empty: df_table_out.to_excel(writer, sheet_name="Table_Extraction", index=False)

    wb = openpyxl.load_workbook(output_excel)
    ws = wb["Layout_Comparison"]
    fills = {"MATCH": PatternFill(start_color="C6EFCE", fill_type="solid"), "MISMATCH": PatternFill(start_color="FFC7CE", fill_type="solid")}
    for row in ws.iter_rows(min_row=2):
        fill = fills.get(row[1].value, PatternFill(start_color="FFEB9C", fill_type="solid"))
        for cell in row: cell.fill = fill
    wb.save(output_excel)

    for rec in plan_records + table_records:
        if "BBox_Pts" in rec:
            page = doc[rec["Page"] - 1]
            status = rec.get("Status", "MISSING")
            color = (0.0, 0.8, 0.2) if status == "MATCH" else ((1.0, 0.2, 0.2) if status == "MISMATCH" else (1.0, 0.9, 0.1))
            draw_translucent_highlighter(page, rec["BBox_Pts"], color)

    doc.save(marked_pdf_output)
    doc.close()


def process_and_compare_top_elevations(
    pdf_path, table_pages, plan_page_config, dynamic_config, target_var,
    margin_left_in=1.5, margin_right_in=2.5, margin_top_in=0.5, margin_bottom_in=0.5,
    row_tolerance=5, col_tolerance=20, search_width_mult=2.0, search_height_mult=1.25,
    skip_header_rows=3, structure_pattern=r"\b[A-Za-z]{2}-[A-Za-z]{2}-\d{1,4}\b",
    orientation="horizontal",
    x_separators=None,
    output_excel="profile_elevation_comparison.xlsx", marked_pdf_output="marked_profile_output.pdf"
):
    
    col_name_key = next((c['Table_Column'] for c in dynamic_config if c['Variable'] == target_var), 'Col_1')
    var_elev = next((c['Variable'] for c in dynamic_config if 'elevation' in c['Variable'].lower()), 'Top_Elevation')
    col_elev_key = next((c['Table_Column'] for c in dynamic_config if c['Variable'] == var_elev), 'Col_5')

    doc = fitz.open(pdf_path)
    plan_records, table_records = [], []

    default_portion = plan_page_config.get("default", "full").lower()
    overrides = plan_page_config.get("overrides", {})

    if "range" in plan_page_config:
        start_p, end_p = plan_page_config["range"]
        target_pages = list(range(start_p, end_p + 1))
    elif "pages" in plan_page_config:
        target_pages = plan_page_config["pages"]
    else:
        target_pages = []

    all_pages_to_scan = sorted(list(set(target_pages).union(set(overrides.keys()))))

    for p_num in all_pages_to_scan:
        page_idx = p_num - 1
        if page_idx < 0 or page_idx >= len(doc): continue
        page = doc[page_idx]
        page_w, page_h = page.rect.width, page.rect.height

        portion = str(overrides.get(p_num, default_portion)).lower()
        if portion == "top": crop_rect = fitz.Rect(0, 0, page_w, page_h / 2.0)
        elif portion == "bottom": crop_rect = fitz.Rect(0, page_h / 2.0, page_w, page_h)
        else: crop_rect = fitz.Rect(0, 0, page_w, page_h)

        words = page.get_text("words", clip=crop_rect)

        # Uses the dynamic engine now!
        # records = parse_callout_stack_dynamic(
        #     words, structure_pattern=structure_pattern,
        #     dynamic_config=dynamic_config, target_var=target_var,
        #     x_width_mult=search_width_mult, y_height_mult=search_height_mult
        # )
        
        records = parse_callout_stack_dynamic(
        words, structure_pattern=structure_pattern,
        dynamic_config=dynamic_config, target_var=target_var,
        x_width_mult=search_width_mult, y_height_mult=search_height_mult,
        orientation=orientation
        )

        for r in records:
            r["Page"] = p_num
            # standardizing keys for comparison engine
            #r["Structure_Name"] = r.get(target_var, "")
            r["Top_Elevation"] = r.get(var_elev, "")
            plan_records.append(r)

    for p_num in table_pages:
        page_idx = p_num - 1
        if page_idx < 0 or page_idx >= len(doc): continue

        page = doc[page_idx]
        page_w, page_h = page.rect.width, page.rect.height

        crop_x0, crop_y0 = margin_left_in * 72.0, margin_top_in * 72.0
        crop_x1 = max(crop_x0 + 72.0, page_w - (margin_right_in * 72.0))
        crop_y1 = max(crop_y0 + 72.0, page_h - (margin_bottom_in * 72.0))
        crop_rect = fitz.Rect(crop_x0, crop_y0, crop_x1, crop_y1)

        words = page.get_text("words", clip=crop_rect)
        df_table, row_bboxes_pt = group_ocr_into_table(words, row_tolerance=row_tolerance, col_tolerance=col_tolerance, skip_header_rows=skip_header_rows, x_separators=x_separators, page_width=page.rect.width)

        for idx, row in df_table.iterrows():
            name_val = row.get(col_name_key, "")
            if not name_val or not str(name_val).strip(): continue
            top_el_val = row.get(col_elev_key, "")
            table_records.append({
                "Page": p_num,
                target_var: sanitize_excel_text(str(name_val)),
                "Top_Elevation_Raw": sanitize_excel_text(str(top_el_val)),
                "BBox_Pts": row_bboxes_pt[idx]
            })

    plan_dict = {}
    for r in plan_records:
        key = normalize_structure_name(r[target_var])
        plan_dict.setdefault(key, []).append(r)

    table_dict = {normalize_structure_name(r[target_var]): r for r in table_records}
    all_keys = sorted(list(set(plan_dict.keys()).union(set(table_dict.keys()))))
    comparison_rows = []

    for key in all_keys:
        p_items, t_item = plan_dict.get(key, []), table_dict.get(key)
        if p_items and t_item:
            t_el_num = parse_elevation_to_numeric(t_item["Top_Elevation_Raw"])
            all_match, plan_pages_str, plan_elevations_str = True, [], []

            for p_item in p_items:
                plan_pages_str.append(str(p_item["Page"]))
                plan_elevations_str.append(p_item["Top_Elevation"])
                p_el_num = parse_elevation_to_numeric(p_item["Top_Elevation"])

                if p_el_num is not None and t_el_num is not None and abs(p_el_num - t_el_num) <= 0.02:
                    p_item["Status"] = "MATCH"
                else:
                    all_match = False
                    p_item["Status"] = "MISMATCH"

            status = "MATCH" if all_match else "MISMATCH"
            t_item["Status"] = status
            comparison_rows.append({
                target_var: t_item[target_var], "Status": status,
                "Plan_Page": ", ".join(sorted(list(set(plan_pages_str)))),
                "Plan_Top_Elevation": " | ".join(sorted(list(set(plan_elevations_str)))),
                "Table_Page": t_item["Page"], "Table_Top_Elevation": t_item["Top_Elevation_Raw"],
            })
        elif p_items and not t_item:
            plan_pages_str, plan_elevations_str = [], []
            for p_item in p_items:
                p_item["Status"] = "MISSING_IN_TABLE"
                plan_pages_str.append(str(p_item["Page"]))
                plan_elevations_str.append(p_item["Top_Elevation"])

            comparison_rows.append({
                target_var: p_items[0][target_var], "Status": "MISSING_IN_TABLE",
                "Plan_Page": ", ".join(sorted(list(set(plan_pages_str)))),
                "Plan_Top_Elevation": " | ".join(sorted(list(set(plan_elevations_str)))),
                "Table_Page": "-", "Table_Top_Elevation": "-",
            })
        elif t_item and not p_items:
            t_item["Status"] = "MISSING_IN_PLAN"
            comparison_rows.append({
                target_var: t_item[target_var], "Status": "MISSING_IN_PLAN",
                "Plan_Page": "-", "Plan_Top_Elevation": "-",
                "Table_Page": t_item["Page"], "Table_Top_Elevation": t_item["Top_Elevation_Raw"],
            })

    df_comp = pd.DataFrame(comparison_rows).replace(ILLEGAL_XML_CHARS_RE, '', regex=True)
    df_plan = pd.DataFrame(plan_records).drop(columns=["BBox_Pts"], errors="ignore").replace(ILLEGAL_XML_CHARS_RE, '', regex=True)
    df_table_out = pd.DataFrame(table_records).drop(columns=["BBox_Pts"], errors="ignore").replace(ILLEGAL_XML_CHARS_RE, '', regex=True)

    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_comp.to_excel(writer, sheet_name="Elevation_Comparison", index=False)
        if not df_plan.empty: df_plan.to_excel(writer, sheet_name="Plan_Extraction", index=False)
        if not df_table_out.empty: df_table_out.to_excel(writer, sheet_name="Table_Extraction", index=False)

    wb = openpyxl.load_workbook(output_excel)
    ws = wb["Elevation_Comparison"]
    fills = {"MATCH": PatternFill(start_color="C6EFCE", fill_type="solid"), "MISMATCH": PatternFill(start_color="FFC7CE", fill_type="solid")}
    for row in ws.iter_rows(min_row=2):
        fill = fills.get(row[1].value, PatternFill(start_color="FFEB9C", fill_type="solid"))
        for cell in row: cell.fill = fill
    wb.save(output_excel)

    for rec in plan_records + table_records:
        if "BBox_Pts" in rec:
            page = doc[rec["Page"] - 1]
            status = rec.get("Status", "MISSING")
            color = (0.0, 0.8, 0.2) if status == "MATCH" else ((1.0, 0.2, 0.2) if status == "MISMATCH" else (1.0, 0.9, 0.1))
            draw_translucent_highlighter(page, rec["BBox_Pts"], color)

    doc.save(marked_pdf_output)
    doc.close()

# =============================================================================
# 4. APP PARSING HELPERS
# =============================================================================

def parse_page_ranges(range_str):
    if not range_str or not str(range_str).strip():
        return []
    pages = []
    for part in str(range_str).split(','):
        part = part.strip()
        if '-' in part:
            try:
                s, e = map(int, part.split('-'))
                pages.extend(range(s, e + 1))
            except ValueError:
                continue
        elif part.isdigit():
            pages.append(int(part))
    return pages

def parse_annotation_loc(val):
    val = str(val).strip().lower()
    if val == "target":
        return {"type": "target"}
    if val == "none":
        return {"type": "none"}
    parts = val.split(" ")
    if len(parts) == 2 and parts[0].isdigit() and parts[1] in ["above", "below"]:
        return {"type": "relative", "index": int(parts[0]) - 1, "dir": parts[1]}
    return {"type": "none"}