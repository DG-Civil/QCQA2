# -*- coding: utf-8 -*-
"""
High-Resolution Dual-Viewer Dashboard Generator (Optimized for Speed)
Generates a self-contained HTML audit dashboard from Excel comparisons and Marked PDFs.
Features left-aligned table crops, crisp red bounding rectangles, 300 DPI rendering,
independent table name column mapping, left-sidebar column sorting, full-width bottom table crop viewer,
exception-based page overrides, and Cartesian explosion protection.
"""

import base64
import json
import os
import re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>/*__DASHBOARD_TITLE__*/</title>
    <style>
        :root {
            --bg-dark: #0f172a;
            --bg-panel: #1e293b;
            --border-color: #334155;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --accent: #3b82f6;
            --match: #22c55e;
            --mismatch: #ef4444;
            --warning: #f59e0b;
            --target-red: #ef4444;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { background: var(--bg-dark); color: var(--text-main); height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
        
        header { background: var(--bg-panel); border-bottom: 1px solid var(--border-color); padding: 10px 20px; display: flex; justify-content: space-between; align-items: center; min-height: 50px; flex-shrink: 0; }
        header h1 { font-size: 1.15rem; font-weight: 600; display: flex; align-items: center; gap: 10px; }
        .stats { display: flex; gap: 10px; font-size: 0.8rem; }
        .badge { padding: 3px 8px; border-radius: 10px; font-weight: 600; font-size: 0.7rem; }
        .badge-match { background: rgba(34,197,94,0.2); color: var(--match); }
        .badge-mismatch { background: rgba(239,68,68,0.2); color: var(--mismatch); }
        .badge-warning { background: rgba(245,158,11,0.2); color: var(--warning); }

        .main-container { 
            display: grid; 
            grid-template-columns: 580px 1fr; 
            grid-template-rows: 1fr 240px; 
            flex: 1; 
            height: calc(100vh - 50px); 
            overflow: hidden; 
        }
        
        .sidebar { grid-column: 1; grid-row: 1; background: var(--bg-panel); border-right: 1px solid var(--border-color); display: flex; flex-direction: column; overflow: hidden; }
        .filter-section { padding: 12px; border-bottom: 1px solid var(--border-color); display: flex; flex-direction: column; gap: 8px; flex-shrink: 0; }
        .search-box { width: 100%; padding: 7px 10px; background: var(--bg-dark); border: 1px solid var(--border-color); color: white; border-radius: 6px; outline: none; font-size: 0.8rem; }
        .filter-buttons { display: flex; gap: 4px; flex-wrap: wrap; }
        .filter-btn { flex: 1; min-width: 80px; padding: 6px 4px; background: var(--bg-dark); border: 1px solid var(--border-color); color: var(--text-muted); border-radius: 4px; cursor: pointer; font-size: 0.68rem; font-weight: 600; text-align: center; }
        .filter-btn.active { background: var(--accent); color: white; border-color: var(--accent); }

        .table-wrapper { flex: 1; overflow: auto; }
        table { width: max-content; min-width: 100%; border-collapse: collapse; text-align: left; font-size: 0.75rem; }
        th { background: #111827; position: sticky; top: 0; padding: 8px 10px; z-index: 10; color: var(--text-muted); white-space: nowrap; border-bottom: 1px solid var(--border-color); cursor: pointer; user-select: none; }
        th:hover { background: #1f2937; color: var(--text-main); }
        td { padding: 7px 10px; border-bottom: 1px solid var(--border-color); cursor: pointer; white-space: nowrap; }
        tr:hover td { background: rgba(59,130,246,0.1); }
        tr.selected td { background: rgba(59,130,246,0.25); border-left: 4px solid var(--accent); }

        .plan-viewer-pane { grid-column: 2; grid-row: 1; display: flex; flex-direction: column; background: #090d16; overflow: hidden; position: relative; border-bottom: 1px solid var(--border-color); }
        .table-crop-pane { grid-column: 1 / -1; grid-row: 2; display: flex; flex-direction: column; background: #0d1322; border-top: 1px solid var(--border-color); overflow: hidden; }

        .interactive-viewport { flex: 1; width: 100%; height: 100%; overflow: auto; position: relative; cursor: grab; }
        .interactive-viewport.dragging { cursor: grabbing; }
        
        .img-container { position: relative; transform-origin: top left; }
        .img-container img { position: absolute; top: 0; left: 0; width: 100%; height: 100%; object-fit: contain; transform-origin: top left; }
        
        .crop-header { background: var(--bg-panel); padding: 6px 15px; font-size: 0.75rem; text-transform: uppercase; color: var(--text-muted); font-weight: 600; display: flex; justify-content: space-between; border-bottom: 1px solid var(--border-color); z-index: 10; flex-shrink: 0; }
        
        @keyframes highlight-pulse-red {
            0% { background-color: rgba(239, 68, 68, 0.65); box-shadow: 0 0 22px rgba(239, 68, 68, 0.95); }
            50% { background-color: rgba(239, 68, 68, 0.15); box-shadow: 0 0 10px rgba(239, 68, 68, 0.3); }
            100% { background-color: rgba(239, 68, 68, 0.35); box-shadow: 0 0 15px rgba(239, 68, 68, 0.7); }
        }
        
        .highlight-box { 
            position: absolute; 
            border: 3.5px solid var(--target-red); 
            pointer-events: none; 
            z-index: 20; 
            opacity: 0; 
            border-radius: 4px;
            box-shadow: 0 0 12px rgba(239, 68, 68, 0.8);
        }
        .highlight-flash { 
            animation: highlight-pulse-red 0.5s ease-in-out 3 forwards; 
            opacity: 1 !important; 
            background-color: rgba(239, 68, 68, 0.3);
        }

        .zoom-controls { position: absolute; top: 10px; right: 10px; display: flex; gap: 4px; z-index: 100; background: rgba(15,23,42,0.85); padding: 4px; border-radius: 6px; border: 1px solid var(--border-color); }
        .zoom-btn { background: var(--bg-panel); color: white; border: 1px solid var(--border-color); width: 26px; height: 26px; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 0.8rem; }
        .zoom-btn:hover { background: var(--accent); }

        .table-row-strip { background: var(--bg-panel); padding: 6px 20px; display: flex; justify-content: flex-start; align-items: center; flex-wrap: wrap; gap: 15px; border-top: 1px solid var(--border-color); font-size: 0.75rem; overflow-y: auto; flex-shrink: 0; max-height: 45px;}
        .strip-item span { display: inline-block; font-size: 0.65rem; color: var(--text-muted); text-transform: uppercase; margin-right: 5px; }
        .strip-item strong { color: var(--text-main); }
    </style>
</head>
<body>

    <header>
        <h1>🔍 <span id="dashboard-title-text">/*__DASHBOARD_TITLE__*/</span></h1>
        <div class="stats" id="stats-container"></div>
    </header>

    <div class="main-container">
        <div class="sidebar">
            <div class="filter-section">
                <input type="text" id="search-input" class="search-box" placeholder="Search across all columns..." onkeyup="filterTable()">
                <div class="filter-buttons">
                    <button class="filter-btn active" onclick="setFilter('ALL', this)">All</button>
                    <button class="filter-btn" onclick="setFilter('MATCH', this)">Matches</button>
                    <button class="filter-btn" onclick="setFilter('MISMATCH', this)">Mismatches</button>
                    <button class="filter-btn" onclick="setFilter('MISSING_PLAN', this)">Missing (Plan)</button>
                    <button class="filter-btn" onclick="setFilter('MISSING_TABLE', this)">Missing (Table)</button>
                </div>
            </div>
            <div class="table-wrapper">
                <table id="comparison-table">
                    <thead id="table-head"></thead>
                    <tbody id="table-body"></tbody>
                </table>
            </div>
        </div>

        <div class="plan-viewer-pane">
            <div class="zoom-controls">
                <button class="zoom-btn" onclick="adjustZoom(0.5)">+</button>
                <button class="zoom-btn" onclick="adjustZoom(-0.5)">-</button>
                <button class="zoom-btn" onclick="resetZoom()">⟲</button>
            </div>
            <div class="crop-header">
                <span>Plan / Profile Viewer (Source PDF)</span>
                <span id="plan-page-label" style="color: var(--accent);">Page: -</span>
            </div>
            <div class="interactive-viewport" id="plan-viewport">
                <div class="img-container" id="plan-container">
                    <img id="plan-img" src="" alt="Select a row to view plan">
                    <div id="plan-highlight" class="highlight-box"></div>
                </div>
            </div>
        </div>

        <div class="table-crop-pane">
            <div class="crop-header">
                <span>Schedule Table Row Verification (Source PDF Crop) — Full Width View</span>
                <span id="crop-page-label" style="color: var(--accent);">Table Page: -</span>
            </div>
            <div class="interactive-viewport" id="table-viewport" style="display: flex; justify-content: center; align-items: center;">
                <span id="crop-placeholder" style="color: var(--text-muted); font-style: italic; font-size: 0.8rem; margin: auto;">No table record found for this structure.</span>
                <div class="img-container" id="table-container" style="display: none;">
                    <img id="table-crop-img" src="" alt="Table Row Crop">
                    <div id="table-highlight" class="highlight-box"></div>
                </div>
            </div>
            <div class="table-row-strip" id="dynamic-strip"></div>
        </div>
    </div>

    <script>
        const dashboardData = /*__DASHBOARD_DATA__*/;
        const tableColumns = /*__TABLE_COLUMNS__*/;
        const pageMeta = /*__PAGE_META__*/;

        let currentFilter = 'ALL';
        let selectedIndex = null;
        let currentZoom = 2.5;
        let sortColumn = null;
        let sortAsc = true;

        function setupDraggableViewport(viewportId) {
            const viewport = document.getElementById(viewportId);
            let isDragging = false;
            let startX, startY, scrollLeft, scrollTop;

            viewport.addEventListener('mousedown', (e) => {
                isDragging = true;
                viewport.classList.add('dragging');
                startX = e.pageX - viewport.offsetLeft;
                startY = e.pageY - viewport.offsetTop;
                scrollLeft = viewport.scrollLeft;
                scrollTop = viewport.scrollTop;
            });

            viewport.addEventListener('mousemove', (e) => {
                if (!isDragging) return;
                e.preventDefault();
                const x = e.pageX - viewport.offsetLeft;
                const y = e.pageY - viewport.offsetTop;
                viewport.scrollLeft = scrollLeft - (x - startX);
                viewport.scrollTop = scrollTop - (y - startY);
            });

            const stopDragging = () => {
                isDragging = false;
                viewport.classList.remove('dragging');
            };

            viewport.addEventListener('mouseup', stopDragging);
            viewport.addEventListener('mouseleave', stopDragging);
        }

        setupDraggableViewport('plan-viewport');
        setupDraggableViewport('table-viewport');

        function sortTable(col) {
            if (sortColumn === col) {
                sortAsc = !sortAsc;
            } else {
                sortColumn = col;
                sortAsc = true;
            }
            renderTable();
        }

        function renderTable() {
            const thead = document.getElementById('table-head');
            thead.innerHTML = '';
            const trHead = document.createElement('tr');
            
            tableColumns.forEach(col => {
                const th = document.createElement('th');
                let indicator = '';
                if (sortColumn === col) {
                    indicator = sortAsc ? ' ▲' : ' ▼';
                }
                th.innerText = col + indicator;
                th.onclick = () => sortTable(col);
                trHead.appendChild(th);
            });
            thead.appendChild(trHead);

            const searchVal = document.getElementById('search-input').value.toLowerCase();

            let filteredData = dashboardData.filter((item) => {
                const statusUpper = String(item._status || '').toUpperCase();
                
                if (currentFilter === 'MATCH' && statusUpper !== 'MATCH') return false;
                if (currentFilter === 'MISMATCH' && statusUpper !== 'MISMATCH') return false;
                if (currentFilter === 'MISSING_PLAN' && !(statusUpper.includes('PLAN') || statusUpper.includes('PROFILE'))) return false;
                if (currentFilter === 'MISSING_TABLE' && !statusUpper.includes('TABLE')) return false;

                const rowText = tableColumns.map(col => String(item[col] || '')).join(' ').toLowerCase();
                if (searchVal && !rowText.includes(searchVal)) return false;

                return true;
            });

            if (sortColumn) {
                filteredData.sort((a, b) => {
                    let valA = a[sortColumn] !== undefined && a[sortColumn] !== null ? a[sortColumn] : '';
                    let valB = b[sortColumn] !== undefined && b[sortColumn] !== null ? b[sortColumn] : '';

                    let numA = parseFloat(valA);
                    let numB = parseFloat(valB);
                    if (!isNaN(numA) && !isNaN(numB)) {
                        return sortAsc ? numA - numB : numB - numA;
                    }

                    let strA = String(valA).toLowerCase();
                    let strB = String(valB).toLowerCase();
                    if (strA < strB) return sortAsc ? -1 : 1;
                    if (strA > strB) return sortAsc ? 1 : -1;
                    return 0;
                });
            }

            const tbody = document.getElementById('table-body');
            tbody.innerHTML = '';

            filteredData.forEach((item) => {
                const originalIndex = dashboardData.indexOf(item);
                const statusUpper = String(item._status || '').toUpperCase();

                const tr = document.createElement('tr');
                if (selectedIndex === originalIndex) tr.classList.add('selected');
                tr.onclick = () => selectItem(originalIndex);

                tableColumns.forEach(col => {
                    const td = document.createElement('td');
                    const val = item[col] !== undefined && item[col] !== null ? item[col] : '-';
                    
                    if (col.toUpperCase() === 'STATUS') {
                        let badgeClass = 'badge-warning';
                        if (statusUpper === 'MATCH') badgeClass = 'badge-match';
                        if (statusUpper === 'MISMATCH') badgeClass = 'badge-mismatch';
                        td.innerHTML = `<span class="badge ${badgeClass}">${val}</span>`;
                    } else if (col.toUpperCase().includes('STRUCTURE') || col.toUpperCase() === 'ID') {
                        td.innerHTML = `<strong>${val}</strong>`;
                    } else {
                        td.innerText = val;
                    }
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
            updateStats();
        }

        function updateStats() {
            let matches = dashboardData.filter(i => String(i._status).toUpperCase() === 'MATCH').length;
            let mismatches = dashboardData.filter(i => String(i._status).toUpperCase() === 'MISMATCH').length;
            let missingPlan = dashboardData.filter(i => String(i._status).toUpperCase().includes('PLAN') || String(i._status).toUpperCase().includes('PROFILE')).length;
            let missingTable = dashboardData.filter(i => String(i._status).toUpperCase().includes('TABLE')).length;

            document.getElementById('stats-container').innerHTML = `
                <span>Total: <strong>${dashboardData.length}</strong></span>
                <span class="badge badge-match">Matches: ${matches}</span>
                <span class="badge badge-mismatch">Mismatches: ${mismatches}</span>
                <span class="badge badge-warning">Missing (Plan): ${missingPlan}</span>
                <span class="badge badge-warning">Missing (Table): ${missingTable}</span>
            `;
        }

        function setFilter(filter, btn) {
            currentFilter = filter;
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderTable();
        }

        function filterTable() { renderTable(); }

        function selectItem(index) {
            selectedIndex = index;
            renderTable();

            const item = dashboardData[index];
            currentZoom = 2.5;
            
            const stripContainer = document.getElementById('dynamic-strip');
            stripContainer.innerHTML = '';
            let count = 0;
            for (let col of tableColumns) {
                if (col.toUpperCase() !== 'STATUS' && item[col] && item[col] !== '-') {
                    stripContainer.innerHTML += `<div class="strip-item"><span>${col}:</span><strong>${item[col]}</strong></div>`;
                    count++;
                    if (count >= 12) break; 
                }
            }

            const planImg = document.getElementById('plan-img');
            const planHighlight = document.getElementById('plan-highlight');
            const planPageLabel = document.getElementById('plan-page-label');

            const planPageKey = String(item._plan_page);
            if (item._plan_page !== '-' && pageMeta[planPageKey]) {
                const pInfo = pageMeta[planPageKey];
                planPageLabel.innerText = `Page: ${item._plan_page}`;
                
                const updatePlanView = () => {
                    executeZoomAndHighlight(
                        'plan-viewport', 'plan-container', 'plan-highlight', 
                        item._plan_bbox, pInfo.width, pInfo.height, currentZoom
                    );
                };

                if (planImg.src === pInfo.data_uri && planImg.complete) {
                    updatePlanView();
                } else {
                    planImg.onload = updatePlanView;
                    planImg.src = pInfo.data_uri;
                }
            } else {
                planImg.src = '';
                planPageLabel.innerText = 'Page: -';
                planHighlight.style.opacity = '0';
            }

            const cropContainer = document.getElementById('table-container');
            const cropImg = document.getElementById('table-crop-img');
            const placeholder = document.getElementById('crop-placeholder');
            const cropPageLabel = document.getElementById('crop-page-label');
            const tableViewport = document.getElementById('table-viewport');

            if (item._table_crop_uri) {
                cropContainer.style.display = 'block';
                placeholder.style.display = 'none';
                cropPageLabel.innerText = `Table Page: ${item._table_page}`;
                tableViewport.style.display = 'block';

                const updateTableView = () => {
                    executeZoomAndHighlight(
                        'table-viewport', 'table-container', 'table-highlight', 
                        item._table_bbox, cropImg.naturalWidth, cropImg.naturalHeight, 1.2
                    );
                };

                if (cropImg.src === item._table_crop_uri && cropImg.complete) {
                    updateTableView();
                } else {
                    cropImg.onload = updateTableView;
                    cropImg.src = item._table_crop_uri;
                }
            } else {
                cropContainer.style.display = 'none';
                placeholder.style.display = 'block';
                cropPageLabel.innerText = 'Table Page: -';
                tableViewport.style.display = 'flex'; 
            }
        }

        function executeZoomAndHighlight(viewportId, containerId, highlightId, bbox, natWidth, natHeight, zoomLvl) {
            const viewport = document.getElementById(viewportId);
            const container = document.getElementById(containerId);
            const hb = document.getElementById(highlightId);

            if (!bbox || bbox.length < 4 || (bbox[2] === 0 && bbox[3] === 0)) {
                hb.style.opacity = '0';
                return;
            }

            const scaledW = natWidth * zoomLvl;
            const scaledH = natHeight * zoomLvl;
            container.style.width = scaledW + 'px';
            container.style.height = scaledH + 'px';

            const x0 = bbox[0] * zoomLvl;
            const y0 = bbox[1] * zoomLvl;
            const boxW = (bbox[2] - bbox[0]) * zoomLvl;
            const boxH = (bbox[3] - bbox[1]) * zoomLvl;

            hb.style.left = x0 + 'px';
            hb.style.top = y0 + 'px';
            hb.style.width = Math.max(35, boxW) + 'px';
            hb.style.height = Math.max(22, boxH) + 'px';
            
            hb.className = 'highlight-box';
            void hb.offsetWidth;
            hb.classList.add('highlight-flash');

            const targetY = y0 + (boxH / 2);
            viewport.scrollTop = Math.max(0, targetY - (viewport.clientHeight / 2));

            if (viewportId === 'table-viewport') {
                viewport.scrollLeft = Math.max(0, x0 - 20);
            } else {
                const targetX = x0 + (boxW / 2);
                viewport.scrollLeft = Math.max(0, targetX - (viewport.clientWidth / 2));
            }
        }

        function adjustZoom(delta) {
            if (selectedIndex === null) return;
            currentZoom = Math.max(1.0, Math.min(8.0, currentZoom + delta));
            const item = dashboardData[selectedIndex];
            const pInfo = pageMeta[String(item._plan_page)];
            if (pInfo) {
                executeZoomAndHighlight('plan-viewport', 'plan-container', 'plan-highlight', item._plan_bbox, pInfo.width, pInfo.height, currentZoom);
            }
        }

        function resetZoom() {
            if (selectedIndex === null) return;
            currentZoom = 1.0;
            const item = dashboardData[selectedIndex];
            const pInfo = pageMeta[String(item._plan_page)];
            if (pInfo) {
                executeZoomAndHighlight('plan-viewport', 'plan-container', 'plan-highlight', item._plan_bbox, pInfo.width, pInfo.height, currentZoom);
            }
        }

        window.onload = () => {
            renderTable();
            if (dashboardData.length > 0) selectItem(0);
        };
    </script>
</body>
</html>
"""

def group_ocr_into_table(words, row_tolerance=5, col_tolerance=20, skip_header_rows=3):
    """Groups page OCR text elements into a structured 2D table grid."""
    if not words: return pd.DataFrame(), []
    items = []
    for w in words:
        text = str(w[4]).strip()
        if text:
            items.append({
                'x0': w[0], 'y0': w[1], 'x1': w[2], 'y1': w[3],
                'x_center': (w[0] + w[2]) / 2.0, 'y_center': (w[1] + w[3]) / 2.0,
                'text': text
            })
    if not items: return pd.DataFrame(), []

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
    if current_row: rows.append(current_row)

    col_centers = []
    for x in sorted([item['x_center'] for item in items]):
        if not col_centers: col_centers.append(x)
        else:
            matched = False
            for idx, c in enumerate(col_centers):
                if abs(x - c) <= col_tolerance:
                    col_centers[idx] = (c + x) / 2.0
                    matched = True
                    break
            if not matched: col_centers.append(x)
    col_centers.sort()
    
    grid_data, row_bboxes_pt = [], []
    for row in rows:
        row_dict = {f"Col_{i+1}": "" for i in range(len(col_centers))}
        rx0 = min(i['x0'] for i in row)
        ry0 = min(i['y0'] for i in row)
        rx1 = max(i['x1'] for i in row)
        ry1 = max(i['y1'] for i in row)
        for item in row:
            best_col_idx = min(range(len(col_centers)), key=lambda i: abs(item['x_center'] - col_centers[i]))
            col_key = f"Col_{best_col_idx + 1}"
            row_dict[col_key] = (row_dict[col_key] + " " + item['text']).strip() if row_dict[col_key] else item['text']
        grid_data.append(row_dict)
        row_bboxes_pt.append([rx0, ry0, rx1, ry1])

    df = pd.DataFrame(grid_data)
    if skip_header_rows > 0 and len(df) > skip_header_rows:
        df = df.iloc[skip_header_rows:].reset_index(drop=True)
        row_bboxes_pt = row_bboxes_pt[skip_header_rows:]
    return df, row_bboxes_pt


def get_page_search_clip(page, page_num, config):
    if not config:
        return page.rect

    p_range = config.get("range")
    if p_range and not (p_range[0] <= page_num <= p_range[1]):
        return page.rect

    rect = page.rect
    
    overrides = config.get("overrides", {})
    if page_num in overrides:
        mode = overrides[page_num]
        if mode == "full":
            return rect
        elif mode == "top":
            return fitz.Rect(0, 0, rect.width, rect.height / 2)
        elif mode == "bottom":
            return fitz.Rect(0, rect.height / 2, rect.width, rect.height)
    
    default_mode = config.get("default", "full")
    if default_mode == "bottom":
        return fitz.Rect(0, rect.height / 2, rect.width, rect.height)
    elif default_mode == "top":
        return fitz.Rect(0, 0, rect.width, rect.height / 2)
    
    return rect

def resolve_column(df, explicit_col):
    """Resolves the identifier column using explicit user setting or dynamic search."""
    if df.empty: 
        return None
    if explicit_col and explicit_col in df.columns:
        return explicit_col
    for col in df.columns:
        if str(col).upper() in ["STRUCTURE_NAME", "STRUCTURE", "ID", "NODE", "NAME", "COL_1"]:
            return col
    return df.columns[0]

def find_page_column(df):
    """Dynamically finds the page tracking column."""
    if df.empty: 
        return None
    for col in df.columns:
        if "PAGE" in str(col).upper():
            return col
    return None

def generate_high_res_dashboard(
    excel_bytes, 
    pdf_bytes, 
    #output_html="audit_dashboard.html", 
    dashboard_title="Layout Audit Dashboard",
    main_sheet_name="Comparison_Summary",
    plan_extraction_sheet="Profile_Extraction",
    table_extraction_sheet="Table_Extraction",
    table_name_column="Col_1",
    col_tol_input=27,
    page_config=None,
    render_dpi=300
):
    SCALE = render_dpi / 72.0 

    print(f"Loading Excel file: {excel_bytes}")
    xls = pd.ExcelFile(BytesIO(excel_bytes))
    
    df_comp = pd.read_excel(xls, sheet_name=main_sheet_name).fillna("-")
    df_plan = pd.read_excel(xls, sheet_name=plan_extraction_sheet) if plan_extraction_sheet in xls.sheet_names else pd.DataFrame()
    df_table = pd.read_excel(xls, sheet_name=table_extraction_sheet) if table_extraction_sheet in xls.sheet_names else pd.DataFrame()

    print(f"Opening Marked PDF file: {pdf_bytes}")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    plan_id_col = resolve_column(df_plan, None)
    plan_page_col = find_page_column(df_plan)
    
    table_id_col = resolve_column(df_table, table_name_column)
    table_page_col = find_page_column(df_table)
    
    comp_id_col = resolve_column(df_comp, None)

    referenced_pages = set()
    if not df_plan.empty and plan_page_col:
        for val in df_plan[plan_page_col]:
            for p in re.findall(r'\d+', str(val)):
                referenced_pages.add(int(p))
                
    if not df_table.empty and table_page_col:
        for val in df_table[table_page_col]:
            for p in re.findall(r'\d+', str(val)):
                referenced_pages.add(int(p))
    
    for col in df_comp.columns:
        if "PAGE" in col.upper():
            for val in df_comp[col]:
                for p in re.findall(r'\d+', str(val)):
                    referenced_pages.add(int(p))

    print(f"Rendering high-res pages at {render_dpi} DPI: {sorted(list(referenced_pages))}")
    page_data_map = {}
    for p_num in referenced_pages:
        if 1 <= p_num <= len(doc):
            page = doc[p_num - 1]
            pix = page.get_pixmap(dpi=render_dpi)
            encoded_img = base64.b64encode(pix.tobytes("png")).decode('utf-8')
            page_data_map[p_num] = {
                "data_uri": f"data:image/png;base64,{encoded_img}",
                "width": pix.width,
                "height": pix.height
            }

    plan_records = []
    if not df_plan.empty:
        plan_records_list = df_plan.to_dict('records')
        for r in plan_records_list:
            s_name = str(r.get(plan_id_col, "")).strip() if plan_id_col else ""
            
            if not s_name or s_name == "-" or s_name.upper() == "NAN":
                continue 
                
            p_val = str(r.get(plan_page_col, "1")).strip() if plan_page_col else "1"
            match_p = re.search(r'\d+', p_val)
            p_num = int(match_p.group()) if match_p else 1
            
            bbox = [0, 0, 0, 0]
            if 1 <= p_num <= len(doc) and s_name:
                page = doc[p_num - 1]
                search_clip = get_page_search_clip(page, p_num, page_config)
                rects = page.search_for(s_name, clip=search_clip)
                
                if not rects and "-" in s_name:
                    rects = page.search_for(s_name.replace("-", " "), clip=search_clip)
                if not rects and " " in s_name:
                    rects = page.search_for(s_name.replace(" ", ""), clip=search_clip)

                if rects:
                    r0 = rects[0]
                    bbox = [(r0.x0 - 6) * SCALE, (r0.y0 - 6) * SCALE, (r0.x1 + 6) * SCALE, (r0.y1 + 6) * SCALE]
            
            plan_records.append({"Structure_Name": s_name, "Page": p_num, "BBox_Pts": bbox})

    # --- SPEED-OPTIMIZED TABLE CROPS RECONSTRUCTION ---
    table_records = []
    table_crop_map = {}

    if not df_table.empty and table_page_col:
        table_pages_to_process = set()
        for val in df_table[table_page_col]:
            for p in re.findall(r'\d+', str(val)):
                table_pages_to_process.add(int(p))

        margin_left_in, margin_right_in, margin_top_in, margin_bottom_in = 1.5, 2.5, 0.5, 0.5
        row_tol, col_tol, skip_headers = 5, col_tol_input, 3

        for p_num in sorted(table_pages_to_process):
            if 1 <= p_num <= len(doc):
                page = doc[p_num - 1]
                crop_rect = fitz.Rect(
                    margin_left_in * 72.0, 
                    margin_top_in * 72.0, 
                    page.rect.width - (margin_right_in * 72.0), 
                    page.rect.height - (margin_bottom_in * 72.0)
                )
                words = page.get_text("words", clip=crop_rect)
                df_grid, row_bboxes = group_ocr_into_table(
                    words, row_tolerance=row_tol, col_tolerance=col_tol, skip_header_rows=skip_headers
                )

                if df_grid.empty:
                    continue

                # Render page once at render_dpi and convert to PIL Image once per table page
                # This completely eliminates heavy per-row page.get_pixmap() calls
                pix_page = page.get_pixmap(dpi=render_dpi)
                pil_page = Image.frombytes("RGB", [pix_page.width, pix_page.height], pix_page.samples)

                target_col = table_name_column if table_name_column in df_grid.columns else "Col_1"
                grid_records = df_grid.to_dict('records')
                
                for idx, row in enumerate(grid_records):
                    s_name = str(row.get(target_col, "")).strip()
                    if not s_name or s_name.upper() in ["-", "NAN"]:
                        continue

                    r0 = row_bboxes[idx]
                    crop_x0 = 18
                    crop_y0 = max(0, r0[1] - 18)
                    clip_x1 = page.rect.width - 18
                    clip_y1 = min(page.rect.height, r0[3] + 18)

                    px_x0 = int(crop_x0 * SCALE)
                    px_y0 = int(crop_y0 * SCALE)
                    px_x1 = int(clip_x1 * SCALE)
                    px_y1 = int(clip_y1 * SCALE)

                    # Lightning-fast PIL crop from the pre-rendered page image
                    pil_crop = pil_page.crop((px_x0, px_y0, px_x1, px_y1))
                    
                    buffered = BytesIO()
                    pil_crop.save(buffered, format="PNG")
                    crop_b64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
                    crop_uri = f"data:image/png;base64,{crop_b64}"
                    
                    t_bbox = [
                        (r0[0] - crop_x0 - 4) * SCALE,
                        (r0[1] - crop_y0 - 4) * SCALE,
                        (r0[2] - crop_x0 + 4) * SCALE,
                        (r0[3] - crop_y0 + 4) * SCALE
                    ]

                    t_rec = {"Structure_Name": s_name, "Page": p_num, "Crop_URI": crop_uri, "Table_BBox": t_bbox}
                    table_records.append(t_rec)
                    table_crop_map[s_name.upper()] = t_rec
                    table_crop_map[re.sub(r'[^A-Z0-9]', '', s_name.upper())] = t_rec

    dashboard_items = []
    plan_dict_by_name = {}
    for pr in plan_records:
        k = pr["Structure_Name"].upper()
        if k not in plan_dict_by_name:
            plan_dict_by_name[k] = []
        plan_dict_by_name[k].append(pr)

    table_columns = [str(col) for col in df_comp.columns]

    comp_records_list = df_comp.to_dict('records')
    for comp in comp_records_list:
        item_data = comp.copy()
        s_name = str(comp.get(comp_id_col, "")).strip() if comp_id_col else ""
        norm_name = s_name.upper()
        clean_key = re.sub(r'[^A-Z0-9]', '', norm_name)
        
        item_data["_structure_name"] = s_name
        item_data["_status"] = str(comp.get("Status", "MATCH")).upper()
        
        if not norm_name or norm_name == "-" or norm_name == "NAN":
            p_items = []
            t_item = None
        else:
            p_items = plan_dict_by_name.get(norm_name, [])
            t_item = table_crop_map.get(norm_name) or table_crop_map.get(clean_key)

        if p_items:
            for p_item in p_items:
                row_copy = item_data.copy()
                row_copy["_plan_page"] = p_item["Page"]
                row_copy["_plan_bbox"] = p_item["BBox_Pts"]
                row_copy["_table_page"] = t_item["Page"] if t_item else "-"
                row_copy["_table_crop_uri"] = t_item["Crop_URI"] if t_item else ""
                row_copy["_table_bbox"] = t_item["Table_BBox"] if t_item else [0, 0, 0, 0]
                dashboard_items.append(row_copy)
        else:
            row_copy = item_data.copy()
            row_copy["_plan_page"] = "-"
            row_copy["_plan_bbox"] = [0, 0, 0, 0]
            row_copy["_table_page"] = t_item["Page"] if t_item else "-"
            row_copy["_table_crop_uri"] = t_item["Crop_URI"] if t_item else ""
            row_copy["_table_bbox"] = t_item["Table_BBox"] if t_item else [0, 0, 0, 0]
            dashboard_items.append(row_copy)

    final_html = HTML_TEMPLATE.replace(
        "/*__DASHBOARD_TITLE__*/", dashboard_title
    ).replace(
        "/*__DASHBOARD_DATA__*/", json.dumps(dashboard_items)
    ).replace(
        "/*__TABLE_COLUMNS__*/", json.dumps(table_columns)
    ).replace(
        "/*__PAGE_META__*/", json.dumps({str(k): v for k, v in page_data_map.items()})
    )

    # with open(output_html, "w", encoding="utf-8") as f:
    #     f.write(final_html)
    
    return final_html

    print(f"\n==================================================")
    print(f"SUCCESS! Created '{dashboard_title}'")
    #print(f"Output File: {output_html}")
    print(f"Universal Headers mapped: {table_columns}")
    print(f"Active Table Name Column: {table_name_column}")
    print(f"==================================================")