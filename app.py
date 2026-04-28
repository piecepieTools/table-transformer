"""
Table Transformer — Streamlit version
Steps: Upload → Select → Configure (numeric) → Rename (multi_value labels) → Arrange (order) → Download
"""

import streamlit as st
import pandas as pd
import io, os
from collections import Counter
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Table Transformer", page_icon="📊", layout="centered")

st.markdown("""
<style>
  html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
  #MainMenu, footer, header { visibility: hidden; }
  .block-container {
    max-width: 780px !important;
    padding-left: 2rem !important; padding-right: 2rem !important;
    margin-left: auto !important; margin-right: auto !important;
  }
  .step-bar { display:flex; gap:0; margin-bottom:1.5rem; }
  .step-pill {
    flex:1; text-align:center; padding:7px 4px; font-size:11px;
    background:#F2F2F2; color:#AAAAAA; border:1px solid #E0E0E0;
  }
  .step-pill.active { background:#222; color:#FFF; font-weight:600; }
  .step-pill.done   { background:#555; color:#FFF; }
  .section-title { font-size:22px; font-weight:700; color:#111; margin-bottom:4px; }
  .section-sub   { font-size:14px; color:#777; margin-bottom:1.5rem; }
  .stats-strip {
    display:flex; background:#EFEFEF; border-radius:4px;
    padding:12px 0; margin:12px 0 20px; text-align:center;
  }
  .stat-item { flex:1; }
  .stat-val  { font-size:18px; font-weight:700; color:#111; }
  .stat-lbl  { font-size:11px; color:#888; margin-top:2px; }
  .success-box {
    background:#FFF; border:2px solid #222; border-radius:8px;
    padding:48px; text-align:center; margin-top:2rem;
  }
  .success-check { font-size:56px; }
  .success-title { font-size:24px; font-weight:700; color:#111; margin:12px 0 4px; }
  .success-sub   { font-size:14px; color:#888; }
  div.stButton > button {
    background:#222 !important; color:#FFF !important;
    border:none !important; border-radius:4px !important;
    padding:8px 20px !important; font-weight:600 !important;
  }
  div.stButton > button:hover { background:#444 !important; }
  div.stDownloadButton > button {
    background:#222 !important; color:#FFF !important;
    border:none !important; border-radius:4px !important;
    padding:10px 24px !important; font-weight:600 !important; font-size:15px !important;
  }
  .raw-val {
    font-size:13px; color:#555; background:#F5F5F5;
    border-radius:4px; padding:2px 8px; font-family:monospace;
  }
  .cnt-badge { font-size:11px; color:#999; margin-left:8px; }
</style>
""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────
CAT_THRESHOLD = 10
ARIAL = "Arial"
OSYM = {"<": "<", "<=": "≤", ">": ">", ">=": "≥"}
CSYM = {"<": "≥", "<=": ">", ">": "≤", ">=": "<"}

# Steps: 1=Upload 2=Select 3=Configure(numeric) 4=Rename(multi_value) 5=Arrange 6=Download
STEPS = ["Upload", "Select", "Configure", "Rename", "Arrange", "Download"]

# ── Session state ──────────────────────────────────────────────
def _init():
    defaults = {
        "step":      1,
        "df":        None,
        "col_types": {},
        "selected":  [],
        "configs":   {},   # col → {type, label, order, label_map, modes…}
        "cat_order": {},   # col → [val, val, …]  (order only, used in Arrange)
        "num_queue": [],
        "cur_num":   None,
        "num_total": 0,
        "word_bytes": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()
S = st.session_state

# ── Type detection ─────────────────────────────────────────────
def _is_multi_value(series):
    non_null = series.dropna().astype(str)
    if non_null.empty:
        return False
    return non_null.str.contains(",").mean() >= 0.1

def detect_type(series):
    if not pd.api.types.is_numeric_dtype(series):
        if _is_multi_value(series):
            return "multi_value"
        return "categorical"
    return "categorical" if series.nunique(dropna=True) <= CAT_THRESHOLD else "numeric"

# ── Multi-value helpers ────────────────────────────────────────
def split_multi(series):
    terms = []
    for val in series.dropna().astype(str):
        for t in val.split(","):
            t = t.strip()
            if t:
                terms.append(t.lower())
    return terms

def multi_value_counts(series):
    return Counter(split_multi(series))

# ── Formatting ─────────────────────────────────────────────────
def fmt_pct(pct):
    return f"{int(pct)}%" if pct == int(pct) else f"{pct:.1f}%"

def go(step):
    S.step = step
    st.rerun()

# ── Step pill bar ──────────────────────────────────────────────
def render_steps():
    pills = ""
    for i, name in enumerate(STEPS, 1):
        if i < S.step:    cls = "step-pill done"
        elif i == S.step: cls = "step-pill active"
        else:             cls = "step-pill"
        pills += f'<div class="{cls}">{i}. {name}</div>'
    st.markdown(f'<div class="step-bar">{pills}</div>', unsafe_allow_html=True)

# ── Word export ────────────────────────────────────────────────
def _set_arial(run):
    rPr = run._r.get_or_add_rPr()
    rf = OxmlElement("w:rFonts")
    for k in ("w:ascii", "w:hAnsi", "w:cs"): rf.set(qn(k), ARIAL)
    rPr.insert(0, rf)

def _wcell(cell, text, bold=False, indent=False, center=False):
    para = cell.paragraphs[0]; para.clear()
    if center: para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if indent: para.paragraph_format.left_indent = Pt(16)
    run = para.add_run(text)
    run.font.name = ARIAL; run.font.size = Pt(10); run.bold = bold
    _set_arial(run)

def _clr(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for s in tcPr.findall(qn("w:shd")): tcPr.remove(s)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "auto")
    tcPr.append(shd)

def _rh(row, t=400):
    trPr = row._tr.get_or_add_trPr()
    rh = OxmlElement("w:trHeight"); rh.set(qn("w:val"), str(t)); trPr.append(rh)

def _hborder(row):
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcB = tcPr.find(qn("w:tcBorders"))
        if tcB is None: tcB = OxmlElement("w:tcBorders"); tcPr.append(tcB)
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0"); b.set(qn("w:color"), "666666")
        tcB.append(b)

def _drow(tbl, c1, c2="", c3="", bold1=False, ind1=False):
    row = tbl.add_row(); _rh(row)
    for c in row.cells: _clr(c)
    _wcell(row.cells[0], c1, bold=bold1, indent=ind1)
    _wcell(row.cells[1], c2, center=True)
    _wcell(row.cells[2], c3, center=True)

def build_word_bytes(col_order, configs, df):
    tpl = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.docx")
    if os.path.exists(tpl):
        doc = Document(tpl)
        for t in list(doc.tables):     t._element.getparent().remove(t._element)
        for p in list(doc.paragraphs): p._element.getparent().remove(p._element)
    else:
        doc = Document()

    tbl = doc.add_table(rows=1, cols=3)
    try: tbl.style = "List Table 1 Light"
    except: pass

    ws = [4536, 900, 1134]
    grid = OxmlElement("w:tblGrid")
    for w in ws:
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(w)); grid.append(gc)
    tblPr = tbl._tbl.find(qn("w:tblPr")); tblPr.addnext(grid)
    for row in tbl.rows:
        for i, cell in enumerate(row.cells):
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None: tcW = OxmlElement("w:tcW"); tcPr.insert(0, tcW)
            tcW.set(qn("w:w"), str(ws[i])); tcW.set(qn("w:type"), "dxa")

    hdr = tbl.rows[0]; _rh(hdr, 500)
    for c in hdr.cells: _clr(c)
    _wcell(hdr.cells[0], "Variable", bold=True)
    _wcell(hdr.cells[1], "N",        bold=True, center=True)
    _wcell(hdr.cells[2], "%",        bold=True, center=True)
    _hborder(hdr)

    OPS_fn = {"<": lambda s,v: s<v, "<=": lambda s,v: s<=v,
              ">": lambda s,v: s>v, ">=": lambda s,v: s>=v}

    for col in col_order:
        cfg    = configs[col]
        series = df[col]
        total  = len(series.dropna())

        if cfg["type"] == "categorical":
            _drow(tbl, cfg["label"], "", "", bold1=True)
            for val in cfg.get("order", list(series.value_counts(dropna=True).index)):
                cnt = (series == val).sum()
                pct = cnt / total * 100 if total else 0
                _drow(tbl, str(val), str(cnt), fmt_pct(pct), ind1=True)
            n_miss = series.isna().sum()
            if n_miss:
                _drow(tbl, "Missing", str(n_miss), fmt_pct(n_miss/len(series)*100), ind1=True)

        elif cfg["type"] == "multi_value":
            _drow(tbl, cfg["label"], str(total), "", bold1=True)
            counts    = multi_value_counts(series)
            order     = cfg.get("order", sorted(counts.keys(), key=lambda k: -counts[k]))
            label_map = cfg.get("label_map", {})
            for term in order:
                cnt  = counts.get(term, 0)
                pct  = cnt / total * 100 if total else 0
                display = label_map.get(term, term.title())
                _drow(tbl, display, str(cnt), fmt_pct(pct), ind1=True)
            n_miss = series.isna().sum()
            if n_miss:
                _drow(tbl, "Missing", str(n_miss), fmt_pct(n_miss/len(series)*100), ind1=True)

        else:  # numeric
            for mc in cfg["modes"]:
                mode  = mc["mode"]
                label = mc["label"]
                s     = series.dropna()
                if mode == "mean_sd":
                    _drow(tbl, label, f"{s.mean():.2f} ± {s.std():.2f}", "", bold1=True)
                elif mode == "median":
                    q1, q3 = s.quantile(0.25), s.quantile(0.75)
                    _drow(tbl, label, f"{s.median():.2f} [{q1:.2f}–{q3:.2f}]", "", bold1=True)
                elif mode == "threshold":
                    op = mc["op"]; v = mc["threshold"]
                    _drow(tbl, label, "", "", bold1=True)
                    mask1 = OPS_fn[op](series, v)
                    cnt1  = mask1.sum(); pct1 = cnt1/total*100 if total else 0
                    mask2 = ~mask1 & series.notna()
                    cnt2  = mask2.sum(); pct2 = cnt2/total*100 if total else 0
                    _drow(tbl, mc["label_true"],  str(cnt1), fmt_pct(pct1), ind1=True)
                    _drow(tbl, mc["label_false"], str(cnt2), fmt_pct(pct2), ind1=True)
                    n_miss = series.isna().sum()
                    if n_miss:
                        _drow(tbl, "Missing", str(n_miss),
                              fmt_pct(n_miss/len(series)*100), ind1=True)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════
#  STEP 1 — Upload
# ══════════════════════════════════════════════════════════════
def step_upload():
    st.markdown('<p class="section-title">Table Transformer</p>', unsafe_allow_html=True)
    st.markdown('<p class="section-sub">Build a formatted Word summary table from your Excel data.</p>',
                unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload your Excel file", type=["xlsx","xls","xlsm"],
                                label_visibility="collapsed")
    if uploaded:
        try:
            df = pd.read_excel(uploaded)
        except Exception as e:
            st.error(f"Could not read file: {e}"); return
        if df.empty:
            st.warning("The file appears to be empty."); return
        S.df        = df
        S.col_types = {c: detect_type(df[c]) for c in df.columns}
        go(2)

# ══════════════════════════════════════════════════════════════
#  STEP 2 — Select columns
# ══════════════════════════════════════════════════════════════
def step_select():
    df = S.df
    st.markdown('<p class="section-title">Select columns</p>', unsafe_allow_html=True)
    st.markdown('<p class="section-sub">Choose which columns to include in the summary table.</p>',
                unsafe_allow_html=True)

    selected = []
    all_cols = list(df.columns)
    for row_start in range(0, len(all_cols), 3):
        row_cols = all_cols[row_start:row_start+3]
        ui_cols  = st.columns(3)
        for ui_col, col in zip(ui_cols, row_cols):
            with ui_col:
                dtype = S.col_types[col]
                badge = {"multi_value": "multi-value"}.get(dtype, dtype)
                checked = st.checkbox(f"**{col}**", value=True, key=f"sel_{col}")
                st.caption(f"`{badge}` · {df[col].nunique(dropna=True)} unique values")
                if checked:
                    selected.append(col)

    st.markdown("---")
    st.markdown("**Data preview** — first 8 rows")
    st.dataframe(df.head(8), use_container_width=True, height=220)
    st.caption(f"{len(df)} rows · {len(df.columns)} columns")

    st.markdown("---")
    cl, cr = st.columns([1,5])
    with cl:
        if st.button("← Back"): go(1)
    with cr:
        if st.button("Continue →", type="primary"):
            if not selected:
                st.error("Please select at least one column."); return
            S.selected = selected

            # Init configs for categorical / multi_value columns
            for col in selected:
                dtype = S.col_types[col]
                if dtype == "categorical":
                    order = list(df[col].value_counts(dropna=True).index)
                    S.configs[col]   = {"type":"categorical","label":col,"order":order}
                    S.cat_order[col] = list(order)
                elif dtype == "multi_value":
                    counts    = multi_value_counts(df[col])
                    order     = sorted(counts.keys(), key=lambda k: -counts[k])
                    label_map = {t: t.title() for t in order}
                    S.configs[col]   = {"type":"multi_value","label":col,
                                        "order":order,"label_map":label_map}
                    S.cat_order[col] = list(order)

            S.num_queue = [c for c in selected if S.col_types[c] == "numeric"]
            S.num_total = len(S.num_queue)
            if S.num_queue:
                S.cur_num = None
                go(3)
            else:
                go(4)   # skip Configure, go to Rename

# ══════════════════════════════════════════════════════════════
#  STEP 3 — Configure numeric columns
# ══════════════════════════════════════════════════════════════
def step_configure():
    if not S.num_queue:
        go(4); return
    if S.cur_num is None or S.cur_num not in S.num_queue:
        S.cur_num = S.num_queue[0]

    col  = S.cur_num
    df   = S.df
    s    = df[col].dropna()
    done = S.num_total - len(S.num_queue) + 1

    st.markdown(
        f'<p class="section-title">Configure: {col}</p>'
        f'<p class="section-sub">Numeric column {done} of {S.num_total}</p>',
        unsafe_allow_html=True)

    st.markdown(
        f'<div class="stats-strip">'
        f'<div class="stat-item"><div class="stat-val">{s.min():.2f}</div><div class="stat-lbl">Min</div></div>'
        f'<div class="stat-item"><div class="stat-val">{s.max():.2f}</div><div class="stat-lbl">Max</div></div>'
        f'<div class="stat-item"><div class="stat-val">{s.mean():.2f}</div><div class="stat-lbl">Mean</div></div>'
        f'<div class="stat-item"><div class="stat-val">{s.median():.2f}</div><div class="stat-lbl">Median</div></div>'
        f'</div>', unsafe_allow_html=True)

    st.markdown("**Select one or more presentation modes:**")
    use_mean   = st.checkbox("Mean ± SD",          key=f"m_mean_{col}")
    use_median = st.checkbox("Median [IQR]",        key=f"m_med_{col}")
    use_thresh = st.checkbox("Groups by threshold", key=f"m_thr_{col}")

    lbl_mean   = st.text_input("Label for Mean ± SD row:", value=f"{col} (Mean ± SD)",
                               key=f"lbl_mean_{col}", disabled=not use_mean)
    lbl_median = st.text_input("Label for Median row:",    value=f"{col} (Median)",
                               key=f"lbl_med_{col}",  disabled=not use_median)

    lbl_thresh = col; thresh_val = None; op = "<"; lbl1 = ""; lbl2 = ""
    auto1 = auto2 = ""

    if use_thresh:
        st.markdown("---")
        st.markdown("**Threshold settings:**")
        lbl_thresh = st.text_input("Label for threshold row:", value=col, key=f"lbl_thr_{col}")
        thresh_val = st.number_input("Threshold value:", value=float(s.median()),
                                     key=f"thresh_{col}")
        op = st.radio(
            "First group condition:",
            options=["<", "<=", ">", ">="],
            format_func=lambda x: {
                "<":  "Smaller than  ( < )",
                "<=": "Smaller than or equal to  ( ≤ )",
                ">":  "Greater than  ( > )",
                ">=": "Greater than or equal to  ( ≥ )",
            }[x],
            key=f"op_{col}", horizontal=False,
        )

        auto1 = f"{OSYM.get(op,'?')} {thresh_val:g}"
        auto2 = f"{CSYM.get(op,'?')} {thresh_val:g}"

        # Force-write label widgets when op or threshold changes.
        # Writing to st.session_state[key] before the widget renders is the
        # only reliable way to update a keyed widget's displayed value.
        snap = st.session_state.get(f"_snap_{col}")
        if snap is None:
            st.session_state[f"l1_{col}"] = auto1
            st.session_state[f"l2_{col}"] = auto2
            st.session_state[f"_snap_{col}"] = (op, thresh_val, auto1, auto2)
        elif op != snap[0] or thresh_val != snap[1]:
            if st.session_state.get(f"l1_{col}", snap[2]) == snap[2]:
                st.session_state[f"l1_{col}"] = auto1
            if st.session_state.get(f"l2_{col}", snap[3]) == snap[3]:
                st.session_state[f"l2_{col}"] = auto2
            st.session_state[f"_snap_{col}"] = (op, thresh_val, auto1, auto2)

        lbl1 = st.text_input("Group 1 label:", key=f"l1_{col}")
        lbl2 = st.text_input("Group 2 label:", key=f"l2_{col}")

    st.markdown("---")
    cl, cr = st.columns([1,5])
    with cl:
        if st.button("← Back"):
            if col not in S.num_queue:
                S.num_queue.insert(0, col)
            S.cur_num = None
            go(2)
    with cr:
        if st.button("Next →", type="primary", key=f"next_{col}"):
            if not (use_mean or use_median or use_thresh):
                st.error("Please select at least one presentation mode."); return
            modes = []
            if use_mean:
                modes.append({"mode":"mean_sd",  "label": lbl_mean.strip() or f"{col} (Mean ± SD)"})
            if use_median:
                modes.append({"mode":"median",   "label": lbl_median.strip() or f"{col} (Median)"})
            if use_thresh:
                if thresh_val is None:
                    st.error("Enter a threshold value."); return
                modes.append({
                    "mode":        "threshold",
                    "label":       lbl_thresh.strip() or col,
                    "op":          op,
                    "threshold":   float(thresh_val),
                    "label_true":  lbl1.strip() or auto1,
                    "label_false": lbl2.strip() or auto2,
                })
            S.configs[col] = {"type":"numeric","label":col,"modes":modes}
            S.num_queue.pop(0)
            S.cur_num = S.num_queue[0] if S.num_queue else None
            if S.num_queue:
                st.rerun()
            else:
                go(4)

# ══════════════════════════════════════════════════════════════
#  STEP 4 — Rename
#  Only shown if there are multi_value columns.
#  Each multi_value column gets an expander with:
#    raw term  |  text input for display label  |  n · %
#  No reorder buttons here — that is step 5.
#  Saves directly into cfg["label_map"][term], keyed by TERM not position.
# ══════════════════════════════════════════════════════════════
def step_rename():
    df = S.df
    mv_cols = [c for c in S.selected if S.col_types[c] == "multi_value"]

    # Skip this step entirely if no multi_value columns
    if not mv_cols:
        go(5); return

    st.markdown('<p class="section-title">Rename values</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="section-sub">Set a display label for each term. '
        'You will set the order in the next step.</p>',
        unsafe_allow_html=True)

    for col in mv_cols:
        cfg    = S.configs[col]
        counts = multi_value_counts(df[col])
        total  = len(df[col].dropna())
        order  = cfg.get("order", sorted(counts.keys(), key=lambda k: -counts[k]))

        with st.expander(f"🏷️  {col}", expanded=True):
            # Section heading
            new_heading = st.text_input("Section heading in table:", value=cfg["label"],
                                        key=f"rn_heading_{col}")
            cfg["label"] = new_heading

            st.markdown("---")
            # Column headers for the rename table
            hc = st.columns([3, 3, 2])
            hc[0].markdown("<span style='font-size:12px;color:#888'>Raw value in spreadsheet</span>",
                           unsafe_allow_html=True)
            hc[1].markdown("<span style='font-size:12px;color:#888'>Display label in Word table</span>",
                           unsafe_allow_html=True)
            hc[2].markdown("<span style='font-size:12px;color:#888'>Count</span>",
                           unsafe_allow_html=True)

            label_map = cfg.get("label_map", {t: t.title() for t in order})

            for term in order:
                cnt = counts.get(term, 0)
                pct = cnt / total * 100 if total else 0
                r = st.columns([3, 3, 2])
                r[0].markdown(
                    f"<div style='padding-top:8px'><span class='raw-val'>{term}</span></div>",
                    unsafe_allow_html=True)
                # Key is term-based, NOT position-based — no drift possible
                new_lbl = r[1].text_input(
                    "label", value=label_map.get(term, term.title()),
                    key=f"rn_lbl_{col}_{term}", label_visibility="collapsed")
                label_map[term] = new_lbl
                r[2].markdown(
                    f"<div style='padding-top:8px;font-size:13px;color:#555'>"
                    f"n={cnt} <span class='cnt-badge'>({fmt_pct(pct)})</span></div>",
                    unsafe_allow_html=True)

            cfg["label_map"] = label_map

    st.markdown("---")
    cl, cr = st.columns([1,5])
    with cl:
        if st.button("← Back"):
            # Go back to configure if there were numeric cols, else to select
            if any(S.col_types[c] == "numeric" for c in S.selected):
                S.num_queue = [c for c in S.selected if S.col_types[c] == "numeric"]
                S.cur_num   = None
                go(3)
            else:
                go(2)
    with cr:
        if st.button("Continue →", type="primary"):
            go(5)

# ══════════════════════════════════════════════════════════════
#  STEP 5 — Arrange (order only, no rename)
#  Shows every selected column.
#  categorical  → reorder raw values (label = value, no rename)
#  multi_value  → reorder terms, shows the LABEL set in step 4
#  numeric      → shows threshold summary, no reorder needed
# ══════════════════════════════════════════════════════════════
def step_arrange():
    df = S.df
    st.markdown('<p class="section-title">Arrange order</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="section-sub">Set the order rows appear in the table. '
        'Labels were fixed in the previous step.</p>',
        unsafe_allow_html=True)

    for col in S.selected:
        cfg   = S.configs[col]
        dtype = S.col_types[col]
        icon  = {"categorical":"🔤","multi_value":"🏷️","numeric":"🔢"}.get(dtype,"•")

        with st.expander(f"{icon}  {col}", expanded=True):

            if dtype == "categorical":
                order = S.cat_order.get(col,
                        list(df[col].value_counts(dropna=True).index))
                total = len(df[col].dropna())
                for idx, val in enumerate(order):
                    cnt = (df[col] == val).sum()
                    pct = cnt / total * 100 if total else 0
                    r = st.columns([6, 1, 1])
                    r[0].markdown(
                        f"<div style='padding:6px 0'>"
                        f"<b>{idx+1}.</b> {val} "
                        f"<span style='color:#aaa;font-size:12px'>n={cnt} · {fmt_pct(pct)}</span>"
                        f"</div>", unsafe_allow_html=True)
                    if r[1].button("▲", key=f"up_{col}_{idx}", disabled=(idx==0)):
                        order[idx], order[idx-1] = order[idx-1], order[idx]
                        S.cat_order[col] = order; st.rerun()
                    if r[2].button("▼", key=f"dn_{col}_{idx}", disabled=(idx==len(order)-1)):
                        order[idx], order[idx+1] = order[idx+1], order[idx]
                        S.cat_order[col] = order; st.rerun()
                S.cat_order[col] = order
                cfg["order"] = order

            elif dtype == "multi_value":
                counts    = multi_value_counts(df[col])
                total     = len(df[col].dropna())
                order     = S.cat_order.get(col, cfg.get("order", list(counts.keys())))
                label_map = cfg.get("label_map", {t: t.title() for t in order})

                st.caption("Percentages are per patient and can sum >100%.")

                for idx, term in enumerate(order):
                    cnt   = counts.get(term, 0)
                    pct   = cnt / total * 100 if total else 0
                    label = label_map.get(term, term.title())
                    r = st.columns([6, 1, 1])
                    r[0].markdown(
                        f"<div style='padding:6px 0'>"
                        f"<b>{idx+1}.</b> {label} "
                        f"<span style='color:#aaa;font-size:12px'>"
                        f"<span class='raw-val' style='font-size:11px'>{term}</span> "
                        f"n={cnt} · {fmt_pct(pct)}</span>"
                        f"</div>", unsafe_allow_html=True)
                    if r[1].button("▲", key=f"mv_up_{col}_{idx}", disabled=(idx==0)):
                        order[idx], order[idx-1] = order[idx-1], order[idx]
                        S.cat_order[col] = order; cfg["order"] = order; st.rerun()
                    if r[2].button("▼", key=f"mv_dn_{col}_{idx}", disabled=(idx==len(order)-1)):
                        order[idx], order[idx+1] = order[idx+1], order[idx]
                        S.cat_order[col] = order; cfg["order"] = order; st.rerun()

                S.cat_order[col] = order
                cfg["order"] = order

            else:  # numeric — just show summary, nothing to reorder
                mode_names = {"mean_sd":"Mean ± SD","median":"Median [IQR]",
                              "threshold":"Threshold groups"}
                for mc in cfg["modes"]:
                    st.markdown(f"**{mode_names.get(mc['mode'], mc['mode'])}:** {mc['label']}")
                    if mc["mode"] == "threshold":
                        st.caption(
                            f"Group 1: {mc['label_true']}  "
                            f"({OSYM.get(mc['op'],'')} {mc['threshold']:g})  ·  "
                            f"Group 2: {mc['label_false']}  "
                            f"({CSYM.get(mc['op'],'')} {mc['threshold']:g})")

    st.markdown("---")
    cl, cr = st.columns([1,5])
    with cl:
        if st.button("← Back"):
            mv_cols = [c for c in S.selected if S.col_types[c] == "multi_value"]
            if mv_cols:
                go(4)
            elif any(S.col_types[c] == "numeric" for c in S.selected):
                S.num_queue = [c for c in S.selected if S.col_types[c] == "numeric"]
                S.cur_num   = None
                go(3)
            else:
                go(2)
    with cr:
        if st.button("Build Word table →", type="primary"):
            S.word_bytes = build_word_bytes(S.selected, S.configs, df)
            go(6)

# ══════════════════════════════════════════════════════════════
#  STEP 6 — Download
# ══════════════════════════════════════════════════════════════
def step_download():
    if not S.word_bytes:
        S.word_bytes = build_word_bytes(S.selected, S.configs, S.df)

    st.markdown(
        '<div class="success-box">'
        '<div class="success-check">✓</div>'
        '<div class="success-title">Your table is ready</div>'
        '<div class="success-sub">Click the button below to download the Word file.</div>'
        '</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        label="📄  Download Word file",
        data=S.word_bytes,
        file_name="summary_table.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Edit", use_container_width=True):
            S.word_bytes = None; go(5)
    with col2:
        if st.button("🔄  Start over", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

# ══════════════════════════════════════════════════════════════
#  Router
# ══════════════════════════════════════════════════════════════
render_steps()

if   S.step == 1: step_upload()
elif S.step == 2: step_select()
elif S.step == 3: step_configure()
elif S.step == 4: step_rename()
elif S.step == 5: step_arrange()
elif S.step == 6: step_download()
