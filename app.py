import streamlit as st
import pandas as pd
import re
import io
from openpyxl import load_workbook

st.set_page_config(page_title="KLD Excel → SVG Generator (Grey-detect, Seals)", layout="wide")
st.title("📏 KLD Excel → SVG Generator (Grey-detect + Seals + Strict Validation)")
st.caption("Detects grey header region, extracts header until numeric table, applies strict dimension validation, and generates SVG dieline.")

# ==========================================================
# Helpers
# ==========================================================
def clean_numeric_list(seq):
    out = []
    for v in seq:
        s = str(v).strip().replace(",", "")
        if not s or s.lower() in ("nan", "none"):
            continue
        try:
            out.append(float(s))
        except:
            m = re.search(r"(-?\d+(?:\.\d+)?)", s)
            if m:
                out.append(float(m.group(1)))
    return out


def first_pair_from_text(text):
    text = str(text)
    m = re.search(r"(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)", text)
    if m:
        return float(m.group(1)), float(m.group(2))
    return 0, 0


def auto_trim_to_target(values, target, tol=1.0):
    vals = values.copy()
    while len(vals) > 1 and target > 0 and sum(vals) > target + tol:
        vals.pop()
    return vals


# ==========================================================
# Grey Detection
# ==========================================================
def cell_is_filled(cell):
    """Any non-white excel fill is part of grey header."""
    try:
        fl = getattr(cell, "fill", None)
        if not fl:
            return False

        patt = getattr(fl, "patternType", None)
        if patt and str(patt).lower() != "none":
            return True

        fg = getattr(fl, "fgColor", None)
        if fg:
            rgb = getattr(fg, "rgb", None)
            if rgb:
                rgb = str(rgb).upper()
                if rgb not in ("FFFFFFFF", "00FFFFFF", "00000000", "FFFFFF"):
                    return True
        return False
    except:
        return False


# ==========================================================
# Extraction
# ==========================================================
def extract_kld_data_from_bytes(xl_bytes):

    bytes_io = io.BytesIO(xl_bytes)
    wb = load_workbook(bytes_io, data_only=True, read_only=False)
    ws = wb.active

    # Detect filled positions
    filled_positions = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if cell_is_filled(ws.cell(r, c)):
                filled_positions.append((r, c))

    if filled_positions:
        rows = [r for r, _ in filled_positions]
        cols = [c for _, c in filled_positions]
        r_min, r_max = min(rows), max(rows)
        c_min, c_max = min(cols), max(cols)
    else:
        r_min, r_max, c_min, c_max = 2, 68, 2, 28
    

    # ======================================================
    # STRICT HEADER EXTRACTION (critical fix)
    # Only include cells that are actually grey-filled
    # ======================================================
    header_lines = []
    detected_dim_line = ""

    for r in range(r_min, r_max + 1):

        # Skip entire row if no grey cells
        if not any(cell_is_filled(ws.cell(r, cc)) for cc in range(c_min, c_max + 1)):
            continue

        row_vals = []
        numeric_count = 0

        for c in range(c_min, c_max + 1):
            cell = ws.cell(r, c)

            # NEW: must be grey-filled
            if not cell_is_filled(cell):
                continue

            val = cell.value
            if val is None:
                continue

            sval = str(val).strip()
            if sval == "":
                continue

            if re.match(r"^-?\d+(\.\d+)?$", sval):
                numeric_count += 1

            row_vals.append(sval)

        if numeric_count >= 3:
            break

        line_text = " ".join(row_vals).strip()
        if line_text:
            header_lines.append(line_text)

        if not detected_dim_line and re.search(r"dimension|width|cut", line_text, re.IGNORECASE):
            detected_dim_line = line_text


    # ======================================================
    # Parse Dimensions
    # ======================================================
    if detected_dim_line:
        w_val, c_val = first_pair_from_text(detected_dim_line)
    else:
        w_val, c_val = (0, 0)

    for ln in header_lines:
        t1, t2 = first_pair_from_text(ln)
        if t1 and t2:
            w_val, t2 = t1, t2
            break

    width_mm = int(w_val) if w_val else 0
    cut_length_mm = int(c_val) if c_val else 0

    dimension_text = f"Dimension ( Width * Cut-off length ) : {width_mm} * {cut_length_mm} ( in mm )"

    job_name = next((ln for ln in header_lines if re.search(r"[A-Za-z]", ln)), "KLD Layout")


    # ======================================================
    # Load numeric table via pandas
    # ======================================================
    bytes_io.seek(0)
    df = pd.read_excel(bytes_io, header=None, engine="openpyxl")
    df = df.fillna("").astype(str)
    df = df[df.apply(lambda r: any(str(x).strip() for x in r), axis=1)]
    df = df.reset_index(drop=True)

    start_row = 0
    for i in range(min(120, len(df))):
        numeric_count = sum(1 for c in df.iloc[i].tolist() if re.match(r"^\d+(\.\d+)?$", c.strip()))
        if numeric_count >= 3:
            start_row = i
            break

    df_num = df.iloc[start_row:].reset_index(drop=True)


    # ======================================================
    # Extract sequences
    # ======================================================
    top_seq_nums = []
    best_diff = float("inf")

    for i in range(len(df_num)):
        nums = clean_numeric_list(df_num.iloc[i].tolist())
        if len(nums) >= 4:
            diff = abs(sum(nums) - cut_length_mm)
            if diff < best_diff:
                best_diff = diff
                top_seq_nums = nums

    side_seq_nums = []
    best_diff = float("inf")

    for c in df_num.columns:
        nums = clean_numeric_list(df_num[c].tolist())
        for i in range(len(nums)):
            s = 0
            for j in range(i, len(nums)):
                s += nums[j]
                if j - i + 1 >= 3:
                    diff = abs(s - width_mm)
                    if diff < best_diff:
                        best_diff = diff
                        side_seq_nums = nums[i:j+1]

    top_seq_trimmed = auto_trim_to_target(top_seq_nums, cut_length_mm)
    side_seq_trimmed = auto_trim_to_target(side_seq_nums, width_mm)

    top_seq_str = ",".join(str(v).rstrip("0").rstrip(".") for v in top_seq_trimmed)
    side_seq_str = ",".join(str(v).rstrip("0").rstrip(".") for v in side_seq_trimmed)

    return {
        "job_name": job_name,
        "header_lines": header_lines,
        "dimension_text": dimension_text,
        "width_mm": width_mm,
        "cut_length_mm": cut_length_mm,
        "top_seq": top_seq_str,
        "side_seq": side_seq_str,
    }


# ==========================================================
# SVG Generator (unchanged)
# ==========================================================
def make_svg(data, line_spacing_mm=5.0):
    # (SVG code unchanged — omitted here because user already confirmed it works)
    return "<svg>...</svg>"


# ==========================================================
# Streamlit App
# ==========================================================
uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        raw = uploaded_file.read()
        data = extract_kld_data_from_bytes(raw)

        def parse_seq_list(src):
            parts = re.split(r"[,;|]", str(src))
            return [float(p.strip()) for p in parts if re.match(r"^\d+(\.\d+)?$", p.strip())]

        top_seq = parse_seq_list(data["top_seq"])
        side_seq = parse_seq_list(data["side_seq"])

        sum_top = sum(top_seq)
        sum_side = sum(side_seq)

        cut_length_mm = data["cut_length_mm"]
        width_mm = data["width_mm"]

        errors = []
        if cut_length_mm and sum_top and sum_top != float(cut_length_mm):
            errors.append(f"Sum of top_seq = {sum_top} mm does NOT match cut_length_mm = {cut_length_mm} mm.")

        if width_mm and sum_side and sum_side != float(width_mm):
            errors.append(f"Sum of side_seq = {sum_side} mm does NOT match width_mm = {width_mm} mm.")

        if errors:
            st.error("❌ SVG generation blocked due to mismatched dimensions:")
            for e in errors:
                st.write(f"- {e}")
            st.stop()

        svg = make_svg(data)

        st.success("✅ Dimensions validated. No mismatches detected.")
        st.download_button("⬇️ Download SVG File", svg, f"{uploaded_file.name}_layout.svg", "image/svg+xml")
        st.code(f"side_seq: {data['side_seq']}\ntop_seq: {data['top_seq']}", language="text")

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")

else:
    st.info("Please upload the Excel (.xlsx) file to begin.")
