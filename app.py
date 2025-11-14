# app.py
import streamlit as st
import pandas as pd
import re
import io
from openpyxl import load_workbook

st.set_page_config(page_title="KLD Excel → SVG Generator (Grey-detect, Seals)", layout="wide")
st.title("📏 KLD Excel → SVG Generator (Grey-detect + Seals)")
st.caption("Detects grey header region (any filled cell != white), extracts header until numeric table, and generates SVG dieline.")

# ------------------------
# Helpers
# ------------------------
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
        return int(float(m.group(1))), int(float(m.group(2)))
    return 0, 0

def auto_trim_to_target(values, target, tol=1.0):
    vals = values.copy()
    while len(vals) > 1 and target > 0 and sum(vals) > target + tol:
        vals.pop()
    return vals

# ------------------------
# Grey detection helpers (Option A: any filled color != white)
# ------------------------
def cell_is_filled(cell):
    """
    Treat any non-white fill as part of the header region.
    Returns True for pattern fills or fgColor.rgb that is not white.
    """
    try:
        fl = getattr(cell, "fill", None)
        if not fl:
            return False
        patt = getattr(fl, "patternType", None)
        # If pattern type exists and not 'none' it's filled
        if patt and str(patt).lower() != "none":
            return True
        # fallback: check fgColor rgb if available and not white
        fg = getattr(fl, "fgColor", None)
        if fg:
            rgb = getattr(fg, "rgb", None)
            if rgb:
                rgb = str(rgb).upper()
                # treat common white values as not-filled; anything else -> filled
                if rgb not in ("FFFFFFFF", "00FFFFFF", "00000000", "FFFFFF"):
                    return True
        return False
    except Exception:
        return False

# ------------------------
# Excel parsing + sequence extraction
# ------------------------
def extract_kld_data_from_bytes(xl_bytes):
    """
    Input: bytes of xlsx file
    Returns dict with job_name, header_lines (list), dimension_text, width_mm, cut_length_mm, top_seq, side_seq (CSV strings)
    """
    bytes_io = io.BytesIO(xl_bytes)

    # load with openpyxl to inspect fills
    try:
        wb = load_workbook(bytes_io, data_only=True, read_only=False)
        ws = wb.active
    except Exception as e:
        raise RuntimeError(f"openpyxl failed to load workbook: {e}")

    # find filled cells (Option A)
    filled_positions = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if cell_is_filled(cell):
                filled_positions.append((r, c))

    # bounding rectangle of filled cells
    if filled_positions:
        rows = [r for r, _ in filled_positions]
        cols = [c for _, c in filled_positions]
        r_min, r_max = min(rows), max(rows)
        c_min, c_max = min(cols), max(cols)
    else:
        # fallback region B2:AB68
        r_min, r_max, c_min, c_max = 2, 68, 2, 28
        r_max = min(r_max, ws.max_row)
        c_max = min(c_max, ws.max_column)

    # collect header lines from bounding rectangle ROW-WISE
    # Stop when the first numeric row (>=3 numeric cells) is encountered
    header_lines = []
    detected_dim_line = ""
    detected_width = 0
    detected_cut = 0

    for r in range(r_min, r_max + 1):
        # compute numeric count in this row (within bounding cols)
        numeric_count = 0
        row_texts = []
        for c in range(c_min, c_max + 1):
            val = ws.cell(row=r, column=c).value
            if val is None:
                continue
            sval = str(val).strip()
            if sval == "":
                continue
            # count numeric-looking cells
            if re.match(r"^-?\d+(\.\d+)?$", sval):
                numeric_count += 1
            row_texts.append(sval)
        # If numeric row (>=3 numeric values), stop header collection
        if numeric_count >= 3:
            # try to detect dimension pair in this row (if any)
            joined = " ".join(row_texts)
            wtmp,ctmp = first_pair_from_text(joined)
            if wtmp and ctmp:
                detected_width = wtmp
                detected_cut = ctmp
            break
        # else treat row as header (if it contains any text)
        joined_text = " ".join(row_texts).strip()
        if joined_text:
            header_lines.append(joined_text)
            # detect dimension line among header rows too
            if not detected_dim_line and re.search(r"dimension|width|cut", joined_text, re.IGNORECASE):
                detected_dim_line = joined_text

    # fallback for dimension: if detected_dim_line empty, check detected_width/detected_cut
    if detected_dim_line:
        w_val, c_val = first_pair_from_text(detected_dim_line)
    elif detected_width and detected_cut:
        w_val, c_val = detected_width, detected_cut
    else:
        # look through header_lines for any pair
        w_val, c_val = (0, 0)
        for ln in header_lines:
            wtmp, ctmp = first_pair_from_text(ln)
            if wtmp and ctmp:
                w_val, c_val = wtmp, ctmp
                break

    width_mm = int(w_val) if w_val else 0
    cut_length_mm = int(c_val) if c_val else 0
    dimension_text = f"Dimension ( Width * Cut-off length ) : {width_mm} * {cut_length_mm} ( in mm )"

    # derive job_name: first alphabetic header line
    job_name = next((ln for ln in header_lines if re.search(r"[A-Za-z]", ln)), "KLD Layout")

    # =================================================
    # Now load into pandas to detect numeric table & sequences
    # =================================================
    bytes_io.seek(0)
    try:
        df = pd.read_excel(bytes_io, header=None, engine="openpyxl")
    except Exception as e:
        raise RuntimeError(f"pandas.read_excel failed: {e}")

    # clean df
    df = df.fillna("").astype(str)
    df = df[df.apply(lambda r: any(str(x).strip() for x in r), axis=1)].reset_index(drop=True)

    # find numeric table start row heuristically (global scan)
    start_row = 0
    for i in range(min(120, len(df))):
        row = df.iloc[i].tolist()
        numeric_count = sum(1 for c in row if re.match(r"^\d+(\.\d+)?$", str(c).strip()))
        if numeric_count >= 3:
            start_row = i
            break

    df_num = df.iloc[start_row:].reset_index(drop=True)

    # top_seq: pick best row (row with numbers matching cut_length_mm)
    top_seq_nums = []
    best_diff = float("inf")
    for i in range(len(df_num)):
        nums = clean_numeric_list(df_num.iloc[i].tolist())
        if len(nums) >= 4:
            diff = abs(sum(nums) - cut_length_mm) if cut_length_mm else sum(nums)
            if diff < best_diff:
                best_diff, top_seq_nums = diff, nums

    # side_seq: find contiguous column subsequence that sums to width_mm
    side_seq_nums = []
    best_diff = float("inf")
    Wtarget = width_mm
    for c in df_num.columns:
        nums = clean_numeric_list(df_num[c].tolist())
        n = len(nums)
        if n < 3:
            continue
        for i in range(n):
            s = 0.0
            for j in range(i, n):
                s += nums[j]
                if j - i + 1 >= 3:
                    diff = abs(s - Wtarget) if Wtarget else s
                    if diff < best_diff:
                        best_diff = diff
                        side_seq_nums = nums[i:j+1]

    top_seq_trimmed = auto_trim_to_target(top_seq_nums, cut_length_mm)
    side_seq_trimmed = auto_trim_to_target(side_seq_nums, width_mm)

    top_seq_str = ",".join(str(int(v)) if float(v).is_integer() else str(v) for v in top_seq_trimmed)
    side_seq_str = ",".join(str(int(v)) if float(v).is_integer() else str(v) for v in side_seq_trimmed)

    return {
        "job_name": job_name,
        "header_lines": header_lines,
        "dimension_text": dimension_text,
        "width_mm": width_mm,
        "cut_length_mm": cut_length_mm,
        "top_seq": top_seq_str,
        "side_seq": side_seq_str
    }

# ------------------------
# SVG generator (complete)
# ------------------------
def make_svg(data, line_spacing_mm=5.0):
    def parse_seq(src):
        if not src:
            return []
        if isinstance(src, (list, tuple)):
            return [float(x) for x in src if x]
        parts = re.split(r"[,;|]", str(src))
        return [float(p.strip()) for p in parts if re.match(r"^\d+(\.\d+)?$", p.strip())]

    # Read inputs
    W = float(data.get("cut_length_mm") or 0)
    H = float(data.get("width_mm") or 0)
    top_seq = parse_seq(data.get("top_seq"))
    side_seq = parse_seq(data.get("side_seq"))
    job_name = data.get("job_name", "KLD Layout")
    header_lines = data.get("header_lines", [])
    dimension_text = data.get("dimension_text", f"Dimension ( Width * Cut-off length ) : {int(W)} * {int(H)} ( in mm )")

    # Artboard expansion
    extra = 60.0
    margin = extra / 2.0
    canvas_W = W + extra
    canvas_H = H + extra

    # Styles
    dieline = "#92278f"
    stroke_pt = 0.356
    font_mm = 1.5           # approx 8pt visual in Illustrator
    tick_short = 5.0
    top_shift_up = 5.0
    left_shift_left = 5.0
    left_text_shift_right = 6.0
    top_text_shift_down = 4.0
    crop_off = 5.0
    crop_len = 5.0

    out = []
    out.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{canvas_W}mm" height="{canvas_H}mm" viewBox="0 0 {canvas_W} {canvas_H}">')
    out.append('<defs><style><![CDATA[')
    out.append(f'.dieline{{stroke:{dieline};stroke-width:{stroke_pt}pt;fill:none;}}')
    out.append(f'.text{{font-family:Arial; font-size:{font_mm}mm; fill:{dieline};}}')
    out.append(']]></style></defs>')

    # Header block: use header_lines extracted (all lines above numeric table)
    header_y = 0.0
    center_x = canvas_W / 2.0
    out.append('<g id="HeaderBlock">')
    # if header_lines empty, fallback to job_name + dimension_text
    if not header_lines:
        header_lines = [job_name, dimension_text]
    for i, line in enumerate(header_lines):
        y = header_y + (i + 1) * line_spacing_mm
        out.append(f'<text x="{center_x}" y="{y}" text-anchor="middle" class="text">{line}</text>')
    out.append('</g>')

    # Outer dieline rectangle (shifted by margin)
    out.append(f'<rect x="{margin}" y="{margin}" width="{W}" height="{H}" class="dieline"/>')

    # Measurements group
    out.append('<g id="Measurements">')

    # TOP ticks & labels
    x = 0.0
    out.append(f'<line x1="{margin}" y1="{margin - top_shift_up}" x2="{margin}" y2="{margin - top_shift_up - tick_short}" class="dieline"/>')
    for v in top_seq:
        x += v
        out.append(f'<line x1="{margin + x}" y1="{margin - top_shift_up}" x2="{margin + x}" y2="{margin - top_shift_up - tick_short}" class="dieline"/>')
        mid = x - v / 2.0
        out.append(f'<text x="{margin + mid}" y="{margin - top_shift_up - tick_short - 1 + top_text_shift_down}" text-anchor="middle" class="text">{int(v)}</text>')

    # LEFT ticks & labels
    y = 0.0
    out.append(f'<line x1="{margin - left_shift_left}" y1="{margin}" x2="{margin - left_shift_left - tick_short}" y2="{margin}" class="dieline"/>')
    for v in side_seq:
        y += v
        out.append(f'<line x1="{margin - left_shift_left}" y1="{margin + y}" x2="{margin - left_shift_left - tick_short}" y2="{margin + y}" class="dieline"/>')
        midY = y - v / 2.0
        lx = margin - left_shift_left - tick_short - 2 + left_text_shift_right
        out.append(f'<text x="{lx}" y="{margin + midY}" transform="rotate(-90 {lx} {margin + midY})" text-anchor="middle" class="text">{int(v)}</text>')
    out.append('</g>')

    # Crop marks (right side outward marks)
    out.append('<g id="CropMarks">')
    out.append(f'<line x1="{margin + W + crop_off}" y1="{margin + H}" x2="{margin + W + crop_off + crop_len}" y2="{margin + H}" class="dieline"/>')
    out.append(f'<line x1="{margin + W + crop_off}" y1="{margin}" x2="{margin + W + crop_off + crop_len}" y2="{margin}" class="dieline"/>')
    out.append('</g>')

    # Photocell mark (top-right) with diagonal and label
    photocell_w, photocell_h = 6, 12
    pc_x = margin + W - photocell_w
    pc_y = margin
    out.append('<g id="PhotocellMark">')
    out.append(f'<rect x="{pc_x}" y="{pc_y}" width="{photocell_w}" height="{photocell_h}" class="dieline"/>')
    out.append(f'<line x1="{pc_x + photocell_w}" y1="{pc_y}" x2="{pc_x + photocell_w + 3}" y2="{pc_y - 3}" class="dieline"/>')
    out.append(f'<text x="{pc_x + photocell_w + 2}" y="{pc_y - 4}" class="text">Photocell Mark {photocell_w}×{photocell_h} mm</text>')
    out.append('</g>')

    # Width indicator (vertical) and label
    total_width = sum(side_seq) if side_seq else 0
    width_line_x = pc_x + photocell_w + 4
    midY = margin + total_width / 2.0 if total_width else margin
    out.append('<g id="WidthMarker">')
    out.append(f'<line x1="{width_line_x}" y1="{margin}" x2="{width_line_x}" y2="{margin + total_width}" class="dieline"/>')
    out.append(f'<text x="{width_line_x + 6}" y="{midY}" transform="rotate(-90 {width_line_x + 6} {midY})" text-anchor="middle" class="text">width = {int(total_width)} mm</text>')
    out.append('</g>')

    # Height indicator (horizontal) and label
    total_height = sum(top_seq) if top_seq else 0
    height_y = margin + total_width + 5
    out.append('<g id="HeightMarker">')
    out.append(f'<line x1="{margin}" y1="{height_y}" x2="{margin + total_height}" y2="{height_y}" class="dieline"/>')
    out.append(f'<text x="{margin + total_height/2.0}" y="{height_y + 6}" text-anchor="middle" class="text">height = {int(total_height)} mm</text>')
    out.append('</g>')

    # Dynamic boxes (width = max(top_seq))
    out.append('<g id="DynamicBoxes">')
    if top_seq and side_seq:
        max_top = max(top_seq)
        # index of first max_top
        try:
            max_idx = next((i for i, v in enumerate(top_seq) if v == max_top), 0)
        except StopIteration:
            max_idx = 0
        left_x = margin + sum(top_seq[:max_idx]) if max_idx > 0 else margin
        skip = 2
        while skip < len(side_seq):
            top_y = margin + sum(side_seq[:skip])
            if skip < len(side_seq):
                height = side_seq[skip]
                out.append(f'<rect x="{left_x}" y="{top_y}" width="{max_top}" height="{height}" class="dieline"/>')
            skip += 3
    out.append('</g>')

    # Seal labels (End Seal left/right, Center Seal top/bottom)
    out.append('<g id="Seals">')
    total_side = sum(side_seq) if side_seq else 0
    mid_side = margin + total_side / 2.0

    # Left END SEAL - horizontally centered to 1st top_seq index
    first_top = top_seq[0] if top_seq else 0
    left_end_x = margin + first_top / 2.0
    out.append(f'<text x="{left_end_x}" y="{mid_side}" transform="rotate(-90 {left_end_x} {mid_side})" text-anchor="middle" class="text">END SEAL</text>')

    # Right END SEAL - horizontally centered to last top_seq index
    last_top = top_seq[-1] if top_seq else 0
    right_end_x = margin + W - last_top / 2.0
    out.append(f'<text x="{right_end_x}" y="{mid_side}" transform="rotate(90 {right_end_x} {mid_side})" text-anchor="middle" class="text">END SEAL</text>')

    # Centre Seal - top (horiz centered to sum(top_seq), vertically center to first side_seq)
    total_top = sum(top_seq) if top_seq else 0
    mid_top_x = margin + total_top / 2.0
    first_side = side_seq[0] if side_seq else 0
    top_center_y = margin + first_side / 2.0
    out.append(f'<text x="{mid_top_x}" y="{top_center_y}" text-anchor="middle" class="text">CENTER SEAL</text>')

    # Centre Seal - bottom (horiz centered same as top, vertically center to last side_seq)
    last_side = side_seq[-1] if side_seq else 0
    bottom_center_y = margin + total_side - last_side / 2.0
    out.append(f'<text x="{mid_top_x}" y="{bottom_center_y}" text-anchor="middle" class="text">CENTER SEAL</text>')
    out.append('</g>')

    # End SVG
    out.append('</svg>')
    return "\n".join(out)

# ------------------------
# Streamlit app
# ------------------------
uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        raw = uploaded_file.read()
        data = extract_kld_data_from_bytes(raw)

        # parse sequences for validation
        def parse_seq_list(src):
            if not src:
                return []
            parts = re.split(r"[,;|]", str(src))
            return [float(p.strip()) for p in parts if re.match(r"^\d+(\.\d+)?$", p.strip())]

        top_seq = parse_seq_list(data.get("top_seq"))
        side_seq = parse_seq_list(data.get("side_seq"))
        cut_length_mm = data.get("cut_length_mm") or 0
        width_mm = data.get("width_mm") or 0

        sum_top = sum(top_seq) if top_seq else 0
        sum_side = sum(side_seq) if side_seq else 0

        # validation warnings (Option A: show warning but still allow download)
        if cut_length_mm and sum_top and sum_top != float(cut_length_mm):
            st.warning(f"⚠️ Sum of top_seq ({sum_top}) does NOT match cut_length_mm ({cut_length_mm}). Please verify the Excel.")
        if width_mm and sum_side and sum_side != float(width_mm):
            st.warning(f"⚠️ Sum of side_seq ({sum_side}) does NOT match width_mm ({width_mm}). Please verify the Excel.")

        svg = make_svg(data, line_spacing_mm=5.0)
        st.success("✅ Processed successfully.")
        st.download_button("⬇️ Download SVG File", svg, f"{uploaded_file.name}_layout.svg", "image/svg+xml")
        st.code(f"side_seq: {data['side_seq']}\ntop_seq: {data['top_seq']}", language="text")
    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")
else:
    st.info("Please upload the Excel (.xlsx) file (not CSV) to use grey-detection.")
