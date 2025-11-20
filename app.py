
import streamlit as st
import pandas as pd
import re
import io
from openpyxl import load_workbook

# Sample uploaded file path (available in this session)
# You can use this path to test locally if needed:
# /mnt/data/Tentative KLD Doodh Marie 142.5g- 207 count.xlsx

st.set_page_config(page_title="KLD Excel → SVG Generator (Grey-detect, Seals + Strict Validation)", layout="wide")
st.title("📏 KLD Excel → SVG Generator (Grey-detect + Seals + Strict Validation)")
st.caption("Detects grey header region (bounding box), extracts header until numeric table, applies strict dimension validation, and generates SVG dieline.")

# ===========================================
# Helpers
# ===========================================
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

# ===========================================
# Grey detection helpers
# ===========================================
def cell_is_filled(cell):
    """
    Consider a cell filled if it has a non-empty patternType or an fgColor
    that's not white (several variants checked).
    """
    try:
        fl = getattr(cell, "fill", None)
        if not fl:
            return False

        patt = getattr(fl, "patternType", None)
        if patt and str(patt).lower() != "none":
            return True

        fg = getattr(fl, "fgColor", None)
        if fg:
            # prefer rgb if present
            rgb = getattr(fg, "rgb", None)
            if rgb:
                rgb = str(rgb).upper()
                # common "white" representations
                if rgb not in ("FFFFFFFF", "00FFFFFF", "00000000", "FFFFFF"):
                    return True
            # fallback: indexed/theme may indicate fill
            # treat indexed or theme values as filled (conservative)
            if getattr(fg, "indexed", None) is not None or getattr(fg, "theme", None) is not None:
                return True
        return False
    except Exception:
        return False

# ===========================================
# Extraction (Option A: bounding-box logic)
# ===========================================
def extract_kld_data_from_bytes(xl_bytes):
    bytes_io = io.BytesIO(xl_bytes)
    wb = load_workbook(bytes_io, data_only=True, read_only=False)
    ws = wb.active

    # Find all filled cells and compute bounding box
    filled_positions = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            try:
                if cell_is_filled(ws.cell(r, c)):
                    filled_positions.append((r, c))
            except Exception:
                # ignore errors when probing cell fill
                continue

    if filled_positions:
        rows = [r for r, _ in filled_positions]
        cols = [c for _, c in filled_positions]
        r_min, r_max = min(rows), max(rows)
        c_min, c_max = min(cols), max(cols)
    else:
        # sensible defaults if no grey found
        r_min, r_max, c_min, c_max = 2, min(68, ws.max_row), 2, min(28, ws.max_column)

    # Collect header lines: any non-numeric text inside bounding box (Option A)
    header_lines = []
    detected_dim_line = ""

    for r in range(r_min, r_max + 1):
        numeric_count = 0
        row_texts = []

        for c in range(c_min, c_max + 1):
            val = ws.cell(r, c).value
            if val is None:
                continue
            sval = str(val).strip()
            if sval == "":
                continue

            # count numeric-like cells
            if re.match(r"^-?\d+(\.\d+)?$", sval):
                numeric_count += 1
            else:
                # non-numeric text inside bounding box counts as header text (Option A)
                row_texts.append(sval)

        # If this row has >=3 numeric cells, assume numeric table begins -> stop header extraction
        if numeric_count >= 3:
            break

        if row_texts:
            line_text = " ".join(row_texts).strip()
            header_lines.append(line_text)
            if not detected_dim_line and re.search(r"dimension|width|cut", line_text, re.IGNORECASE):
                detected_dim_line = line_text

    # Dimension parsing: prefer detected_dim_line, else scan header_lines for first pair
    if detected_dim_line:
        w_val, c_val = first_pair_from_text(detected_dim_line)
    else:
        w_val, c_val = 0, 0

    # If not found from detected_dim_line, scan header lines
    if not (w_val and c_val):
        for ln in header_lines:
            t1, t2 = first_pair_from_text(ln)
            if t1 and t2:
                w_val, c_val = t1, t2
                break

    width_mm = int(w_val) if w_val else 0
    cut_length_mm = int(c_val) if c_val else 0

    dimension_text = f"Dimension ( Width * Cut-off length ) : {width_mm} * {cut_length_mm} ( in mm )"
    job_name = next((ln for ln in header_lines if re.search(r"[A-Za-z]", ln)), "KLD Layout")

    # Load numeric table through pandas to extract sequences
    bytes_io.seek(0)
    df = pd.read_excel(bytes_io, header=None, engine="openpyxl")
    df = df.fillna("").astype(str)
    df = df[df.apply(lambda r: any(str(x).strip() for x in r), axis=1)].reset_index(drop=True)

    # find first row that looks numeric (>=3 numeric cells)
    start_row = 0
    for i in range(min(120, len(df))):
        numeric_count = sum(1 for c in df.iloc[i].tolist() if re.match(r"^\d+(\.\d+)?$", str(c).strip()))
        if numeric_count >= 3:
            start_row = i
            break
    df_num = df.iloc[start_row:].reset_index(drop=True)

    # Extract top_seq: find a row with >=4 numeric fields whose sum ~ cut_length_mm
    top_seq_nums = []
    best_diff = float("inf")
    for i in range(len(df_num)):
        nums = clean_numeric_list(df_num.iloc[i].tolist())
        if len(nums) >= 4:
            diff = abs(sum(nums) - cut_length_mm)
            if diff < best_diff:
                best_diff = diff
                top_seq_nums = nums

    # Extract side_seq: search columns for contiguous runs of numbers whose sum ~ width_mm
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

# ===========================================
# SVG generator (full)
# ===========================================
def make_svg(data, line_spacing_mm=5.0):
    def parse_seq(src):
        if not src:
            return []
        parts = re.split(r"[,;|]", str(src))
        out = []
        for p in parts:
            p = p.strip()
            if re.match(r"^\d+(\.\d+)?$", p):
                out.append(float(p))
        return out

    W = float(data["cut_length_mm"]) if data["cut_length_mm"] else 0.0
    H = float(data["width_mm"]) if data["width_mm"] else 0.0
    top_seq = parse_seq(data["top_seq"])
    side_seq = parse_seq(data["side_seq"])
    header_lines = data.get("header_lines", [])
    if not header_lines:
        header_lines = [data.get("job_name", "KLD Layout"), data.get("dimension_text", "")]

    # Basic canvas
    extra = 60
    margin = extra / 2
    canvas_W = W + extra
    canvas_H = H + extra

    # Spot colour definition (CMYK, name: Dieline)
    spot_color_def = """
    <defs>
          <icc-color-profile name="Dieline" xlink:href="data:application/vnd.iccprofile;base64,">
          </icc-color-profile>
    </defs>
    """

    dieline = "icc-color(Dieline, 0.5, 1, 0, 0)"
    stroke_pt = 0.356
    font_mm = 1.5
    tick_short = 5
    top_shift_up = 5
    left_shift_left = 5
    left_text_shift_right = 6
    top_text_shift_down = 4
    crop_off = 5
    crop_len = 5

    out = []
    out.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{canvas_W}mm" height="{canvas_H}mm" viewBox="0 0 {canvas_W} {canvas_H}">')
    out.append('<defs><style><![CDATA[')
    out.append(f'.dieline{{stroke:icc-color(Dieline, 0.5, 1, 0, 0);stroke-width:{stroke_pt}pt;fill:none;}}')
    out.append(f'.text{{font-family:Arial; font-size:{font_mm}mm; fill:icc-color(Dieline, 0.5, 1, 0, 0);}}')
    out.append(']]></style></defs>')

    # Header
    out.append('<g id="Header">')
    cx = canvas_W / 2
    for i, line in enumerate(header_lines):
        y = (i + 1) * line_spacing_mm
        out.append(f'<text x="{cx}" y="{y}" text-anchor="middle" class="text">{line}</text>')
    out.append('</g>')

    # Dieline
    out.append(f'<rect x="{margin}" y="{margin}" width="{W}" height="{H}" class="dieline"/>')

    # Measurements group (top)
    out.append('<g id="Measurements">')
    x = 0
    out.append(f'<line x1="{margin}" y1="{margin-top_shift_up}" x2="{margin}" y2="{margin-top_shift_up-tick_short}" class="dieline"/>')
    for v in top_seq:
        x += v
        out.append(f'<line x1="{margin+x}" y1="{margin-top_shift_up}" x2="{margin+x}" y2="{margin-top_shift_up-tick_short}" class="dieline"/>')
        mid = x - v / 2
        out.append(f'<text x="{margin+mid}" y="{margin-top_shift_up-tick_short-1+top_text_shift_down}" text-anchor="middle" class="text">{round(v, 2)}</text>')

    # Left measurements
    y = 0
    out.append(f'<line x1="{margin-left_shift_left}" y1="{margin}" x2="{margin-left_shift_left-tick_short}" y2="{margin}" class="dieline"/>')
    for v in side_seq:
        y += v
        out.append(f'<line x1="{margin-left_shift_left}" y1="{margin+y}" x2="{margin-left_shift_left-tick_short}" y2="{margin+y}" class="dieline"/>')
        midY = y - v / 2
        lx = margin - left_shift_left - tick_short - 2 + left_text_shift_right
        out.append(f'<text x="{lx}" y="{margin+midY}" transform="rotate(-90 {lx} {margin+midY})" text-anchor="middle" class="text">{int(v)}</text>')
    out.append('</g>')

    # Crop marks
    out.append('<g id="CropMarks">')
    out.append(f'<line x1="{margin+W+crop_off}" y1="{margin}" x2="{margin+W+crop_off+crop_len}" y2="{margin}" class="dieline"/>')
    out.append(f'<line x1="{margin+W+crop_off}" y1="{margin+H}" x2="{margin+W+crop_off+crop_len}" y2="{margin+H}" class="dieline"/>')
    out.append('</g>')

    # Photocell
    photocell_w, photocell_h = 6, 12
    pc_x = margin + W - photocell_w
    pc_y = margin
    out.append('<g id="Photocell">')
    out.append(f'<rect x="{pc_x}" y="{pc_y}" width="{photocell_w}" height="{photocell_h}" class="dieline"/>')
    out.append(f'<line x1="{pc_x+photocell_w}" y1="{pc_y}" x2="{pc_x+photocell_w+3}" y2="{pc_y-3}" class="dieline"/>')
    out.append(f'<text x="{pc_x+photocell_w+2}" y="{pc_y-4}" class="text">Photocell Mark {photocell_w}×{photocell_h} mm</text>')
    out.append('</g>')

    # Width marker
    total_width = sum(side_seq) if side_seq else 0
    wx = pc_x + photocell_w + 4
    midY = margin + total_width / 2
    out.append('<g id="WidthMarker">')
    out.append(f'<line x1="{wx}" y1="{margin}" x2="{wx}" y2="{margin+total_width}" class="dieline"/>')
    out.append(f'<text x="{wx+6}" y="{midY}" transform="rotate(-90 {wx+6} {midY})" class="text" text-anchor="middle">width = {int(total_width)} mm</text>')
    out.append('</g>')

    # Height marker
    total_height = sum(top_seq) if top_seq else 0
    hy = margin + total_width + 5
    out.append('<g id="HeightMarker">')
    out.append(f'<line x1="{margin}" y1="{hy}" x2="{margin+total_height}" y2="{hy}" class="dieline"/>')
    out.append(f'<text x="{margin+total_height/2}" y="{hy+6}" text-anchor="middle" class="text">height = {int(total_height)} mm</text>')
    out.append('</g>')

    # Dynamic boxes
    out.append('<g id="DynamicBoxes">')
    if top_seq and side_seq:
        max_top = max(top_seq)
        max_idx = top_seq.index(max_top)
        left_x = margin + sum(top_seq[:max_idx])
        skip = 2
        while skip < len(side_seq):
            top_y = margin + sum(side_seq[:skip])
            height = side_seq[skip]
            out.append(f'<rect x="{left_x}" y="{top_y}" width="{max_top}" height="{height}" class="dieline"/>')
            skip += 3
    out.append('</g>')

    # Seals
    out.append('<g id="Seals">')
    total_side = sum(side_seq) if side_seq else 0
    mid_side = margin + total_side / 2
    first_top = top_seq[0] if top_seq else 0
    last_top = top_seq[-1] if top_seq else 0
    left_end_x = margin + first_top / 2 if first_top else margin + 10
    out.append(f'<text x="{left_end_x}" y="{mid_side}" transform="rotate(-90 {left_end_x} {mid_side})" text-anchor="middle" class="text">END SEAL</text>')
    right_end_x = margin + W - last_top / 2 if last_top else margin + W - 10
    out.append(f'<text x="{right_end_x}" y="{mid_side}" transform="rotate(90 {right_end_x} {mid_side})" text-anchor="middle" class="text">END SEAL</text>')

    total_top = sum(top_seq) if top_seq else 0
    mid_top_x = margin + total_top / 2
    first_side = side_seq[0] if side_seq else 0
    last_side = side_seq[-1] if side_seq else 0
    top_center_y = margin + first_side / 2 if first_side else margin + 10
    bottom_center_y = margin + total_side - last_side / 2 if last_side else margin + total_side - 10
    out.append(f'<text x="{mid_top_x}" y="{top_center_y}" text-anchor="middle" class="text">CENTER SEAL</text>')
    out.append(f'<text x="{mid_top_x}" y="{bottom_center_y}" text-anchor="middle" class="text">CENTER SEAL</text>')
    out.append('</g>')

    out.append('</svg>')
    return "\\n".join(out)

# ===========================================
# Streamlit App (using strict validation)
# ===========================================
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
        if cut_length_mm and sum_top and abs(sum_top - float(cut_length_mm)) > 0.0001:
            errors.append(f"Sum of top_seq = {sum_top} mm does NOT match cut_length_mm = {cut_length_mm} mm.")

        if width_mm and sum_side and abs(sum_side - float(width_mm)) > 0.0001:
            errors.append(f"Sum of side_seq = {sum_side} mm does NOT match width_mm = {width_mm} mm.")

        if errors:
            st.error("❌ SVG generation blocked due to mismatched dimensions:")
            for e in errors:
                st.write(f"- {e}")
            st.stop()

        svg = make_svg(data)

        st.success("✅ Dimensions validated. No mismatches detected.")
        st.download_button("⬇️ Download SVG File", svg, f"{uploaded_file.name}_layout.svg", "image/svg+xml")
        st.code(f"side_seq: {data['side_seq']}\ntop_seq: {data['top_seq']}\nheader_lines: {data['header_lines']}", language="text")

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")

else:
    st.info("Please upload the Excel (.xlsx) file to begin.")
