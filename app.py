import streamlit as st
import pandas as pd
import re
import io
from openpyxl import load_workbook

# ===========================================================
# Streamlit Config
# ===========================================================
st.set_page_config(page_title="KLD Excel → SVG Generator (Spot Dieline)", layout="wide")
st.title("📏 KLD Excel → SVG Generator (Spot-Colour Dieline)")
st.caption("Extracts header via grey bounding box (Option A), validates dimensions, and outputs SVG with SPOT COLOUR ‘Dieline’ (C50 M100 Y0 K0).")

# ===========================================================
# Helpers
# ===========================================================
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

# ===========================================================
# Grey Cell Detector
# ===========================================================
def cell_is_filled(cell):
    try:
        fl = cell.fill
        if not fl:
            return False

        if fl.patternType and str(fl.patternType).lower() != "none":
            return True

        fg = fl.fgColor
        if fg:
            rgb = getattr(fg, "rgb", None)
            if rgb:
                rgb = str(rgb).upper()
                if rgb not in ("FFFFFFFF", "FFFFFF", "00FFFFFF"):
                    return True
            if fg.indexed is not None or fg.theme is not None:
                return True
        return False
    except:
        return False

# ===========================================================
# Extractor – Option A (grey bounding box)
# ===========================================================
def extract_kld_data_from_bytes(xl_bytes):
    bytes_io = io.BytesIO(xl_bytes)
    wb = load_workbook(bytes_io, data_only=True, read_only=False)
    ws = wb.active

    filled = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if cell_is_filled(ws.cell(r, c)):
                filled.append((r, c))

    if filled:
        rows = [r for r, _ in filled]
        cols = [c for _, c in filled]
        r_min, r_max = min(rows), max(rows)
        c_min, c_max = min(cols), max(cols)
    else:
        r_min, r_max, c_min, c_max = 2, 68, 2, 28

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

            if re.match(r"^-?\d+(\.\d+)?$", sval):
                numeric_count += 1
            else:
                row_texts.append(sval)

        if numeric_count >= 3:
            break

        if row_texts:
            line_text = " ".join(row_texts).strip()
            header_lines.append(line_text)
            if not detected_dim_line and re.search(r"(dimension|width|cut)", line_text, re.IGNORECASE):
                detected_dim_line = line_text

    if detected_dim_line:
        w_val, c_val = first_pair_from_text(detected_dim_line)
    else:
        w_val, c_val = 0, 0

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

    bytes_io.seek(0)
    df = pd.read_excel(bytes_io, header=None, engine="openpyxl")
    df = df.fillna("").astype(str)
    df = df[df.apply(lambda r: any(str(x).strip() for x in r), axis=1)].reset_index(drop=True)

    start_row = 0
    for i in range(min(120, len(df))):
        numeric_count = sum(1 for v in df.iloc[i] if re.match(r"^\d+(\.\d+)?$", v.strip()))
        if numeric_count >= 3:
            start_row = i
            break

    df_num = df.iloc[start_row:].reset_index(drop=True)

    top_seq_nums = []
    best_diff = 999999
    for i in range(len(df_num)):
        nums = clean_numeric_list(df_num.iloc[i].tolist())
        if len(nums) >= 4:
            diff = abs(sum(nums) - cut_length_mm)
            if diff < best_diff:
                best_diff = diff
                top_seq_nums = nums

    side_seq_nums = []
    best_diff = 999999
    for col in df_num.columns:
        nums = clean_numeric_list(df_num[col])
        for i in range(len(nums)):
            s = 0
            for j in range(i, len(nums)):
                s += nums[j]
                if j - i + 1 >= 3:
                    diff = abs(s - width_mm)
                    if diff < best_diff:
                        best_diff = diff
                        side_seq_nums = nums[i:j+1]

    top_seq = auto_trim_to_target(top_seq_nums, cut_length_mm)
    side_seq = auto_trim_to_target(side_seq_nums, width_mm)

    return {
        "job_name": job_name,
        "header_lines": header_lines,
        "dimension_text": dimension_text,
        "width_mm": width_mm,
        "cut_length_mm": cut_length_mm,
        "top_seq": ",".join(map(str, top_seq)),
        "side_seq": ",".join(map(str, side_seq)),
    }

# ===========================================================
# SVG Generator with SPOT COLOUR “Dieline”
# ===========================================================
def make_svg(data):
    def parse_seq(txt):
        return [float(x) for x in re.split(r"[ ,;]+", txt) if re.match(r"^\d+(\.\d+)?$", x)]

    W = float(data["cut_length_mm"])
    H = float(data["width_mm"])
    top = parse_seq(data["top_seq"])
    side = parse_seq(data["side_seq"])
    headers = data["header_lines"]

    # SPOT COLOUR DEFINITION – Illustrator compatible
    # Name: Dieline
    # Model: spot
    # Tint Build: C50 M100 Y0 K0
    spot_colour = "cmyk(0.5,1,0,0)"   # Illustrator interprets CMYK 0-1 scale

    out = []
    out.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{W+60}mm" height="{H+60}mm">')

    # Define spot colour swatch
    out.append(f'''
<defs>
    <swatch id="Dieline" inkscape:swatch="solid">
        <color profile="CMYK" model="spot" name="Dieline" c="0.5" m="1" y="0" k="0"/>
    </swatch>
    <style>
        .dieline {{ stroke: {spot_colour}; fill:none; stroke-width:0.4mm; }}
        .text {{ fill:{spot_colour}; font-family:Arial; font-size:4mm; }}
    </style>
</defs>
''')

    cx = (W+60)/2
    for i,line in enumerate(headers):
        out.append(f'<text class="text" x="{cx}" y="{10+i*6}" text-anchor="middle">{line}</text>')

    margin = 30
    out.append(f'<rect class="dieline" x="{margin}" y="{margin}" width="{W}" height="{H}"/>')

    return "\n".join(out)

# ===========================================================
# Streamlit App
# ===========================================================
uploaded = st.file_uploader("Upload KLD Excel", type=["xlsx"])

if uploaded:
    raw = uploaded.read()
    data = extract_kld_data_from_bytes(raw)

    top = list(map(float, data["top_seq"].split(","))) if data["top_seq"] else []
    side = list(map(float, data["side_seq"].split(","))) if data["side_seq"] else []

    if sum(top) != data["cut_length_mm"]:
        st.error("Top sequence mismatch.")
        st.stop()

    if sum(side) != data["width_mm"]:
        st.error("Side sequence mismatch.")
        st.stop()

    svg = make_svg(data)

    st.success("SVG ready (Spot Colour: Dieline)")
    st.download_button("Download SVG", svg, "layout.svg", "image/svg+xml")
    st.code(svg[:800], language="xml")
else:
    st.info("Upload Excel to begin.")
