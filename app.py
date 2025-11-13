import streamlit as st
import pandas as pd
import re
import io
import html

st.set_page_config(page_title="KLD Excel → SVG Generator", layout="wide")
st.title("📏 KLD Excel → SVG Generator (Final v7)")
st.caption("Reads Excel, extracts dimensions, sequences & generates an editable SVG dieline for Illustrator QC.")

# ---------------------------------------------------
# Helper Functions
# ---------------------------------------------------

def clean_numeric_list(seq):
    """Extract numeric floats from list cells."""
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
    """Extract numeric pair like 15*20, 15 x 20."""
    text = str(text)
    m = re.search(r"(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)", text)
    if m:
        return int(float(m.group(1))), int(float(m.group(2)))
    return 0, 0


def auto_trim_to_target(values, target, tol=1.0):
    """Trim numeric sequence to sum near target ± tol."""
    vals = values.copy()
    while len(vals) > 1 and target > 0 and sum(vals) > target + tol:
        vals.pop()
    return vals


def extract_print_areas(lines):
    """Extract Print Area (left/right) and Printing Area (main)."""
    finarea_left = finarea_right = printarea = ""
    for ln in lines:
        if re.search(r"print\s*area", ln, re.IGNORECASE):
            pairs = re.findall(r"(\d+)\s*[*xX]\s*(\d+)", ln)
            if len(pairs) == 1:
                finarea_left = f"{pairs[0][0]}x{pairs[0][1]} mm"
            elif len(pairs) >= 2:
                finarea_left = f"{pairs[0][0]}x{pairs[0][1]} mm"
                finarea_right = f"{pairs[1][0]}x{pairs[1][1]} mm"

        if re.search(r"printing\s*area", ln, re.IGNORECASE):
            w, h = first_pair_from_text(ln)
            if w and h:
                printarea = f"{w}x{h} mm"
    return finarea_left, finarea_right, printarea


def extract_dimensions(lines):
    """Extract main width x cut_length dimensions."""
    for ln in lines:
        if re.search(r"dimension|width|cut", ln, re.IGNORECASE):
            w, c = first_pair_from_text(ln)
            if w and c:
                return w, c
    return 0, 0


# ---------------------------------------------------
# Data Extraction
# ---------------------------------------------------

def extract_kld_data(df):
    df = df.fillna("").astype(str)
    df = df[df.apply(lambda r: any(str(x).strip() for x in r), axis=1)].reset_index(drop=True)

    header_lines = []
    start_row = 0

    for i in range(min(60, len(df))):
        row = df.iloc[i].tolist()
        line_text = " ".join([s.strip() for s in row if s.strip()])
        numeric_count = sum(1 for c in row if re.match(r"^\d+(\.\d+)?$", str(c)))
        if numeric_count >= 3:
            start_row = i
            break
        if line_text:
            header_lines.append(line_text)

    job_name = "\n".join(header_lines) if header_lines else "Unknown"

    search_lines = [
        " ".join([str(x).strip() for x in df.iloc[i].tolist() if str(x).strip()])
        for i in range(max(0, start_row - 10), min(len(df), start_row + 80))
    ]

    width_mm, cut_length_mm = extract_dimensions(search_lines)
    finarea_left, finarea_right, printarea = extract_print_areas(search_lines)

    df_num = df.iloc[start_row:].reset_index(drop=True)
    top_seq_nums, side_seq_nums = [], []

    # --- Find top sequence (rows) — matches cut_length
    best_diff = float("inf")
    for i in range(len(df_num)):
        nums = clean_numeric_list(df_num.iloc[i].tolist())
        if len(nums) >= 4:
            diff = abs(sum(nums) - cut_length_mm) if cut_length_mm else sum(nums)
            if diff < best_diff:
                best_diff, top_seq_nums = diff, nums

    # --- Find side sequence (columns) — contiguous subsequence that sums to width ---
    best_diff = float("inf")
    W = width_mm
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
                    diff = abs(s - W) if W else s
                    if diff < best_diff:
                        best_diff = diff
                        side_seq_nums = nums[i:j+1]

    top_seq_trimmed = auto_trim_to_target(top_seq_nums, cut_length_mm)
    side_seq_trimmed = auto_trim_to_target(side_seq_nums, width_mm)

    top_seq = ",".join(str(int(v)) if v.is_integer() else str(v) for v in top_seq_trimmed)
    side_seq = ",".join(str(int(v)) if v.is_integer() else str(v) for v in side_seq_trimmed)

    return {
        "job_name": job_name,
        "width_mm": width_mm,
        "cut_length_mm": cut_length_mm,
        "top_seq": top_seq,
        "side_seq": side_seq,
        "finarea_left": finarea_left,
        "finarea_right": finarea_right,
        "printarea": printarea,
        "photocell_w": 6,
        "photocell_h": 12,
        "photocell_offset_right_mm": 12,
        "stroke_mm": 0.25,
        "brand_label": "BRANDING",
    }


# ---------------------------------------------------
# SVG Generator (your final spec)
# ---------------------------------------------------

def make_svg(data):
    def parse_seq(src):
        if src is None:
            return []
        if isinstance(src, (list, tuple)):
            return [float(x) for x in src if x]
        s = str(src).replace(";", ",").replace("|", ",")
        parts = [p.strip() for p in s.split(",") if p.strip()]
        return [float(p) for p in parts if re.match(r"^\d+(\.\d+)?$", p)]

    W = float(data.get("cut_length_mm") or 0)
    H = float(data.get("width_mm") or 0)
    top_seq = parse_seq(data.get("top_seq"))
    side_seq = parse_seq(data.get("side_seq"))

    dieline = "#92278f"
    stroke_pt = 0.356
    font_pt = 8         # pt for correct Illustrator scale
    tick_short = 5.0
    top_shift_up = 5.0
    left_shift_left = 5.0
    crop_off = 5.0
    crop_len = 5.0

out = []
out.append(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}">')
out.append('<defs><style><![CDATA[')
out.append(f'.dieline{{stroke:{dieline};stroke-width:{stroke_pt}pt;fill:none;}}')
out.append(f'.text{{font-family:Arial; font-size:8pt; fill:{dieline};}}')
out.append(']]></style></defs>')

# --- Outer dieline box ---
out.append(f'<rect x="0" y="0" width="{W}" height="{H}" class="dieline"/>')

    # --- Measurement ticks and labels ---
    out.append('<g id="Measurements">')
    # TOP ticks and labels
    x = 0
    out.append(f'<line x1="0" y1="{-top_shift_up}" x2="0" y2="{-top_shift_up - tick_short}" class="dieline"/>')
    for v in top_seq:
        x += v
        out.append(f'<line x1="{x}" y1="{-top_shift_up}" x2="{x}" y2="{-top_shift_up - tick_short}" class="dieline"/>')
        mid = x - v / 2
        out.append(f'<text x="{mid}" y="{-top_shift_up - tick_short - 1}" text-anchor="middle" class="text">{int(v)}</text>')

    # LEFT ticks and labels
    y = 0
    out.append(f'<line x1="{-left_shift_left}" y1="0" x2="{-left_shift_left - tick_short}" y2="0" class="dieline"/>')
    for v in side_seq:
        y += v
        out.append(f'<line x1="{-left_shift_left}" y1="{y}" x2="{-left_shift_left - tick_short}" y2="{y}" class="dieline"/>')
        midY = y - v / 2
        lx = -left_shift_left - tick_short - 2
        out.append(f'<text x="{lx}" y="{midY}" transform="rotate(-90 {lx} {midY})" text-anchor="middle" class="text">{int(v)}</text>')
    out.append('</g>')

    # --- Crop Marks ---
    out.append('<g id="CropMarks">')
    # Right horizontal (outward, +X)
    out.append(f'<line x1="{W + crop_off}" y1="{H}" x2="{W + crop_off + crop_len}" y2="{H}" class="dieline"/>')
    # Bottom left vertical (outward, +Y)
    out.append(f'<line x1="0" y1="{0 - crop_off}" x2="0" y2="{-crop_off - crop_len}" class="dieline"/>')
    # Bottom right vertical (outward, +Y)
    out.append(f'<line x1="{W}" y1="{0 - crop_off}" x2="{W}" y2="{-crop_off - crop_len}" class="dieline"/>')
    # Bottom right horizontal (outward, +X)
    out.append(f'<line x1="{W + crop_off}" y1="0" x2="{W + crop_off + crop_len}" y2="0" class="dieline"/>')
    out.append('</g>')

    out.append('</svg>')
    return "\n".join(out)

# ---------------------------------------------------
# Streamlit Execution
# ---------------------------------------------------

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])
if uploaded_file:
    data = uploaded_file.read()
    uploaded_file.seek(0)
    df = pd.read_excel(io.BytesIO(data), header=None, engine="openpyxl")

    res = extract_kld_data(df)
    svg_data = make_svg(res)

    st.success("✅ Processed successfully.")
    st.download_button("⬇️ Download SVG File", svg_data, f"{uploaded_file.name}_layout.svg", "image/svg+xml")
    st.code(res["side_seq"], language="text")
    st.caption("Previewed side_seq above — verify matches Excel structure.")
else:
    st.info("Please upload a KLD Excel file to begin.")
