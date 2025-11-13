import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="KLD Excel → SVG Generator", layout="wide")
st.title("📏 KLD Excel → SVG Generator (with Dynamic Boxes)")
st.caption("Reads KLD Excel, extracts dimensions and sequences, and generates editable SVG dieline with dynamic boxes for Illustrator QC.")

# ---------------------------------------------------
# Helper functions
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


def extract_dimensions(lines):
    """Extract main width x cut_length dimensions."""
    for ln in lines:
        if re.search(r"dimension|width|cut", ln, re.IGNORECASE):
            w, c = first_pair_from_text(ln)
            if w and c:
                return w, c
    return 0, 0


# ---------------------------------------------------
# Data extraction
# ---------------------------------------------------

def extract_kld_data(df):
    df = df.fillna("").astype(str)
    df = df[df.apply(lambda r: any(str(x).strip() for x in r), axis=1)].reset_index(drop=True)

    header_lines = []
    start_row = 0

    # find header and numeric table start
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
    }


# ---------------------------------------------------
# SVG generator
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

    # --- Style ---
    dieline = "#92278f"
    stroke_pt = 0.356
    font_mm = 8 / 2.8346
    tick_short = 5.0
    top_shift_up = 5.0
    left_shift_left = 5.0
    crop_off = 5.0
    crop_len = 5.0

    out = []
    out.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}mm" height="{H}mm" viewBox="0 0 {W} {H}">')
    out.append('<defs><style><![CDATA[')
    out.append(f'.dieline{{stroke:{dieline};stroke-width:{stroke_pt}pt;fill:none;}}')
    out.append(f'.text{{font-family:Arial; font-size:{font_mm}mm; fill:{dieline};}}')
    out.append(']]></style></defs>')

    # --- Outer dieline box ---
    out.append(f'<rect x="0" y="0" width="{W}" height="{H}" class="dieline"/>')

    # --- Dynamic Rectangles based on sequences ---
    out.append('<g id="DynamicBoxes">')
    max_top = max(top_seq) if top_seq else 0
    x_pos = 0

    skip_pattern = [2, 5, 8, 11, 14, 17]
    for skip in skip_pattern:
        if skip >= len(side_seq):
            break
        top_y = sum(side_seq[:skip])
        height = sum(side_seq[skip:])
        if height <= 0:
            continue
        out.append(f'<rect x="{x_pos}" y="{top_y}" width="{max_top}" height="{height}" class="dieline"/>')
    out.append('</g>')

    out.append('</svg>')
    return "\n".join(out)


# ---------------------------------------------------
# Streamlit execution
# ---------------------------------------------------

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])

if uploaded_file:
    data = uploaded_file.read()
    uploaded_file.seek(0)
    df = pd.read_excel(io.BytesIO(data), header=None, engine="openpyxl")

    try:
        res = extract_kld_data(df)
        svg_data = make_svg(res)

        st.success("✅ Processed successfully.")
        st.download_button("⬇️ Download SVG File", svg_data, f"{uploaded_file.name}_layout.svg", "image/svg+xml")
        st.code(f"side_seq: {res['side_seq']}", language="text")
    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")
else:
    st.info("Please upload a KLD Excel file to begin.")
