import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="KLD Excel → SVG Generator", layout="wide")
st.title("📏 KLD Excel → SVG Generator (Final Illustrator 8 pt Version)")
st.caption("Reads KLD Excel, extracts sequences, and generates editable SVG dieline for Illustrator QC.")

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

    # --- Crop Marks (all corners, full directional set) ---
    out.append('<g id="CropMarks">')

    # --- TOP-LEFT ---
    # Horizontal (→ left)
    out.append(f'<line x1="0" y1="{H}" x2="{-crop_off - crop_len}" y2="{H}" class="dieline"/>')
    # Vertical (↑ up)
    out.append(f'<line x1="0" y1="{H}" x2="0" y2="{H + crop_off + crop_len}" class="dieline"/>')

    # --- TOP-RIGHT ---
    # Horizontal (→ right)
    out.append(f'<line x1="{W + crop_off}" y1="{H}" x2="{W + crop_off + crop_len}" y2="{H}" class="dieline"/>')
    # Vertical (↑ up)
    out.append(f'<line x1="{W}" y1="{H}" x2="{W}" y2="{H + crop_off + crop_len}" class="dieline"/>')

    # --- BOTTOM-LEFT ---
    # Horizontal (→ left)
    out.append(f'<line x1="0" y1="0" x2="{-crop_off - crop_len}" y2="0" class="dieline"/>')
    # Vertical (↓ down)
    out.append(f'<line x1="0" y1="0" x2="0" y2="{-crop_off - crop_len}" class="dieline"/>')

    # --- BOTTOM-RIGHT ---
    # Horizontal (→ right)
    out.append(f'<line x1="{W + crop_off}" y1="0" x2="{W + crop_off + crop_len}" y2="0" class="dieline"/>')
    # Vertical (↓ down)
    out.append(f'<line x1="{W}" y1="0" x2="{W}" y2="{-crop_off - crop_len}" class="dieline"/>')

    out.append('</g>')


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
