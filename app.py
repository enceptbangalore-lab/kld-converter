import streamlit as st
import pandas as pd
import re
from io import BytesIO
import numpy as np

# ---------------------------------------------------
# Streamlit App Configuration
# ---------------------------------------------------
st.set_page_config(page_title="KLD Excel → CSV Converter", layout="wide")
st.title("📏 KLD Excel → CSV Converter (Auto Trim by Dimension Sum ±1 mm)")
st.caption("Detects job name, dimensions, sequences, and trims edges so total length matches KLD width/height (±1 mm).")

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])
show_debug = st.checkbox("Show debug info", value=False)

# ---------------------------------------------------
# Helpers
# ---------------------------------------------------
def is_number(x):
    try:
        float(x)
        return True
    except:
        return False

def clean_numeric_list(seq):
    out = []
    for v in seq:
        if is_number(v):
            f = float(v)
            out.append(f)
    return out

def auto_trim_to_target(values, target, tol=1.0):
    """Trim from the end until sum <= target + tol."""
    vals = values.copy()
    while sum(vals) > target + tol and len(vals) > 1:
        vals.pop()  # remove last segment (right or bottom)
    return vals

# ---------------------------------------------------
# Core Extraction Logic
# ---------------------------------------------------
def extract_kld_data(df):
    df = df.fillna("").astype(str)

    # --- 1. Job name ---
    job_name = "Unknown"
    for row in df.values:
        joined = " ".join(row)
        m = re.search(r"Lam(?:inate)?\s*KLD\s*for\s*(.+)", joined, re.IGNORECASE)
        if m:
            job_name = m.group(1).strip()
            break

    # --- 2. Dimensions ---
    width_mm = cut_length_mm = 0
    for row in df.values:
        joined = " ".join(row)
        if "DIMENSION" in joined.upper():
            nums = re.findall(r"\d+", joined)
            if len(nums) >= 2:
                width_mm, cut_length_mm = map(int, nums[:2])
                break

    # --- 3. Pack note ---
    pack_note = ""
    for row in df.values:
        joined = " ".join(row)
        if re.search(r"biscuits\s+on\s+edge", joined, re.IGNORECASE):
            pack_note = joined.strip()
            break

    # --- 4. Print Area ---
    print_areas = []
    for row in df.values:
        joined = " ".join(row)
        matches = re.findall(r"Print\s*Area[^,;]*", joined, re.IGNORECASE)
        for m in matches:
            clean = re.sub(r"\s+", " ", m.strip())
            if clean not in print_areas:
                print_areas.append(clean)
    print_area_str = ", ".join(print_areas)

    # --- 5. Photocell ---
    photocell_w, photocell_h = 6, 12  # defaults
    for row in df.values:
        joined = " ".join(row)
        upper = joined.upper()
        if ("PHOTO" in upper or "MARK" in upper) and not re.search(r"KLD|COUNT|G\b", upper):
            nums = [float(n) for n in re.findall(r"(\d+(?:\.\d+)?)", joined)]
            nums = [n for n in nums if 2 <= n <= 50]
            if len(nums) >= 2:
                nums = sorted(nums)
                photocell_w, photocell_h = nums[0], nums[1]
                break
            elif len(nums) == 1:
                photocell_w = nums[0]
                photocell_h = max(12.0, photocell_w)
                break

    # --- 6. Top sequence ---
    top_seq_nums = []
    max_count = 0
    for i in range(len(df)):
        row = df.iloc[i].tolist()
        nums = clean_numeric_list(row)
        if len(nums) > max_count and len(nums) >= 5:
            max_count = len(nums)
            top_seq_nums = nums

    # --- 7. Side sequence ---
    side_seq_nums = []
    max_col = 0
    for c in df.columns:
        col = df[c].tolist()
        nums = clean_numeric_list(col)
        if len(nums) > max_col and len(nums) >= 3:
            max_col = len(nums)
            side_seq_nums = nums

    # --- 8. Trim sequences to match width/height within ±1 mm ---
    top_seq_trimmed = auto_trim_to_target(top_seq_nums, width_mm, tol=1.0)
    side_seq_trimmed = auto_trim_to_target(side_seq_nums, cut_length_mm, tol=1.0)

    if show_debug:
        st.write(f"Top seq raw sum = {sum(top_seq_nums):.1f}, target = {width_mm}")
        st.write(f"Trimmed top seq sum = {sum(top_seq_trimmed):.1f}")
        st.write(f"Side seq raw sum = {sum(side_seq_nums):.1f}, target = {cut_length_mm}")
        st.write(f"Trimmed side seq sum = {sum(side_seq_trimmed):.1f}")

    top_seq = ",".join([str(int(v)) if v.is_integer() else str(v) for v in top_seq_trimmed])
    side_seq = ",".join([str(int(v)) if v.is_integer() else str(v) for v in side_seq_trimmed])

    return (
        job_name,
        width_mm,
        cut_length_mm,
        top_seq,
        side_seq,
        pack_note,
        print_area_str,
        photocell_w,
        photocell_h,
    )

# ---------------------------------------------------
# Streamlit App Main Flow
# ---------------------------------------------------
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    try:
        (
            job_name,
            width_mm,
            cut_length_mm,
            top_seq,
            side_seq,
            pack_note,
            print_area_str,
            photocell_w,
            photocell_h,
        ) = extract_kld_data(df)

        output_df = pd.DataFrame([{
            "job_name": job_name,
            "width_mm": width_mm,
            "cut_length_mm": cut_length_mm,
            "top_seq": top_seq,
            "side_seq": side_seq,
            "pack_note": pack_note,
            "print_area": print_area_str,
            "photocell_w": photocell_w,
            "photocell_h": photocell_h,
            "photocell_offset_right_mm": 12,
            "stroke_mm": 0.25,
            "brand_label": "BRANDING"
        }])

        csv_bytes = output_df.to_csv(index=False).encode("utf-8")
        st.success(f"✅ Processed successfully for **{job_name}**")
        st.download_button("⬇️ Download CSV File", csv_bytes, f"{job_name}_converted.csv", "text/csv")
        st.dataframe(output_df)

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")

else:
    st.info("Please upload a KLD Excel file to begin.")
