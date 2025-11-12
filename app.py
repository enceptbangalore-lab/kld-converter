import streamlit as st
import pandas as pd
import re
import numpy as np
import io

# ---------------------------------------------------
# Streamlit App Configuration
# ---------------------------------------------------
st.set_page_config(page_title="KLD Excel → CSV Converter", layout="wide")
st.title("📏 KLD Excel → CSV Converter (Refined Header Detection)")
st.caption(
    "Reads KLD Excel (.xls/.xlsx), detects header lines cleanly, "
    "keeps them under job_name, moves Print Area to finarea, and trims sequences (±1 mm)."
)

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
            out.append(float(v))
    return out


def auto_trim_to_target(values, target, tol=1.0):
    vals = values.copy()
    while sum(vals) > target + tol and len(vals) > 1:
        vals.pop()
    return vals


# ---------------------------------------------------
# Core Extraction Logic
# ---------------------------------------------------
def extract_kld_data(df):
    df = df.fillna("").astype(str)

    header_lines = []
    finarea = ""
    start_row = 0

    # --- Improved header detection ---
    for i in range(min(20, len(df))):
        row = df.iloc[i].tolist()
        line_text = " ".join([s.strip() for s in row if str(s).strip() != ""])
        numeric_values = clean_numeric_list(row)

        # Identify print area separately
        if re.search(r"Print\s*Area", line_text, re.IGNORECASE):
            finarea = line_text.strip()
            continue

        # If a line has 2+ numbers, assume numeric data starts here
        if len(numeric_values) >= 2:
            start_row = i
            break

        # Otherwise treat as header text
        if line_text:
            header_lines.append(line_text)

    job_name = "\n".join(header_lines) if header_lines else "Unknown"

    # Slice off header section
    df = df.iloc[start_row:].reset_index(drop=True)

    # --- Dimensions ---
    width_mm = cut_length_mm = 0
    for row in df.values:
        joined = " ".join(row)
        if "DIMENSION" in joined.upper():
            nums = re.findall(r"\d+", joined)
            if len(nums) >= 2:
                width_mm, cut_length_mm = map(int, nums[:2])
                break

    # --- Pack note ---
    pack_note = ""
    for row in df.values:
        joined = " ".join(row)
        if re.search(r"biscuits\s+on\s+edge", joined, re.IGNORECASE):
            pack_note = joined.strip()
            break

    # --- Photocell ---
    photocell_w, photocell_h = 6, 12
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

    # --- Top sequence (closest to cut length) ---
    top_seq_nums, best_diff = [], float("inf")
    for i in range(len(df)):
        nums = clean_numeric_list(df.iloc[i].tolist())
        if len(nums) >= 5:
            diff = abs(sum(nums) - cut_length_mm)
            if diff < best_diff:
                best_diff, top_seq_nums = diff, nums

    # --- Side sequence (closest to width) ---
    side_seq_nums, best_diff = [], float("inf")
    for c in df.columns:
        nums = clean_numeric_list(df[c].tolist())
        if len(nums) >= 3:
            diff = abs(sum(nums) - width_mm)
            if diff < best_diff:
                best_diff, side_seq_nums = diff, nums

    # --- Trim to match dimensions ---
    top_seq_trimmed = auto_trim_to_target(top_seq_nums, cut_length_mm, tol=1.0)
    side_seq_trimmed = auto_trim_to_target(side_seq_nums, width_mm, tol=1.0)

    if show_debug:
        st.write("Header lines detected:", header_lines)
        st.write(f"finarea: {finarea}")
        st.write(f"Top seq raw sum = {sum(top_seq_nums):.1f}, target (cut length) = {cut_length_mm}")
        st.write(f"Trimmed top seq sum = {sum(top_seq_trimmed):.1f}")
        st.write(f"Side seq raw sum = {sum(side_seq_nums):.1f}, target (width) = {width_mm}")
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
        finarea,
        photocell_w,
        photocell_h,
    )


# ---------------------------------------------------
# Streamlit App Main Flow
# ---------------------------------------------------
if uploaded_file:
    # Smart Excel reader for .xls/.xlsx
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    ext = uploaded_file.name.split(".")[-1].lower()

    try:
        if ext == "xls":
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None, engine="xlrd")
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Failed to read Excel file: {e}")
        st.stop()

    try:
        (
            job_name,
            width_mm,
            cut_length_mm,
            top_seq,
            side_seq,
            pack_note,
            finarea,
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
            "finarea": finarea,
            "photocell_w": photocell_w,
            "photocell_h": photocell_h,
            "photocell_offset_right_mm": 12,
            "stroke_mm": 0.25,
            "brand_label": "BRANDING"
        }])

        csv_bytes = output_df.to_csv(index=False, quoting=1).encode("utf-8")
        st.success(f"✅ Processed successfully for {uploaded_file.name}")
        st.download_button("⬇️ Download CSV File", csv_bytes, f"{uploaded_file.name}_converted.csv", "text/csv")
        st.dataframe(output_df)

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")

else:
    st.info("Please upload a KLD Excel file to begin.")
