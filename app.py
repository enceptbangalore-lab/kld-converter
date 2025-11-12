import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ---------------------------------------------------
# Streamlit App Configuration
# ---------------------------------------------------
st.set_page_config(page_title="KLD Excel → CSV Converter", layout="wide")
st.title("📏 KLD Excel → CSV Converter (Refined with Area Detection)")
st.caption("Automatically detects job name, dimensions, valid KLD area, and photocell (filters irrelevant values).")

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])


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

    # --- 5. Photocell (filter product lines, use only 'photo'/'mark' lines) ---
    photocell_w, photocell_h = 6, 12  # defaults
    for row in df.values:
        joined = " ".join(row)
        upper = joined.upper()
        if ("PHOTO" in upper or "MARK" in upper) and not re.search(r"KLD|COUNT|G\b", upper):
            nums = [float(n) for n in re.findall(r"(\d+(?:\.\d+)?)", joined)]
            nums = [n for n in nums if 2 <= n <= 50]  # valid mm range filter
            if len(nums) >= 2:
                nums = sorted(nums)
                photocell_w, photocell_h = nums[0], nums[1]
                break
            elif len(nums) == 1:
                photocell_w = nums[0]
                photocell_h = max(12.0, photocell_w)
                break

    # --- 6. Top sequence (remove outside KLD area noise) ---
    top_seq = ""
    max_count = 0
    for i in range(len(df)):
        row = df.iloc[i].tolist()
        nums = clean_numeric_list(row)
        if len(nums) > max_count and len(nums) >= 5:
            max_count = len(nums)
            top_seq_nums = nums

    # Filter out any trailing tiny segment (e.g., ≤15 mm or after huge jump)
    if max_count > 0:
        diffs = [abs(top_seq_nums[i+1] - top_seq_nums[i]) for i in range(len(top_seq_nums)-1)]
        cutoff = len(top_seq_nums)
        if top_seq_nums[-1] <= 15 or (len(diffs) > 0 and diffs[-1] > 100):
            cutoff -= 1
        top_seq_nums = top_seq_nums[:cutoff]
        top_seq = ",".join([str(int(v)) if v.is_integer() else str(v) for v in top_seq_nums])

    # --- 7. Side sequence ---
    side_seq = ""
    max_col = 0
    for c in df.columns:
        col = df[c].tolist()
        nums = clean_numeric_list(col)
        if len(nums) > max_col and len(nums) >= 3:
            max_col = len(nums)
            side_seq = ",".join([str(int(v)) if v.is_integer() else str(v) for v in nums])

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

        # Build output dataframe
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
        st.caption(
            f"Detected top_seq length: {len(top_seq.split(','))} | "
            f"side_seq length: {len(side_seq.split(','))}"
        )

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")

else:
    st.info("Please upload a KLD Excel file to begin.")
