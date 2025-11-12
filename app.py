import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ---------------------------------------------------
# Streamlit App Configuration
# ---------------------------------------------------
st.set_page_config(page_title="KLD Excel → CSV Converter", layout="wide")
st.title("📏 KLD Excel → CSV Converter (Multi-Layout Smart Version)")
st.caption("Upload any KLD Excel — app auto-detects job name, dimensions, top & side sequences.")

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])


# ---------------------------------------------------
# Helper functions
# ---------------------------------------------------
def is_number(x):
    try:
        float(x)
        return True
    except:
        return False


def clean_numeric_list(seq):
    """Keep only numbers; cast 1.0→1 if integer."""
    out = []
    for v in seq:
        if is_number(v):
            f = float(v)
            out.append(str(int(f)) if f.is_integer() else str(f))
    return out


# ---------------------------------------------------
# Extraction Logic
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

    # --- 3. Top sequence: find row with most pure numeric cells ---
    top_seq = ""
    max_count = 0
    for i in range(len(df)):
        row = df.iloc[i].tolist()
        nums = clean_numeric_list(row)
        if len(nums) > max_count and len(nums) >= 5:  # avoid short false positives
            max_count = len(nums)
            top_seq = ",".join(nums)

    # --- 4. Side sequence: find column with most pure numeric cells ---
    side_seq = ""
    max_col = 0
    for c in df.columns:
        col = df[c].tolist()
        nums = clean_numeric_list(col)
        if len(nums) > max_col and len(nums) >= 3:
            max_col = len(nums)
            side_seq = ",".join(nums)

    return job_name, width_mm, cut_length_mm, top_seq, side_seq


# ---------------------------------------------------
# Streamlit App Main Flow
# ---------------------------------------------------
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    try:
        job_name, width_mm, cut_length_mm, top_seq, side_seq = extract_kld_data(df)

        output_df = pd.DataFrame([{
            "job_name": job_name,
            "width_mm": width_mm,
            "cut_length_mm": cut_length_mm,
            "top_seq": top_seq,
            "side_seq": side_seq,
            "photocell_w": 8,
            "photocell_h": 12,
            "photocell_offset_right_mm": 12,
            "stroke_mm": 0.25,
            "brand_label": "BRANDING"
        }])

        csv_bytes = output_df.to_csv(index=False).encode("utf-8")
        st.success(f"✅ Processed successfully for **{job_name}**")
        st.download_button("⬇️ Download CSV File", csv_bytes, f"{job_name}_converted.csv", "text/csv")
        st.dataframe(output_df)
        st.caption(f"Detected top_seq length: {len(top_seq.split(','))} | side_seq length: {len(side_seq.split(','))}")

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")

else:
    st.info("Please upload a KLD Excel file to begin.")
