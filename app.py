import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="KLD Excel → CSV Converter", layout="wide")
st.title("📏 KLD Excel → CSV Converter")
st.caption("Upload your Excel file and download the converted CSV in standard format.")

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])

def extract_kld_data(df):
    # --- 1. Job name ---
    job_name_row = df[df.apply(lambda x: x.astype(str).str.contains("Lam KLD for", case=False, na=False)).any(axis=1)].index[0]
    job_name = df.iloc[job_name_row].dropna().iloc[0].replace("Lam KLD for ", "").strip()

    # --- 2. Dimensions ---
    dim_row = df[df.apply(lambda x: x.astype(str).str.contains("Dimension", case=False, na=False)).any(axis=1)].index[0]
    dim_text = df.iloc[dim_row].dropna().iloc[0]
    width_mm, cut_length_mm = map(int, re.findall(r"\d+", dim_text)[:2])

    # --- 3. Detect numeric-only cells ---
    def is_number(x):
        try:
            float(x)
            return True
        except:
            return False

    numeric_mask = df.applymap(is_number)

    # --- 4. Top sequence (row with most numeric values) ---
    top_row = numeric_mask.sum(axis=1).idxmax()
    top_seq_vals = [
        str(int(float(v))) if str(v).replace(".", "", 1).isdigit() else str(v)
        for v in df.iloc[top_row].dropna().tolist()
        if is_number(v)
    ]
    top_seq = ",".join(top_seq_vals)

    # --- 5. Side sequence (column with most numeric values) ---
    side_col = numeric_mask.sum(axis=0).idxmax()
    side_seq_vals = [
        str(int(float(v))) if str(v).replace(".", "", 1).isdigit() else str(v)
        for v in df.iloc[:, side_col].dropna().tolist()
        if is_number(v)
    ]
    side_seq = ",".join(side_seq_vals)

    return job_name, width_mm, cut_length_mm, top_seq, side_seq


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

        csv_data = output_df.to_csv(index=False).encode("utf-8")
        st.success(f"✅ Converted successfully for **{job_name}**")
        st.download_button("⬇️ Download CSV File", csv_data, f"{job_name}_converted.csv", "text/csv")
        st.dataframe(output_df)

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")

else:
    st.info("Please upload a KLD Excel file to begin.")
