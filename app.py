import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ---------------------------------------------------
# Streamlit App Configuration
# ---------------------------------------------------
st.set_page_config(page_title="KLD Excel → CSV Converter", layout="wide")
st.title("📏 KLD Excel → CSV Converter")
st.caption("Upload your Excel file and download the converted CSV in standard format.")

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])


# ---------------------------------------------------
# Core Logic: Extract KLD Data
# ---------------------------------------------------
def extract_kld_data(df):
    # --- 1. Job name ---
    job_name_row = df[df.apply(lambda x: x.astype(str).str.contains("Lam KLD for", case=False, na=False)).any(axis=1)].index[0]
    job_name = df.iloc[job_name_row].dropna().iloc[0].replace("Lam KLD for ", "").strip()

    # --- 2. Dimensions ---
    dim_row = df[df.apply(lambda x: x.astype(str).str.contains("Dimension", case=False, na=False)).any(axis=1)].index[0]
    dim_text = df.iloc[dim_row].dropna().iloc[0]
    width_mm, cut_length_mm = map(int, re.findall(r"\d+", dim_text)[:2])

    # --- 3. Top sequence (from row containing 310 or similar numeric pattern) ---
    top_row_candidates = df[df.apply(lambda x: x.astype(str).str.contains("310", na=False)).any(axis=1)]
    if not top_row_candidates.empty:
        top_row = top_row_candidates.index[0]
    else:
        top_row = 10  # fallback (typical row index for top seq)

    top_vals = [
        str(int(float(v))) if str(v).replace(".", "", 1).isdigit() else str(v)
        for v in df.iloc[top_row].dropna().tolist()
        if re.match(r"^\d+(\.\d+)?$", str(v))
    ]
    top_seq = ",".join(top_vals)

    # --- 4. Side sequence (column 5 between rows 11–48) ---
    side_vals = []
    for i in range(11, 49):
        v = df.iat[i, 5] if 5 < df.shape[1] else ""
        if pd.notna(v) and re.match(r"^\d+(\.\d+)?$", str(v).strip()):
            val = str(int(float(v))) if float(v).is_integer() else str(v)
            side_vals.append(val)
    side_seq = ",".join(side_vals)

    return job_name, width_mm, cut_length_mm, top_seq, side_seq


# ---------------------------------------------------
# Main App Logic
# ---------------------------------------------------
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    try:
        job_name, width_mm, cut_length_mm, top_seq, side_seq = extract_kld_data(df)

        # Create output CSV in template format
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

        # Convert to downloadable CSV
        csv_data = output_df.to_csv(index=False).encode("utf-8")
        st.success(f"✅ Converted successfully for **{job_name}**")
        st.download_button("⬇️ Download CSV File", csv_data, f"{job_name}_converted.csv", "text/csv")
        st.dataframe(output_df)

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")

else:
    st.info("Please upload a KLD Excel file to begin.")
