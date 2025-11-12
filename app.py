import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="KLD Excel → CSV Converter", layout="wide")
st.title("📏 KLD Excel → CSV Converter (Refined v3)")
st.caption(
    "Keeps full job header (including Count/kg), removes non-header lines like Photocell info, "
    "and extracts both finarea (left/right) and printarea correctly."
)

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])
show_debug = st.checkbox("Show debug info", value=False)
show_raw = st.checkbox("Show raw sheet preview", value=False)

# ---------------------------------------------------
# Helpers
# ---------------------------------------------------
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
    """Extract numeric pair like 15*20, 15 x 20, (15*20)."""
    text = str(text)
    m = re.search(r"(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)", text)
    if m:
        return int(float(m.group(1))), int(float(m.group(2)))
    return 0, 0


def extract_print_areas(lines):
    """Extract Print Area (left/right) and Printing Area (main)."""
    finarea_left = finarea_right = printarea = ""
    for ln in lines:
        # PRINT AREA LEFT/RIGHT
        if re.search(r"print\s*area", ln, re.IGNORECASE):
            pairs = re.findall(r"(\d+)\s*[*xX]\s*(\d+)", ln)
            if len(pairs) == 1:
                finarea_left = f"{pairs[0][0]}x{pairs[0][1]} mm"
            elif len(pairs) >= 2:
                finarea_left = f"{pairs[0][0]}x{pairs[0][1]} mm"
                finarea_right = f"{pairs[1][0]}x{pairs[1][1]} mm"

        # PRINTING AREA (separate from Print Area)
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
# Core Extraction
# ---------------------------------------------------
def extract_kld_data(df):
    df = df.fillna("").astype(str)
    df = df[df.apply(lambda r: any(str(x).strip() for x in r), axis=1)].reset_index(drop=True)

    header_lines = []
    start_row = 0

    # Collect header lines until numeric data
    for i in range(min(60, len(df))):
        row = df.iloc[i].tolist()
        line_text = " ".join([s.strip() for s in row if s.strip()])
        numeric_count = sum(1 for c in row if re.match(r"^\d+(\.\d+)?$", str(c)))
        if numeric_count >= 3:
            start_row = i
            break
        if line_text:
            header_lines.append(line_text)

    # --- Clean header lines ---
    exclude_patterns = [r"photocell", r"\[\s*\d+", r"\*\s*\d+\s*\]\s*mm"]
    while header_lines and any(re.search(pat, header_lines[-1], re.IGNORECASE) for pat in exclude_patterns):
        header_lines.pop()
    job_name = "\n".join(header_lines) if header_lines else "Unknown"

    # --- Extract sections ---
    search_lines = [
        " ".join([str(x).strip() for x in df.iloc[i].tolist() if str(x).strip()])
        for i in range(max(0, start_row - 10), min(len(df), start_row + 50))
    ]

    width_mm, cut_length_mm = extract_dimensions(search_lines)
    finarea_left, finarea_right, printarea = extract_print_areas(search_lines)

    # --- Photocell detection ---
    photocell_w, photocell_h = 6, 12
    for ln in search_lines:
        if re.search(r"photo|mark", ln, re.IGNORECASE):
            nums = [float(n) for n in re.findall(r"(\d+(?:\.\d+)?)", ln)]
            nums = [n for n in nums if 2 <= n <= 50]
            if len(nums) >= 2:
                nums.sort()
                photocell_w, photocell_h = nums[0], nums[1]
                break

    # --- Pack note ---
    pack_note = ""
    for ln in search_lines:
        if re.search(r"biscuits\s+on\s+edge", ln, re.IGNORECASE):
            pack_note = ln.strip()
            break

    return {
        "job_name": job_name,
        "width_mm": width_mm,
        "cut_length_mm": cut_length_mm,
        "finarea_left": finarea_left,
        "finarea_right": finarea_right,
        "printarea": printarea,
        "pack_note": pack_note,
        "photocell_w": photocell_w,
        "photocell_h": photocell_h,
    }


# ---------------------------------------------------
# Streamlit Execution
# ---------------------------------------------------
if uploaded_file:
    ext = uploaded_file.name.split(".")[-1].lower()
    data = uploaded_file.read()
    uploaded_file.seek(0)

    if ext == "xls":
        import xlrd  # xlrd==1.2.0 required
        df = pd.read_excel(io.BytesIO(data), header=None, engine="xlrd")
    else:
        df = pd.read_excel(io.BytesIO(data), header=None, engine="openpyxl")

    if show_raw:
        st.subheader("Raw Excel (Top 40 Rows)")
        st.dataframe(df.head(40))

    try:
        res = extract_kld_data(df)
        output_df = pd.DataFrame([{
            "job_name": res["job_name"],
            "width_mm": res["width_mm"],
            "cut_length_mm": res["cut_length_mm"],
            "finarea_left": res["finarea_left"],
            "finarea_right": res["finarea_right"],
            "printarea": res["printarea"],
            "pack_note": res["pack_note"],
            "photocell_w": res["photocell_w"],
            "photocell_h": res["photocell_h"],
            "photocell_offset_right_mm": 12,
            "stroke_mm": 0.25,
            "brand_label": "BRANDING",
        }])

        if show_debug:
            st.write("=== DEBUG INFO ===")
            st.write("Header lines:")
            for i, l in enumerate(res["job_name"].splitlines(), start=1):
                st.write(f"{i}. {l}")
            st.write(f"Width×Cut: {res['width_mm']}×{res['cut_length_mm']}")
            st.write(f"FinArea Left: {res['finarea_left']} | FinArea Right: {res['finarea_right']}")
            st.write(f"PrintArea: {res['printarea']}")
            st.write(f"Pack Note: {res['pack_note']}")
            st.write(f"Photocell: {res['photocell_w']}×{res['photocell_h']}")

        csv_bytes = output_df.to_csv(index=False).encode("utf-8")
        st.success(f"✅ Processed successfully for {uploaded_file.name}")
        st.download_button(
            "⬇️ Download CSV File",
            csv_bytes,
            f"{uploaded_file.name}_converted.csv",
            "text/csv"
        )
        st.dataframe(output_df)

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")
else:
    st.info("Please upload a KLD Excel file to begin.")
