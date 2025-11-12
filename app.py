import streamlit as st
import pandas as pd
import re
import numpy as np
import io

# ---------------------------
# Streamlit config
# ---------------------------
st.set_page_config(page_title="KLD Excel → CSV Converter", layout="wide")
st.title("📏 KLD Excel → CSV Converter (Robust for your file)")
st.caption("Preserves all header lines, extracts dimensions (many formats), and finds Print Area (finarea).")

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])
show_debug = st.checkbox("Show debug info", value=False)
show_raw = st.checkbox("Show raw sheet preview (top 80 rows)", value=False)

# ---------------------------
# Helpers
# ---------------------------
def is_number_string(s):
    try:
        float(str(s).strip().replace(",", ""))
        return True
    except:
        return False

def clean_numeric_list(seq):
    out = []
    for v in seq:
        s = str(v).strip().replace(",", "")
        if s == "" or s.lower() in ("nan", "none"):
            continue
        try:
            out.append(float(s))
        except:
            # try to catch numbers embedded in text like "126.0"
            m = re.search(r"(-?\d+(?:\.\d+)?)", s)
            if m:
                out.append(float(m.group(1)))
    return out

def first_pair_from_text(text):
    m = re.search(r"(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)", text)
    if m:
        return int(round(float(m.group(1)))), int(round(float(m.group(2))))
    # parentheses style (126*72)
    m2 = re.search(r"\(\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\s*\)", text)
    if m2:
        return int(round(float(m2.group(1)))), int(round(float(m2.group(2))))
    # fallback: two numbers anywhere
    nums = re.findall(r"(\d+(?:\.\d+)?)", text)
    if len(nums) >= 2:
        return int(round(float(nums[0]))), int(round(float(nums[1])))
    return 0, 0

def extract_dimensions_from_region(lines):
    """
    Given a list of text lines (strings), try multiple heuristics to extract width & cut_length.
    Returns (width_mm, cut_length_mm)
    """
    width = cut = 0

    # 1) Look for explicit pair in a single line e.g., "416 * 386" or "Printing Area - (126*72)"
    for ln in lines:
        w, c = first_pair_from_text(ln)
        if w and c:
            # try to determine order by neighboring words if possible
            low = ln.lower()
            # if 'width' occurs left of 'cut' in same line, assume order accordingly (rare)
            if re.search(r"width", low) and re.search(r"cut", low):
                # assume first corresponds to width, second to cut
                return w, c
            # if line explicitly says "Printing Area" or "Print Area", treat as printarea pair (width×height)
            return w, c

    # 2) If there are separate labeled lines "Width : 160mm" and "Length : 148mm" (or Cut)
    w_val = None
    c_val = None
    for ln in lines:
        low = ln.lower()
        if "width" in low and re.search(r"(\d+(?:\.\d+)?)", ln):
            v = re.search(r"(\d+(?:\.\d+)?)", ln).group(1)
            w_val = int(round(float(v)))
        if re.search(r"cut|length|cut-off|cutoff|cut off", low) and re.search(r"(\d+(?:\.\d+)?)", ln):
            v = re.search(r"(\d+(?:\.\d+)?)", ln).group(1)
            c_val = int(round(float(v)))

    if w_val and c_val:
        return w_val, c_val

    # 3) If not found, try any line that contains 'dimension' or 'width' etc and get first pair from it
    for ln in lines:
        if re.search(r"dimensi|width|cut|size|print", ln, re.IGNORECASE):
            w, c = first_pair_from_text(ln)
            if w and c:
                return w, c

    return 0, 0

def extract_finarea_line(line):
    """
    From a line that contains 'print' and 'area' (or 'printing area'), extract a clean representation.
    Prefer explicit numeric pair as 'AxB mm' if present; otherwise return trimmed original.
    """
    if not line:
        return ""
    low = line.lower()
    if "print" in low and "area" in low:
        # capture explicit pair if present
        w, c = first_pair_from_text(line)
        if w and c:
            return f"{w}x{c} mm"
        # else return the tail after 'Print Area' or the whole line stripped
        m = re.search(r"print[-\s]*area[\s\-:–—]*(.+)$", line, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        return line.strip()
    return ""

# ---------------------------
# Core extraction
# ---------------------------
def extract_kld_data(df):
    # ensure strings
    df = df.fillna("").astype(str)
    # strip fully blank rows
    df = df[df.apply(lambda r: any(str(x).strip() != "" for x in r), axis=1)].reset_index(drop=True)

    # collect header = consecutive non-empty lines from top until we detect numeric-data block
    header_lines = []
    start_row = 0

    # We'll scan up to first 60 rows to find the start of numeric region
    for i in range(min(60, len(df))):
        row = df.iloc[i].tolist()
        joined = " ".join([s.strip() for s in row if str(s).strip() != ""])
        # consider this a header line unless it is predominantly numeric
        numeric_cells = sum(1 for c in row if re.match(r"^\s*-?\d+(\.\d+)?\s*$", str(c).strip()))
        numeric_ratio = numeric_cells / max(1, len(row))
        if numeric_cells >= 3 or numeric_ratio > 0.6:
            start_row = i
            break
        if joined:
            header_lines.append(joined)
    else:
        # no explicit numeric region found in first 60 rows -> treat whole file as header
        start_row = len(df)

    # Keep all collected header lines (user wanted 4 lines preserved)
    job_name = "\n".join(header_lines) if header_lines else "Unknown"

    # Search region for dimension & print area: include a few header lines and a chunk after start_row
    search_region_start = max(0, start_row - 6)
    search_region_end = min(len(df), start_row + 60)
    region_lines = []
    for i in range(search_region_start, search_region_end):
        region_lines.append(" ".join([str(x).strip() for x in df.iloc[i].tolist() if str(x).strip() != ""]))

    # Extract dimensions robustly
    width_mm, cut_length_mm = extract_dimensions_from_region(region_lines)

    # If not found, try scanning entire file lines for common patterns
    if width_mm == 0 and cut_length_mm == 0:
        all_lines = [" ".join([str(x).strip() for x in df.iloc[i].tolist() if str(x).strip() != ""]) for i in range(len(df))]
        width_mm, cut_length_mm = extract_dimensions_from_region(all_lines)

    # Extract finarea from region lines that mention Print Area (prefer earliest)
    finarea = ""
    for ln in region_lines:
        if re.search(r"print[-\s]*area|printing\s+area", ln, re.IGNORECASE):
            fa = extract_finarea_line(ln)
            if fa:
                finarea = fa
                break
    # fallback: look further down
    if not finarea:
        for i in range(search_region_end, min(len(df), search_region_end + 40)):
            ln = " ".join([str(x).strip() for x in df.iloc[i].tolist() if str(x).strip() != ""])
            if re.search(r"print[-\s]*area|printing\s+area", ln, re.IGNORECASE):
                fa = extract_finarea_line(ln)
                if fa:
                    finarea = fa
                    break

    # If still empty, but dimensions exist, infer finarea
    if not finarea and width_mm and cut_length_mm:
        finarea = f"{width_mm}x{cut_length_mm} mm (inferred)"

    # pack note
    pack_note = ""
    for i in range(search_region_start, min(len(df), search_region_end+40)):
        joined = " ".join([str(x).strip() for x in df.iloc[i].tolist() if str(x).strip() != ""])
        if re.search(r"biscuits\s+on\s+edge", joined, re.IGNORECASE):
            pack_note = joined
            break

    # Photocell detection
    photocell_w, photocell_h = 6, 12
    for i in range(search_region_start, min(len(df), search_region_end+40)):
        joined = " ".join([str(x).strip() for x in df.iloc[i].tolist() if str(x).strip() != ""])
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

    # Numeric sequences:
    df_num = df.iloc[start_row:].reset_index(drop=True) if start_row < len(df) else pd.DataFrame(columns=df.columns)

    # Top sequence (rows)
    top_seq_nums, best_diff = [], float("inf")
    for i in range(min(len(df_num), 120)):
        nums = clean_numeric_list(df_num.iloc[i].tolist())
        if len(nums) >= 4:
            diff = abs(sum(nums) - cut_length_mm) if cut_length_mm > 0 else abs(sum(nums))
            if diff < best_diff:
                best_diff, top_seq_nums = diff, nums

    # Side sequence (columns)
    side_seq_nums, best_diff = [], float("inf")
    for c in df_num.columns:
        nums = clean_numeric_list(df_num[c].tolist())
        if len(nums) >= 3:
            diff = abs(sum(nums) - width_mm) if width_mm > 0 else abs(sum(nums))
            if diff < best_diff:
                best_diff, side_seq_nums = diff, nums

    # Trim sequences toward target
    def auto_trim(values, target):
        vals = values.copy()
        while len(vals) > 1 and target > 0 and sum(vals) > target + 1.0:
            vals.pop()
        return vals

    top_trim = auto_trim(top_seq_nums, cut_length_mm)
    side_trim = auto_trim(side_seq_nums, width_mm)

    def fmt(vals):
        if not vals:
            return ""
        out = []
        for v in vals:
            if float(v).is_integer():
                out.append(str(int(round(v))))
            else:
                out.append(str(v))
        return ",".join(out)

    top_seq = fmt(top_trim)
    side_seq = fmt(side_trim)

    return {
        "job_name": job_name,
        "width_mm": int(width_mm) if width_mm else 0,
        "cut_length_mm": int(cut_length_mm) if cut_length_mm else 0,
        "top_seq": top_seq,
        "side_seq": side_seq,
        "pack_note": pack_note,
        "finarea": finarea,
        "photocell_w": int(photocell_w) if float(photocell_w).is_integer() else photocell_w,
        "photocell_h": int(photocell_h) if float(photocell_h).is_integer() else photocell_h,
    }

# ---------------------------
# Main flow
# ---------------------------
if uploaded_file:
    ext = uploaded_file.name.split(".")[-1].lower()
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    try:
        if ext == "xls":
            import xlrd  # ensure xlrd==1.2.0
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Failed to read Excel file: {e}")
        st.stop()

    if show_raw:
        st.subheader("Raw Excel preview (top 80 rows)")
        st.dataframe(df.head(80))

    try:
        res = extract_kld_data(df)

        output_df = pd.DataFrame([{
            "job_name": res["job_name"],
            "width_mm": res["width_mm"],
            "cut_length_mm": res["cut_length_mm"],
            "top_seq": res["top_seq"],
            "side_seq": res["side_seq"],
            "pack_note": res["pack_note"],
            "finarea": res["finarea"],
            "photocell_w": res["photocell_w"],
            "photocell_h": res["photocell_h"],
            "photocell_offset_right_mm": 12,
            "stroke_mm": 0.25,
            "brand_label": "BRANDING"
        }])

        if show_debug:
            st.write("=== DEBUG INFO ===")
            st.write("Job name (lines):")
            for i, line in enumerate(res["job_name"].splitlines(), start=1):
                st.write(f"{i}. {line}")
            st.write(f"Detected width_mm: {res['width_mm']}, cut_length_mm: {res['cut_length_mm']}")
            st.write(f"finarea: {res['finarea']}")
            st.write(f"top_seq: {res['top_seq']}")
            st.write(f"side_seq: {res['side_seq']}")
            st.write(f"pack_note: {res['pack_note']}")
            st.write(f"photocell: {res['photocell_w']},{res['photocell_h']}")

        csv_bytes = output_df.to_csv(index=False, quoting=1).encode("utf-8")
        st.success(f"✅ Processed successfully for {uploaded_file.name}")
        st.download_button("⬇️ Download CSV File", csv_bytes, f"{uploaded_file.name}_converted.csv", "text/csv")
        st.dataframe(output_df)

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")
else:
    st.info("Please upload a KLD Excel file to begin.")
