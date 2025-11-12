import streamlit as st
import pandas as pd
import re
import numpy as np
import io

# ---------------------------
# Streamlit config
# ---------------------------
st.set_page_config(page_title="KLD Excel → CSV Converter", layout="wide")
st.title("📏 KLD Excel → CSV Converter (Robust header, dims & finarea extraction)")
st.caption("Keeps all header lines, finds dimensions reliably, extracts Print Area (finarea).")

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])
show_debug = st.checkbox("Show debug info", value=False)
show_raw = st.checkbox("Show raw sheet preview (top 60 rows)", value=False)

# ---------------------------
# Helpers
# ---------------------------
def is_number(x):
    try:
        float(x)
        return True
    except:
        return False


def clean_numeric_list(seq):
    out = []
    for v in seq:
        # Accept values that look numeric (strip commas)
        s = str(v).strip().replace(",", "")
        if s == "":
            continue
        try:
            out.append(float(s))
        except:
            pass
    return out


def auto_trim_to_target(values, target, tol=1.0):
    vals = values.copy()
    # keep dropping last until <= target + tol
    while len(vals) > 1 and sum(vals) > target + tol:
        vals.pop()
    return vals


def extract_dimensions_from_line(line):
    """
    Try multiple heuristics to extract width_mm and cut_length_mm from a text line.
    Returns (width_mm:int or 0, cut_length_mm:int or 0)
    """
    text = line.strip()
    if not text:
        return 0, 0

    # 1) Look for explicit labeled pairs like "Width ... 416" and "Cut ... 386" in the same or adjacent text
    # capture numeric tokens with their positions
    nums = [(m.group(0), m.start()) for m in re.finditer(r"(\d+(?:\.\d+)?)", text)]
    lowered = text.lower()

    # If both words 'width' and 'cut' appear, decide mapping by order
    if "width" in lowered or "cut" in lowered or "cut off" in lowered or "cutoff" in lowered:
        # If pattern like "width * cut off length" then look for the pair format first
        pair = re.search(r"(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)", text)
        if pair:
            a = float(pair.group(1))
            b = float(pair.group(2))
            # Determine label order in nearby text: find index of 'width' and 'cut'
            idx_w = lowered.find("width")
            idx_c = lowered.find("cut")
            if idx_w != -1 and idx_c != -1:
                if idx_w < idx_c:
                    return int(round(a)), int(round(b))
                else:
                    return int(round(b)), int(round(a))
            # fallback: assume first is width, second cut
            return int(round(a)), int(round(b))

        # If labels present but numbers separated, try to find numbers near labels
        # Example: "Width (mm) 416  Cut off length (mm) 386"
        # We'll try to capture "...width...(\d+)" and "...cut...(\d+)"
        w_match = re.search(r"width[^\d\n]*?(\d+(?:\.\d+)?)", lowered, re.IGNORECASE)
        c_match = re.search(r"cut[^\d\n]*?(\d+(?:\.\d+)?)", lowered, re.IGNORECASE)
        if w_match or c_match:
            w = int(round(float(w_match.group(1)))) if w_match else 0
            c = int(round(float(c_match.group(1)))) if c_match else 0
            if w or c:
                return w, c

    # 2) Look for pair with '*' or 'x' without labels, assume order is width * cut
    pair = re.search(r"(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)", text)
    if pair:
        a = float(pair.group(1))
        b = float(pair.group(2))
        return int(round(a)), int(round(b))

    # 3) If the line contains exactly two numbers and no clear labels, assume first = width, second = cut
    all_nums = re.findall(r"(\d+(?:\.\d+)?)", text)
    if len(all_nums) >= 2:
        a = float(all_nums[0])
        b = float(all_nums[1])
        return int(round(a)), int(round(b))

    return 0, 0


def extract_finarea_from_line(line):
    """
    Extract finarea info from a line containing words like 'Print Area' or variants.
    Returns the relevant substring (after colon/dash) or the full line if that's most sensible.
    """
    text = line.strip()
    lower = text.lower()
    if "print" in lower and "area" in lower:
        # If there's a colon or dash, prefer text after it
        m = re.search(r"print\s*[-:–—]?\s*area\s*[:\-–—]?\s*(.+)$", lower, re.IGNORECASE)
        if m:
            # return original-cased remainder for clarity
            remainder = text[m.end(0) - len(m.group(1)):]
            return remainder.strip()
        # else return full line
        return text
    # also accept "printarea" or "print-area"
    if re.search(r"print[-\s]*area", lower):
        return text
    return ""


# ---------------------------
# Extraction main logic
# ---------------------------
def extract_kld_data(df):
    df = df.fillna("").astype(str)

    # drop fully blank rows at top and within detection window
    df = df[df.apply(lambda r: any(str(x).strip() != "" for x in r), axis=1)].reset_index(drop=True)

    header_lines = []
    finarea = ""
    start_row = 0

    # scan top rows to collect header lines and find where numeric section begins
    for i in range(min(60, len(df))):
        row = df.iloc[i].tolist()
        # join non-empty cells
        line_text = " ".join([s.strip() for s in row if str(s).strip() != ""])

        # keep finarea if present anywhere in these header lines
        if "print" in line_text.lower() and "area" in line_text.lower():
            # extract finarea details
            fin_candidate = extract_finarea_from_line(line_text)
            if fin_candidate:
                finarea = fin_candidate

        # Determine if this row is predominantly numeric (start of numeric blocks)
        numeric_cells = [c for c in row if re.match(r"^\s*-?\d+(\.\d+)?\s*$", str(c).strip())]
        numeric_ratio = len(numeric_cells) / max(1, len(row))

        # Heuristic: if many numeric cells or several numeric tokens, that's start of numeric block
        if len(numeric_cells) >= 3 or numeric_ratio > 0.5:
            start_row = i
            break

        # otherwise treat non-empty lines as header lines
        if line_text:
            header_lines.append(line_text)

    # If finarea still empty, scan a bit further (next 40 rows) to catch Print Area lines
    if not finarea:
        for i in range(start_row, min(start_row + 40, len(df))):
            row = df.iloc[i].tolist()
            line_text = " ".join([s.strip() for s in row if str(s).strip() != ""])
            if "print" in line_text.lower() and "area" in line_text.lower():
                fin_candidate = extract_finarea_from_line(line_text)
                if fin_candidate:
                    finarea = fin_candidate
                    break

    # Keep ALL header lines joined with newline (user requested 4 lines preserved)
    job_name = "\n".join(header_lines) if header_lines else "Unknown"

    # slice the dataframe after header start
    df_num = df.iloc[start_row:].reset_index(drop=True)

    # DIMENSION detection: search through the numeric section and a few lines above for dimension patterns
    width_mm, cut_length_mm = 0, 0
    # look in first 30 rows of numeric section + a few rows above
    search_region_start = max(0, start_row - 4)
    search_region_end = min(len(df), start_row + 40)
    for i in range(search_region_start, search_region_end):
        joined = " ".join(df.iloc[i].tolist())
        w, c = extract_dimensions_from_line(joined)
        if w or c:
            width_mm, cut_length_mm = w, c
            break

    # Additional heuristic: if swapped (e.g., width much larger than cut or vice versa), try columns-based detection
    if width_mm == 0 and cut_length_mm == 0:
        # Try scanning for a line containing words like 'dimensions' and numbers anywhere below
        for i in range(search_region_start, search_region_end):
            joined = " ".join(df.iloc[i].tolist())
            if re.search(r"dimensi|size|width|cut", joined, re.IGNORECASE):
                w, c = extract_dimensions_from_line(joined)
                if w or c:
                    width_mm, cut_length_mm = w, c
                    break

    # Pack note (same heuristic as before)
    pack_note = ""
    for i in range(start_row, min(len(df), start_row + 80)):
        joined = " ".join(df.iloc[i].tolist())
        if re.search(r"biscuits\s+on\s+edge", joined, re.IGNORECASE):
            pack_note = joined.strip()
            break

    # Photocell detection - same robust pattern
    photocell_w, photocell_h = 6, 12
    for i in range(start_row, min(len(df), start_row + 60)):
        joined = " ".join(df.iloc[i].tolist())
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

    # Top sequence: find row with many numbers whose sum best matches cut_length_mm
    top_seq_nums, best_diff = [], float("inf")
    for i in range(start_row, min(len(df), start_row + 120)):
        nums = clean_numeric_list(df.iloc[i].tolist())
        if len(nums) >= 4:  # lowered requirement to be more permissive
            if cut_length_mm > 0:
                diff = abs(sum(nums) - cut_length_mm)
            else:
                diff = sum(nums)  # fallback if no cut_length_mm
            if diff < best_diff:
                best_diff, top_seq_nums = diff, nums

    # Side sequence: examine columns for vertical sequences summing to width_mm
    side_seq_nums, best_diff = [], float("inf")
    for c in df_num.columns:
        nums = clean_numeric_list(df_num[c].tolist())
        if len(nums) >= 3:
            if width_mm > 0:
                diff = abs(sum(nums) - width_mm)
            else:
                diff = sum(nums)
            if diff < best_diff:
                best_diff, side_seq_nums = diff, nums

    # Trim sequences toward target dimensions if targets exist
    top_seq_trimmed = auto_trim_to_target(top_seq_nums, cut_length_mm, tol=1.0) if cut_length_mm > 0 else top_seq_nums
    side_seq_trimmed = auto_trim_to_target(side_seq_nums, width_mm, tol=1.0) if width_mm > 0 else side_seq_nums

    # Format sequences to comma separated integers when possible
    def fmt_seq(vals):
        out = []
        for v in vals:
            if float(v).is_integer():
                out.append(str(int(round(v))))
            else:
                out.append(str(v))
        return ",".join(out) if out else ""

    top_seq = fmt_seq(top_seq_trimmed)
    side_seq = fmt_seq(side_seq_trimmed)

    # If finarea empty, attempt to generate it from common dimension lines (fallback)
    if not finarea:
        # If we have width & cut, produce a common print area string
        if width_mm and cut_length_mm:
            finarea = f"{width_mm} x {cut_length_mm} mm (inferred)"
        else:
            # scan near top for any line that contains "print" or "area" words to use as finarea
            for i in range(search_region_start, search_region_end):
                joined = " ".join(df.iloc[i].tolist())
                if "print" in joined.lower() or "area" in joined.lower():
                    finarea = joined.strip()
                    break

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
# Streamlit main flow
# ---------------------------
if uploaded_file:
    ext = uploaded_file.name.split(".")[-1].lower()
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    # read file using appropriate engine
    try:
        if ext == "xls":
            # requires xlrd==1.2.0
            import xlrd  # noqa: F401
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Failed to read Excel file: {e}")
        st.stop()

    # optional raw preview
    if show_raw:
        try:
            st.subheader("Raw Excel preview (top 60 rows)")
            st.dataframe(df.head(60))
        except Exception:
            # fallback to text preview
            txt = "\n".join(" | ".join([str(c) for c in row]) for row in df.head(60).values)
            st.text(txt)

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
            st.write("Detected header lines (kept all):")
            for i, ln in enumerate(res["job_name"].splitlines(), start=1):
                st.write(f"{i}. {ln}")
            st.write(f"finarea: {res['finarea']}")
            st.write(f"width_mm: {res['width_mm']}, cut_length_mm: {res['cut_length_mm']}")
            st.write(f"Top seq (raw/trimmed): {res['top_seq']}")
            st.write(f"Side seq (raw/trimmed): {res['side_seq']}")
            st.write(f"pack_note: {res['pack_note']}")
            st.write(f"photocell_w,h: {res['photocell_w']},{res['photocell_h']}")

        csv_bytes = output_df.to_csv(index=False, quoting=1).encode("utf-8")
        st.success(f"✅ Processed successfully for {uploaded_file.name}")
        st.download_button("⬇️ Download CSV File", csv_bytes, f"{uploaded_file.name}_converted.csv", "text/csv")
        st.dataframe(output_df)

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")

else:
    st.info("Please upload a KLD Excel file to begin.")
