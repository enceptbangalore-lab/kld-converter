import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="KLD Excel → SVG Generator", layout="wide")
st.title("📏 KLD Excel → SVG Generator (Final Combined)")
st.caption("Upload Excel → extract KLD → download SVG (AI-editable, mm units, DIELINE colour, Arial font).")

# === helper functions (same as your original) ==================================
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
    text = str(text)
    m = re.search(r"(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)", text)
    if m:
        return int(float(m.group(1))), int(float(m.group(2)))
    return 0, 0

def auto_trim_to_target(values, target, tol=1.0):
    vals = values.copy()
    while len(vals) > 1 and target > 0 and sum(vals) > target + tol:
        vals.pop()
    return vals

def extract_print_areas(lines):
    finarea_left = finarea_right = printarea = ""
    for ln in lines:
        if re.search(r"print\s*area", ln, re.IGNORECASE):
            pairs = re.findall(r"(\d+)\s*[*xX]\s*(\d+)", ln)
            if len(pairs) == 1:
                finarea_left = f"{pairs[0][0]}x{pairs[0][1]} mm"
            elif len(pairs) >= 2:
                finarea_left = f"{pairs[0][0]}x{pairs[0][1]} mm"
                finarea_right = f"{pairs[1][0]}x{pairs[1][1]} mm"
        if re.search(r"printing\s*area", ln, re.IGNORECASE):
            w, h = first_pair_from_text(ln)
            if w and h:
                printarea = f"{w}x{h} mm"
    return finarea_left, finarea_right, printarea

def extract_dimensions(lines):
    for ln in lines:
        if re.search(r"dimension|width|cut", ln, re.IGNORECASE):
            w, c = first_pair_from_text(ln)
            if w and c:
                return w, c
    return 0, 0

def extract_kld_data(df):
    df = df.fillna("").astype(str)
    df = df[df.apply(lambda r: any(str(x).strip() for x in r), axis=1)].reset_index(drop=True)
    header_lines = []
    start_row = 0
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
    finarea_left, finarea_right, printarea = extract_print_areas(search_lines)

    photocell_w, photocell_h = 6, 12
    for ln in search_lines:
        if re.search(r"photo|mark", ln, re.IGNORECASE):
            nums = [float(n) for n in re.findall(r"(\d+(?:\.\d+)?)", ln)]
            nums = [n for n in nums if 2 <= n <= 50]
            if len(nums) >= 2:
                nums.sort()
                photocell_w, photocell_h = nums[0], nums[1]
                break

    pack_note = ""
    for ln in search_lines:
        if re.search(r"biscuits\s+on\s+edge", ln, re.IGNORECASE):
            pack_note = ln.strip()
            break

    df_num = df.iloc[start_row:].reset_index(drop=True)
    top_seq_nums, side_seq_nums = [], []
    best_diff = float("inf")
    for i in range(len(df_num)):
        nums = clean_numeric_list(df_num.iloc[i].tolist())
        if len(nums) >= 4:
            diff = abs(sum(nums) - cut_length_mm) if cut_length_mm else sum(nums)
            if diff < best_diff:
                best_diff, top_seq_nums = diff, nums
    best_diff = float("inf")
    for c in df_num.columns:
        nums = clean_numeric_list(df_num[c].tolist())
        if len(nums) >= 3:
            diff = abs(sum(nums) - width_mm) if width_mm else sum(nums)
            if diff < best_diff:
                best_diff, side_seq_nums = diff, nums

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
        "finarea_left": finarea_left,
        "finarea_right": finarea_right,
        "printarea": printarea,
        "pack_note": pack_note,
        "photocell_w": photocell_w,
        "photocell_h": photocell_h,
        "photocell_offset_right_mm": 12,
        "stroke_mm": 0.25,
        "brand_label": "BRANDING",
    }

# === SVG builder ==============================================================
def make_svg(data):
    W = data["cut_length_mm"]
    H = data["width_mm"]
    top_seq = [float(x) for x in data["top_seq"].split(",") if x]
    side_seq = [float(x) for x in data["side_seq"].split(",") if x]
    dieline = "#7f00bf"
    stroke = data["stroke_mm"]
    pcw, pch = data["photocell_w"], data["photocell_h"]
    pc_off = data["photocell_offset_right_mm"]
    job = data["job_name"].replace("\n", " | ")

    # SVG header (mm units)
    out = []
    out.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}mm" height="{H}mm" viewBox="0 0 {W} {H}">')
    out.append(f'<style>text{{font-family:Arial;font-size:3.5mm;fill:{dieline};}} .dl{{stroke:{dieline};stroke-width:{stroke}mm;fill:none;}}</style>')
    out.append(f'<g id="MainOutline"><rect x="0" y="0" width="{W}" height="{H}" class="dl"/></g>')

    # vertical fold lines (top_seq)
    x = 0
    for v in top_seq[:-1]:
        x += v
        out.append(f'<line x1="{x}" y1="0" x2="{x}" y2="{H}" class="dl" stroke-dasharray="1,1"/>')

    # horizontal fold lines (side_seq)
    y = 0
    for v in side_seq[:-1]:
        y += v
        out.append(f'<line x1="0" y1="{y}" x2="{W}" y2="{y}" class="dl" stroke-dasharray="1,1"/>')

    # Photocell mark (right side)
    pcx = W - pc_off - pcw
    pcy = H / 2 - pch / 2
    out.append(f'<rect x="{pcx}" y="{pcy}" width="{pcw}" height="{pch}" class="dl"/>')

    # Labels
    out.append(f'<g id="Labels">')
    out.append(f'<text x="{W/2}" y="{-5}" text-anchor="middle">{job}</text>')
    out.append(f'<text x="{W/2}" y="{H+5}" text-anchor="middle">{round(H)}×{round(W)} mm</text>')
    out.append(f'</g>')

    out.append('</svg>')
    return "\n".join(out)

# === main streamlit logic =====================================================
uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])

if uploaded_file:
    ext = uploaded_file.name.split(".")[-1].lower()
    data = uploaded_file.read()
    uploaded_file.seek(0)

    if ext == "xls":
        import xlrd
        df = pd.read_excel(io.BytesIO(data), header=None, engine="xlrd")
    else:
        df = pd.read_excel(io.BytesIO(data), header=None, engine="openpyxl")

    try:
        res = extract_kld_data(df)
        csv_bytes = pd.DataFrame([res]).to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download CSV", csv_bytes, "kld_extracted.csv", "text/csv")

        svg_str = make_svg(res)
        st.success(f"✅ Extracted successfully for {uploaded_file.name}")
        st.download_button(
            "⬇️ Download SVG for Illustrator",
            svg_str.encode("utf-8"),
            "kld_layout.svg",
            "image/svg+xml"
        )

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")
else:
    st.info("Please upload a KLD Excel file to begin.")
