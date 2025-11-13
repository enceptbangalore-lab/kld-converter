import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="KLD Excel → SVG Generator", layout="wide")
st.title("📏 KLD Excel → SVG Generator (All Marks Inside Artboard)")
st.caption("Reads KLD Excel, extracts dimensions and sequences, and generates editable SVG dieline for Illustrator QC with all elements inside the artboard.")

# ---------------------------------------------------
# Helper functions
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

    header_lines, start_row = [], 0
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
    W = width_mm
    for c in df_num.columns:
        nums = clean_numeric_list(df_num[c].tolist())
        n = len(nums)
        if n < 3:
            continue
        for i in range(n):
            s = 0.0
            for j in range(i, n):
                s += nums[j]
                if j - i + 1 >= 3:
                    diff = abs(s - W) if W else s
                    if diff < best_diff:
                        best_diff = diff
                        side_seq_nums = nums[i:j+1]

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
    }


def make_svg(data):
    def parse_seq(src):
        if not src:
            return []
        if isinstance(src, (list, tuple)):
            return [float(x) for x in src if x]
        parts = re.split(r"[,;|]", str(src))
        return [float(p.strip()) for p in parts if re.match(r"^\d+(\.\d+)?$", p.strip())]

    W = float(data.get("cut_length_mm") or 0)
    H = float(data.get("width_mm") or 0)
    top_seq = parse_seq(data.get("top_seq"))
    side_seq = parse_seq(data.get("side_seq"))

    # --- Artboard expansion ---
    extra = 60.0
    margin = extra / 2
    canvas_W = W + extra
    canvas_H = H + extra

    # --- Style ---
    dieline = "#92278f"
    stroke_pt = 0.356
    font_mm = 1.5
    tick_short = 5.0
    top_shift_up = 5.0
    left_shift_left = 5.0
    crop_off = 5.0
    crop_len = 5.0
    left_text_shift_right = 6.0
    top_text_shift_down = 4.0

    out = []
    out.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{canvas_W}mm" height="{canvas_H}mm" viewBox="0 0 {canvas_W} {canvas_H}">')
    out.append('<defs><style><![CDATA[')
    out.append(f'.dieline{{stroke:{dieline};stroke-width:{stroke_pt}pt;fill:none;}}')
    out.append(f'.text{{font-family:Arial; font-size:{font_mm}mm; fill:{dieline};}}')
    out.append(']]></style></defs>')

    # --- Outer dieline ---
    out.append(f'<rect x="{margin}" y="{margin}" width="{W}" height="{H}" class="dieline"/>')

    # --- Measurement ticks + text ---
    out.append('<g id="Measurements">')

    # TOP ticks & labels
    x = 0
    out.append(f'<line x1="{margin}" y1="{margin - top_shift_up}" x2="{margin}" y2="{margin - top_shift_up - tick_short}" class="dieline"/>')
    for v in top_seq:
        x += v
        out.append(f'<line x1="{margin + x}" y1="{margin - top_shift_up}" x2="{margin + x}" y2="{margin - top_shift_up - tick_short}" class="dieline"/>')
        mid = x - v / 2
        out.append(f'<text x="{margin + mid}" y="{margin - top_shift_up - tick_short - 1 + top_text_shift_down}" text-anchor="middle" class="text">{int(v)}</text>')

    # LEFT ticks & labels
    y = 0
    out.append(f'<line x1="{margin - left_shift_left}" y1="{margin}" x2="{margin - left_shift_left - tick_short}" y2="{margin}" class="dieline"/>')
    for v in side_seq:
        y += v
        out.append(f'<line x1="{margin - left_shift_left}" y1="{margin + y}" x2="{margin - left_shift_left - tick_short}" y2="{margin + y}" class="dieline"/>')
        midY = y - v / 2
        lx = margin - left_shift_left - tick_short - 2 + left_text_shift_right
        out.append(f'<text x="{lx}" y="{margin + midY}" transform="rotate(-90 {lx} {margin + midY})" text-anchor="middle" class="text">{int(v)}</text>')

    out.append('</g>')

    # --- Crop Marks ---
    out.append('<g id="CropMarks">')
    out.append(f'<line x1="{margin + W + crop_off}" y1="{margin + H}" x2="{margin + W + crop_off + crop_len}" y2="{margin + H}" class="dieline"/>')
    out.append(f'<line x1="{margin + W + crop_off}" y1="{margin}" x2="{margin + W + crop_off + crop_len}" y2="{margin}" class="dieline"/>')
    out.append('</g>')

    # --- Photocell Mark (TOP-RIGHT) ---
    photocell_w, photocell_h = 6, 12
    pc_x = margin + W - photocell_w
    pc_y = margin
    out.append(f'<rect x="{pc_x}" y="{pc_y}" width="{photocell_w}" height="{photocell_h}" class="dieline"/>')

    # --- Width Marker ---
    total_width = sum(side_seq)
    width_line_x = pc_x + photocell_w + 4
    out.append(f'<line x1="{width_line_x}" y1="{margin}" x2="{width_line_x}" y2="{margin + total_width}" class="dieline"/>')
    midY = margin + total_width / 2
    out.append(f'<text x="{width_line_x + 6}" y="{midY}" transform="rotate(-90 {width_line_x + 6} {midY})" text-anchor="middle" class="text">width = {int(total_width)} mm</text>')

    # --- Height Marker ---
    total_height = sum(top_seq)
    height_y = margin + total_width + 5
    out.append(f'<line x1="{margin}" y1="{height_y}" x2="{margin + total_height}" y2="{height_y}" class="dieline"/>')
    label_y_h = height_y + 6
    out.append(f'<text x="{margin + total_height/2}" y="{label_y_h}" text-anchor="middle" class="text">height = {int(total_height)} mm</text>')

    # --- Dynamic Boxes ---
    out.append('<g id="DynamicBoxes">')
    if top_seq and side_seq:
        max_top = max(top_seq)
        max_idx = next(i for i, v in enumerate(top_seq) if v == max_top)
        left_x = margin + sum(top_seq[:max_idx])
        skip = 2
        while skip < len(side_seq):
            top_y = margin + sum(side_seq[:skip])
            if skip < len(side_seq):
                height = side_seq[skip]
                out.append(f'<rect x="{left_x}" y="{top_y}" width="{max_top}" height="{height}" class="dieline"/>')
            skip += 3
    out.append('</g>')

    out.append('</svg>')
    return "\n".join(out)


# ---------------------------------------------------
# Streamlit execution
# ---------------------------------------------------

uploaded_file = st.file_uploader("Upload KLD Excel file", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None, engine="openpyxl")
    try:
        res = extract_kld_data(df)
        svg_data = make_svg(res)
        st.success("✅ Processed successfully.")
        st.download_button("⬇️ Download SVG File", svg_data, f"{uploaded_file.name}_layout.svg", "image/svg+xml")
        st.code(f"side_seq: {res['side_seq']}", language="text")
    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")
else:
    st.info("Please upload a KLD Excel file to begin.")
