import streamlit as st
import pandas as pd
import re
import io
import html

st.set_page_config(page_title="KLD Excel → SVG Generator", layout="wide")
st.title("📏 KLD Excel → SVG Generator (Final Layout)")
st.caption("Upload Excel → Extract KLD → Download SVG (AI-editable, mm units, DIELINE colour, Arial font)")

# --------------------------------------------------------------------
# Helper functions
# --------------------------------------------------------------------
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

    return {
        "job_name": job_name,
        "width_mm": width_mm,
        "cut_length_mm": cut_length_mm,
        "top_seq": top_seq_nums,
        "side_seq": side_seq_nums,
        "photocell_w": 6,
        "photocell_h": 12,
        "photocell_offset_right_mm": 12,
        "stroke_mm": 0.25,
        "brand_label": "BRANDING",
    }

# --------------------------------------------------------------------
# SVG Generator (final layout - Illustrator-equivalent)
# --------------------------------------------------------------------
def make_svg(data):
    import html

    def parse_seq(src):
        if src is None:
            return []
        if isinstance(src, (list, tuple)):
            return [float(x) for x in src if x is not None and str(x) != ""]
        s = str(src).strip().replace(";", ",").replace("|", ",")
        parts = [p.strip() for p in s.split(",") if p.strip()]
        out = []
        for p in parts:
            try:
                out.append(float(p))
            except:
                for q in p.split():
                    try:
                        out.append(float(q))
                    except:
                        pass
        return out

    # Inputs
    W = float(data.get("cut_length_mm") or 0)
    H = float(data.get("width_mm") or 0)
    top_seq = parse_seq(data.get("top_seq"))
    side_seq = parse_seq(data.get("side_seq"))
    pcw = float(data.get("photocell_w") or 6)
    pch = float(data.get("photocell_h") or 12)
    pc_off = float(data.get("photocell_offset_right_mm") or 12)
    job = str(data.get("job_name") or "Job").replace("\n", " | ")
    brand_label = str(data.get("brand_label") or "BRANDING")

    # Style / offsets per your requirements
    dieline = "#7f00bf"
    stroke_pt = 0.356           # 0.4 pt outline thickness (as requested)
    font_pt = 3.2               # 8 pt font (universal)
    tick_short = 5.0          # tick length in mm (as before)
    top_shift_up = 5.0        # move top ticks/text up by 4 mm
    left_shift_left = 5.0     # move left ticks/text left by 4 mm
    crop_off = 5.0
    crop_len = 5.0

    # Start SVG (mm units)
    out = []
    out.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}mm" height="{H}mm" viewBox="0 0 {W} {H}">')
    out.append('<defs>')
    out.append('<style type="text/css"><![CDATA[')
    out.append(f'.dieline{{stroke:{dieline};stroke-width:{stroke_pt}pt;fill:none;}}')
    out.append(f'.dashed{{stroke:{dieline};stroke-width:{stroke_pt}pt;stroke-dasharray:1,1;fill:none;}}')
    out.append(f'.text{{font-family:Arial, Helvetica, sans-serif; font-size:{font_pt}pt; fill:{dieline};}}')
    out.append(']]></style></defs>')

    # Outer dieline
    out.append(f'<rect x="0" y="0" width="{W}" height="{H}" class="dieline"/>')

    # Vertical dashed folds from top_seq
    x = 0.0
    for v in top_seq[:-1]:
        x += v
        out.append(f'<line x1="{x}" y1="0" x2="{x}" y2="{H}" class="dashed"/>')

    # Horizontal dashed folds from side_seq
    y = 0.0
    for v in side_seq[:-1]:
        y += v
        out.append(f'<line x1="0" y1="{y}" x2="{W}" y2="{y}" class="dashed"/>')

    # Panels (optional exact panel outlines)
    xs = [0.0]; s = 0.0
    for v in top_seq:
        s += v; xs.append(s)
    ys = [0.0]; s2 = 0.0
    for v in side_seq:
        s2 += v; ys.append(s2)
    for r in range(len(side_seq)):
        y0 = ys[r]; y1 = ys[r+1]
        for c in range(len(top_seq)):
            x0 = xs[c]; x1 = xs[c+1]
            out.append(f'<rect x="{x0}" y="{y0}" width="{x1-x0}" height="{y1-y0}" class="dieline" />')

    # Photocell (same placement logic)
    pcx = W - pc_off - pcw
    pcy = H/2.0 - pch/2.0
    out.append(f'<rect x="{pcx}" y="{pcy}" width="{pcw}" height="{pch}" class="dieline"/>')
    out.append(f'<line x1="{pcx+pcw}" y1="{pcy}" x2="{pcx+pcw+3}" y2="{pcy+3}" class="dieline"/>')
    out.append(f'<line x1="{pcx+pcw}" y1="{pcy+pch}" x2="{pcx+pcw+3}" y2="{pcy+pch-3}" class="dieline"/>')

    # Crop marks — apply ±4mm shifts as requested
    out.append('<g id="CropMarks">')
    # Top-left
    # Horizontal mark (moved left by 4mm): base was from -crop_off to -(crop_off+crop_len) at y=0
    out.append(f'<line x1="{-crop_off - 4.0}" y1="0" x2="{-crop_off - crop_len - 4.0}" y2="0" class="dieline"/>')
    # Vertical mark (moved up by 4mm): base vertical was at x=0 from y=-crop_off to -(crop_off+crop_len)
    out.append(f'<line x1="0" y1="{-crop_off - 4.0}" x2="0" y2="{-crop_off - crop_len - 4.0}" class="dieline"/>')

    # Top-right
    # Horizontal mark (moved right by 4mm)
    out.append(f'<line x1="{W + crop_off + 4.0}" y1="0" x2="{W + crop_off + crop_len + 4.0}" y2="0" class="dieline"/>')
    # Vertical mark (moved up by 4mm)
    out.append(f'<line x1="{W}" y1="{-crop_off - 4.0}" x2="{W}" y2="{-crop_off - crop_len - 4.0}" class="dieline"/>')

    # Bottom-left
    # Horizontal mark (moved left by 4mm)
    out.append(f'<line x1="{-crop_off - 4.0}" y1="{H}" x2="{-crop_off - crop_len - 4.0}" y2="{H}" class="dieline"/>')
    # Vertical mark (moved down by 4mm)
    out.append(f'<line x1="0" y1="{H + crop_off + 4.0}" x2="0" y2="{H + crop_off + crop_len + 4.0}" class="dieline"/>')

    # Bottom-right
    # Horizontal mark (moved right by 4mm)
    out.append(f'<line x1="{W + crop_off + 4.0}" y1="{H}" x2="{W + crop_off + crop_len + 4.0}" y2="{H}" class="dieline"/>')
    # Vertical mark (moved down by 4mm)
    out.append(f'<line x1="{W}" y1="{H + crop_off + 4.0}" x2="{W}" y2="{H + crop_off + crop_len + 4.0}" class="dieline"/>')

    out.append('</g>')  # CropMarks

    # Measurement ticks and labels
    out.append('<g id="Measurements">')
    # TOP ticks and numeric labels (moved up by top_shift_up)
    x = 0.0
    # leftmost tick (at x=0)
    out.append(f'<line x1="{0}" y1="{-top_shift_up}" x2="{0}" y2="{-top_shift_up - tick_short}" class="dieline"/>')
    for v in top_seq:
        x += v
        out.append(f'<line x1="{x}" y1="{-top_shift_up}" x2="{x}" y2="{-top_shift_up - tick_short}" class="dieline"/>')
        mid = x - v/2.0
        out.append(f'<text x="{mid}" y="{-top_shift_up - tick_short - 1.0}" text-anchor="middle" class="text">{int(round(v))}</text>')

    # LEFT ticks and numeric labels (moved left by left_shift_left)
    y = 0.0
    # topmost left tick at y=0
    out.append(f'<line x1="{-left_shift_left}" y1="{0}" x2="{-left_shift_left - tick_short}" y2="{0}" class="dieline"/>')
    for v in side_seq:
        y += v
        out.append(f'<line x1="{-left_shift_left}" y1="{y}" x2="{-left_shift_left - tick_short}" y2="{y}" class="dieline"/>')
        midY = y - v/2.0
        # rotate -90 to match Illustrator style and place left by left_shift + tick + small gap
        label_x = -left_shift_left - tick_short - 2.0
        out.append(f'<text x="{label_x}" y="{midY}" transform="rotate(-90 {label_x} {midY})" text-anchor="middle" class="text">{int(round(v))}</text>')

    out.append('</g>')  # Measurements

    # Centre seal labels and Branding and END SEAL (kept similar)
    # Centre seal top/bottom
    out.append(f'<text x="{W/2}" y="{5}" text-anchor="middle" class="text" font-weight="bold">CENTRE SEAL AREA</text>')
    out.append(f'<text x="{W/2}" y="{H-2}" text-anchor="middle" class="text" font-weight="bold">CENTRE SEAL AREA</text>')
    # Branding
    total_side = sum(side_seq) if side_seq else H
    branding_y = total_side/2.0
    out.append(f'<text x="{W/2}" y="{branding_y}" text-anchor="middle" class="text" font-weight="bold">{html.escape(brand_label)}</text>')
    # END SEAL left/right rotated
    midy = H/2.0
    out.append(f'<text x="{-18}" y="{midy}" transform="rotate(-90 {-18} {midy})" text-anchor="middle" class="text" font-weight="bold">END SEAL</text>')
    out.append(f'<text x="{W+18}" y="{midy}" transform="rotate(-90 {W+18} {midy})" text-anchor="middle" class="text" font-weight="bold">END SEAL</text>')

    out.append('</svg>')
    return "\n".join(out)




# --------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------
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
