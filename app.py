    # --- Crop Marks (all corners, full directional set) ---
    out.append('<g id="CropMarks">')

    # --- TOP-LEFT ---
    # Horizontal (→ left)
    out.append(f'<line x1="0" y1="{H}" x2="{-crop_off - crop_len}" y2="{H}" class="dieline"/>')
    # Vertical (↑ up)
    out.append(f'<line x1="0" y1="{H}" x2="0" y2="{H + crop_off + crop_len}" class="dieline"/>')

    # --- TOP-RIGHT ---
    # Horizontal (→ right)
    out.append(f'<line x1="{W + crop_off}" y1="{H}" x2="{W + crop_off + crop_len}" y2="{H}" class="dieline"/>')
    # Vertical (↑ up)
    out.append(f'<line x1="{W}" y1="{H}" x2="{W}" y2="{H + crop_off + crop_len}" class="dieline"/>')

    # --- BOTTOM-LEFT ---
    # Horizontal (→ left)
    out.append(f'<line x1="0" y1="0" x2="{-crop_off - crop_len}" y2="0" class="dieline"/>')
    # Vertical (↓ down)
    out.append(f'<line x1="0" y1="0" x2="0" y2="{-crop_off - crop_len}" class="dieline"/>')

    # --- BOTTOM-RIGHT ---
    # Horizontal (→ right)
    out.append(f'<line x1="{W + crop_off}" y1="0" x2="{W + crop_off + crop_len}" y2="0" class="dieline"/>')
    # Vertical (↓ down)
    out.append(f'<line x1="{W}" y1="0" x2="{W}" y2="{-crop_off - crop_len}" class="dieline"/>')

    out.append('</g>')
