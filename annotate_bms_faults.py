"""
Annotate BMS fault list images with English translations.
Adds translations to the right of each Chinese fault name row.
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

from PIL import Image, ImageDraw, ImageFont
import numpy as np

IMG1_SRC = r"C:\Users\JasonOng\xwechat_files\wxid_g9xziacvqskj22_a87e\temp\RWTemp\2026-07\96a931d15394eb088f88fa5c60259c55.png"
IMG2_SRC = r"C:\Users\JasonOng\OneDrive - Advancer Global Ltd\AST BD\HyESys Dept\2. Software (EMS.BMS)\LEMS.BMS\V2.1 AST OFC.Xiasha\EMS.BMS shutdown faults.png"
IMG1_OUT = r"C:\Users\JasonOng\OneDrive - Advancer Global Ltd\AST BD\HyESys Dept\2. Software (EMS.BMS)\LEMS.BMS\V2.1 AST OFC.Xiasha\96a931_annotated_EN.png"
IMG2_OUT = r"C:\Users\JasonOng\OneDrive - Advancer Global Ltd\AST BD\HyESys Dept\2. Software (EMS.BMS)\LEMS.BMS\V2.1 AST OFC.Xiasha\EMS.BMS shutdown faults_annotated_EN.png"

FONT_PATH = r"C:\Windows\Fonts\arial.ttf"

# ── Translations for Image #1 (36 fault lines) ─────────────────────────────
TRANS_IMG1 = [
    "Level 3 Fault:",
    "Ambient Over-Temperature",
    "Cell Core Over-Voltage",
    "Charging Over-Current",
    "Charging Over-Temperature",
    "Charging Under-Temperature",
    "DC Internal Resistance Deviation Too Large",
    "DC Internal Resistance Too Large",
    "HV Box Terminal Post Over-Temperature",
    "Pre-Charge Fault",
    "HV Acquisition Fault",
    "Insulation Fault",
    "Battery Total Voltage Over-Voltage",
    "Pack Terminal Post Over-Temperature",
    "Cell Temperature Difference Too Large",
    "Temperature Sensor Wire Break",
    "Cell Voltage Difference Too Large",
    "Cell Voltage Sensor Wire Break",
    "6815 Communication Fault",
    "Main Neg. Relay Coil Short to Power Supply",
    "Main Neg. Relay Coil Short to Ground",
    "Main Neg. Relay Coil Open Circuit",
    "Main Neg. Relay Closing Fault: Unable to Close",
    "Main Neg. Relay Opening Fault: Welded",
    "Main Pos. Fuse Fault",
    "Pre-Charge Relay Coil Short to Power Supply",
    "Pre-Charge Relay Coil Short to Ground",
    "Pre-Charge Relay Coil Open Circuit",
    "Pre-Charge Relay Closing Fault: Unable to Close",
    "Pre-Charge Relay Opening Fault: Welded",
    "Main Pos. Relay Coil Short to Power Supply",
    "Main Pos. Relay Coil Short to Ground",
    "Main Pos. Relay Coil Open Circuit",
    "Main Pos. Relay Closing Fault: Unable to Close",
    "Main Pos. Relay Opening Fault: Relay Closing Fault Welded",
    "Fire Suppression",
]

# ── Translations for Image #2 (35 fault lines — no separate 直流内阻过大) ──
TRANS_IMG2 = [
    "Level 3 Fault:",
    "Ambient Over-Temperature:",
    "Cell Core Over-Voltage:",
    "Charging Over-Current:",
    "Charging Over-Temperature:",
    "Charging Under-Temperature:",
    "DC Internal Resistance Deviation Too Large:",
    "HV Box Terminal Post Over-Temperature:",
    "Pre-Charge Fault",
    "HV Acquisition Fault",
    "Insulation Resistance Fault:",
    "Battery Total Voltage Over-Voltage:",
    "Pack Terminal Post Over-Temperature:",
    "Cell Temperature Difference Too Large:",
    "Temperature Sensor Wire Break",
    "Cell Voltage Difference Too Large:",
    "Cell Voltage Sensor Wire Break",
    "6815 Communication Fault",
    "Main Neg. Relay Coil Short to Power Supply",
    "Main Neg. Relay Coil Short to Ground",
    "Main Neg. Relay Coil Open Circuit",
    "Main Neg. Relay Closing Fault: Unable to Close",
    "Main Neg. Relay Opening Fault: Welded",
    "Main Pos. Fuse Fault",
    "Pre-Charge Relay Coil Short to Power Supply",
    "Pre-Charge Relay Coil Short to Ground",
    "Pre-Charge Relay Coil Open Circuit",
    "Pre-Charge Relay Closing Fault: Unable to Close",
    "Pre-Charge Relay Opening Fault: Welded",
    "Main Pos. Relay Coil Short to Power Supply",
    "Main Pos. Relay Coil Short to Ground",
    "Main Pos. Relay Coil Open Circuit",
    "Main Pos. Relay Closing Fault: Unable to Close",
    "Main Pos. Relay Opening Fault: Relay Closing Fault Welded",
    "Fire Suppression",
]


def find_text_lines(img, threshold=100, min_pixels=3, min_gap=5):
    """Return list of (y_center) for each detected text line."""
    arr = np.array(img.convert('L'))
    row_dark = np.sum(arr < threshold, axis=1)
    lines, in_line, group_start, prev = [], False, 0, -99
    for y, count in enumerate(row_dark):
        if count >= min_pixels and not in_line:
            in_line, group_start = True, y
        elif count < min_pixels and in_line:
            in_line = False
            if y - group_start >= 1:
                lines.append((group_start + y) // 2)
        prev = y
    # Filter out toolbar line (last line is usually far below content)
    if len(lines) >= 2:
        gap = lines[-1] - lines[-2]
        if gap > 100:
            lines = lines[:-1]
    return lines


def find_max_right_edge(img, line_y_positions, half_h=12, threshold=100):
    """Find the rightmost dark pixel x across all text line bands."""
    arr = np.array(img.convert('L'))
    h, w = arr.shape
    max_x = 0
    for y in line_y_positions:
        y0, y1 = max(0, y - half_h), min(h, y + half_h)
        band = arr[y0:y1, :]
        dark_cols = np.where(np.any(band < threshold, axis=0))[0]
        if len(dark_cols):
            max_x = max(max_x, int(dark_cols[-1]))
    return max_x


def annotate(src_path, out_path, translations, font_size=15):
    img = Image.open(src_path).convert('RGB')
    orig_w, orig_h = img.size

    lines = find_text_lines(img)
    n_content = min(len(lines), len(translations))
    print(f"  Detected {len(lines)} lines, using {n_content} translations")

    # Determine where English text starts
    right_edge = find_max_right_edge(img, lines[:n_content])
    en_x = right_edge + 30

    # Estimate width needed for English text
    font = ImageFont.truetype(FONT_PATH, font_size)
    dummy = Image.new('RGB', (1, 1))
    dummy_draw = ImageDraw.Draw(dummy)
    max_en_w = max(dummy_draw.textlength(t, font=font) for t in translations[:n_content])

    needed_w = int(en_x + max_en_w + 20)
    canvas_w = max(orig_w, needed_w)

    # Create new canvas (white background)
    canvas = Image.new('RGB', (canvas_w, orig_h), (255, 255, 255))
    canvas.paste(img, (0, 0))

    draw = ImageDraw.Draw(canvas)

    for i, (y, en_text) in enumerate(zip(lines[:n_content], translations[:n_content])):
        # Align text baseline to match the Chinese text row
        bbox = font.getbbox(en_text)
        text_h = bbox[3] - bbox[1]
        ty = y - text_h // 2

        # Draw a subtle grey separator line at the en_x boundary (first line only)
        if i == 0:
            draw.line([(en_x - 15, 15), (en_x - 15, lines[n_content - 1] + 14)],
                      fill=(200, 200, 200), width=1)

        draw.text((en_x, ty), en_text, font=font, fill=(0, 100, 180))

    canvas.save(out_path, 'PNG')
    print(f"  Saved → {out_path}")


print("Annotating Image #1...")
annotate(IMG1_SRC, IMG1_OUT, TRANS_IMG1, font_size=15)

print("Annotating Image #2...")
annotate(IMG2_SRC, IMG2_OUT, TRANS_IMG2, font_size=15)

print("Done.")
