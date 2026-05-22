"""
Draws the HyESys Multi-Site Agent Architecture flowchart.
Saves as PNG to the specified OneDrive path.
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np

DST = (
    r"C:\Users\JasonOng\OneDrive - Advancer Global Ltd"
    r"\AST BD\2024 HyESys\Software\claude code"
    r"\HyESys_MultiSite_Agent_Architecture.png"
)

# ── Colours ───────────────────────────────────────────────────────────────────
C_COORD   = "#1F3763"   # dark navy   — coordinator
C_POOL    = "#2E75B6"   # mid blue    — process pool
C_SITE    = "#375623"   # dark green  — site process boxes
C_SITE_LT = "#E2EFDA"   # light green — site box fill
C_A1_HDR  = "#843C0C"   # dark orange — Agent 1 header
C_A1      = "#F4B183"   # light orange
C_A2_HDR  = "#1F3763"   # dark navy   — Agent 2 header
C_A2      = "#BDD7EE"   # light blue
C_DB      = "#7030A0"   # purple      — database
C_RETRAIN = "#C00000"   # red         — retrain
C_SEP     = "#D9D9D9"   # separator
C_WHITE   = "#FFFFFF"
C_TEXT_DK = "#1A1A1A"
C_ARROW   = "#404040"
C_DIAMOND = "#ED7D31"   # orange — decision diamonds


# ── Helpers ───────────────────────────────────────────────────────────────────

def rect(ax, cx, cy, w, h, label, fc, ec=C_WHITE, fontsize=8.5,
         tc=C_WHITE, bold=False, radius=0.3, alpha=1.0, zorder=3):
    patch = FancyBboxPatch(
        (cx - w / 2, cy - h / 2), w, h,
        boxstyle=f"round,pad=0",
        facecolor=fc, edgecolor=ec,
        linewidth=1.5, zorder=zorder, alpha=alpha
    )
    ax.add_patch(patch)
    ax.text(cx, cy, label,
            ha="center", va="center",
            fontsize=fontsize, color=tc,
            fontweight="bold" if bold else "normal",
            zorder=zorder + 1,
            multialignment="center",
            linespacing=1.4)


def diamond_shape(ax, cx, cy, w, h, label, fc, fontsize=8, tc=C_WHITE):
    pts = np.array([
        [cx,       cy + h/2],
        [cx + w/2, cy],
        [cx,       cy - h/2],
        [cx - w/2, cy]
    ])
    patch = plt.Polygon(pts, closed=True, facecolor=fc,
                        edgecolor=C_WHITE, linewidth=1.5, zorder=3)
    ax.add_patch(patch)
    ax.text(cx, cy, label, ha="center", va="center",
            fontsize=fontsize, color=tc, fontweight="bold",
            zorder=4, multialignment="center", linespacing=1.3)


def arrow_v(ax, cx, y1, y2, color=C_ARROW, lw=1.5, style="->"):
    ax.annotate("", xy=(cx, y2), xytext=(cx, y1),
                arrowprops=dict(arrowstyle=style, color=color,
                                lw=lw, mutation_scale=14),
                zorder=2)


def arrow_h(ax, x1, cy, x2, color=C_ARROW, lw=1.5):
    ax.annotate("", xy=(x2, cy), xytext=(x1, cy),
                arrowprops=dict(arrowstyle="->", color=color,
                                lw=lw, mutation_scale=14),
                zorder=2)


def arrow_custom(ax, x1, y1, x2, y2, color=C_ARROW, lw=1.5,
                 connectionstyle="arc3,rad=0.0"):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="->", color=color, lw=lw,
                                mutation_scale=14,
                                connectionstyle=connectionstyle),
                zorder=2)


def label(ax, x, y, text, fontsize=8, color=C_TEXT_DK, ha="center",
          bold=False, style="normal"):
    ax.text(x, y, text, ha=ha, va="center",
            fontsize=fontsize, color=color,
            fontweight="bold" if bold else "normal",
            fontstyle=style, zorder=5)


def hline(ax, y, x0=0.5, x1=21.5, color=C_SEP, lw=1, ls="--"):
    ax.plot([x0, x1], [y, y], color=color, lw=lw, ls=ls, zorder=1)


# ── Build figure ──────────────────────────────────────────────────────────────

fig, ax = plt.subplots(figsize=(22, 17))
ax.set_xlim(0, 22)
ax.set_ylim(0, 17)
ax.axis("off")
fig.patch.set_facecolor("#FAFAFA")

# Background panels
ax.add_patch(FancyBboxPatch((0.3, 8.8), 21.4, 7.9,
             boxstyle="round,pad=0", facecolor="#EEF4FB",
             edgecolor=C_SEP, linewidth=1, zorder=0))
ax.add_patch(FancyBboxPatch((0.3, 0.2), 21.4, 8.4,
             boxstyle="round,pad=0", facecolor="#FFF8F0",
             edgecolor=C_SEP, linewidth=1, zorder=0))

# Panel labels
label(ax, 1.6, 16.4, "SYSTEM LEVEL", fontsize=8.5,
      color=C_POOL, bold=True)
label(ax, 1.6, 8.6, "PER-SITE PROCESS  (Detail View)", fontsize=8.5,
      color=C_A1_HDR, bold=True)

# ── TITLE ─────────────────────────────────────────────────────────────────────
ax.text(11, 16.55,
        "HyESys  Multi-Site Agent Architecture",
        ha="center", va="center", fontsize=16,
        fontweight="bold", color=C_COORD, zorder=5)
ax.text(11, 16.1,
        "Concurrent Processing  ·  Per-Site Config  ·  Per-Site Database  ·  Staggered Retraining",
        ha="center", va="center", fontsize=9,
        color="#555555", style="italic", zorder=5)

# ── SYSTEM LEVEL ──────────────────────────────────────────────────────────────

# Coordinator
rect(ax, 11, 15.35, 5.5, 0.75,
     "COORDINATOR\nmain.py", fc=C_COORD, fontsize=9, bold=True)
arrow_v(ax, 11, 14.97, 14.67)

# Load configs
rect(ax, 11, 14.3, 5.5, 0.7,
     "Load All Site Configs\n( config/*.yaml )", fc=C_POOL, fontsize=9)
arrow_v(ax, 11, 13.95, 13.65)

# Process Pool
rect(ax, 11, 13.2, 7.0, 0.75,
     "Process Pool Executor\n( max workers = CPU core count )",
     fc=C_POOL, fontsize=9, bold=True)

# Fan-out arrows from pool to site boxes
site_xs = [2.5, 6.5, 11.0, 15.5, 19.5]
for sx in site_xs:
    # Vertical line down from pool
    ax.plot([11, sx], [12.82, 12.0],
            color=C_ARROW, lw=1.3, zorder=2)
    ax.annotate("", xy=(sx, 11.6), xytext=(sx, 12.0),
                arrowprops=dict(arrowstyle="->", color=C_ARROW,
                                lw=1.3, mutation_scale=12), zorder=2)

# "..." between site 3 and 4
ax.text(13.25, 11.2, "···", ha="center", va="center",
        fontsize=16, color=C_ARROW, fontweight="bold")

# Site process boxes
site_labels = [
    "Site 1\nAgent 1 + Agent 2\nconfig/site1.yaml\ndata/site1.db",
    "Site 2\nAgent 1 + Agent 2\nconfig/site2.yaml\ndata/site2.db",
    "Site 3\nAgent 1 + Agent 2\nconfig/site3.yaml\ndata/site3.db",
    "Site N\nAgent 1 + Agent 2\nconfig/siteN.yaml\ndata/siteN.db",
]
for i, (sx, lbl) in enumerate(zip([2.5, 6.5, 11.0, 19.5], site_labels)):
    rect(ax, sx, 10.85, 3.4, 1.5, lbl,
         fc=C_SITE, ec="#A9D18E", fontsize=7.5, zorder=3)

# Dashed arrows from site boxes down to detail section
for sx in [2.5, 6.5, 11.0, 19.5]:
    ax.annotate("", xy=(sx, 8.85), xytext=(sx, 10.1),
                arrowprops=dict(arrowstyle="->", color="#70AD47",
                                lw=1.2, linestyle="dashed",
                                mutation_scale=11), zorder=2)

# Horizontal separator
hline(ax, 8.8, lw=1.5, ls="-", color="#AAAAAA")

# ── PER-SITE DETAIL ───────────────────────────────────────────────────────────

# -- AGENT 1 column (left) --
A1X = 5.5
rect(ax, A1X, 8.2, 4.2, 0.7,
     "AGENT 1  —  Data Quality & Ingestion",
     fc=C_A1_HDR, fontsize=9, bold=True)
arrow_v(ax, A1X, 7.85, 7.55)

rect(ax, A1X, 7.2, 4.0, 0.6,
     "Receive Meter Data\n( CSV  /  Live MQTT feed )",
     fc=C_A1, tc=C_TEXT_DK, fontsize=8.5)
arrow_v(ax, A1X, 6.9, 6.6)

diamond_shape(ax, A1X, 6.15, 3.8, 0.75,
              "Validate Row\n( Agent 1 rules\nfrom site config )",
              fc=C_DIAMOND, fontsize=8)
arrow_v(ax, A1X, 5.77, 5.47)

rect(ax, A1X, 5.1, 4.0, 0.6,
     "Tag Each Row:\nCLEAN  /  SUSPECT  /  REJECTED",
     fc=C_A1, tc=C_TEXT_DK, fontsize=8.5)
arrow_v(ax, A1X, 4.8, 4.5)

rect(ax, A1X, 4.15, 4.0, 0.6,
     "Write CLEAN rows to site.db",
     fc=C_A1, tc=C_TEXT_DK, fontsize=8.5)

# REJECTED branch label
label(ax, A1X + 2.3, 5.95, "CLEAN / SUSPECT",
      fontsize=7.5, color=C_A1_HDR, ha="left")
label(ax, A1X - 2.3, 5.95, "REJECTED → discard",
      fontsize=7.5, color="#C00000", ha="right")
ax.annotate("", xy=(A1X - 3.6, 5.77), xytext=(A1X - 1.9, 6.15),
            arrowprops=dict(arrowstyle="->", color="#C00000",
                            lw=1.2, mutation_scale=11), zorder=2)

# -- AGENT 2 column (right) --
A2X = 16.5
rect(ax, A2X, 8.2, 4.2, 0.7,
     "AGENT 2  —  Analysis & Recommendation",
     fc=C_A2_HDR, fontsize=9, bold=True)
arrow_v(ax, A2X, 7.85, 7.55)

rect(ax, A2X, 7.2, 4.0, 0.6,
     "Read 15-min State Snapshots\nfrom site.db",
     fc=C_A2, tc=C_TEXT_DK, fontsize=8.5)
arrow_v(ax, A2X, 6.9, 6.6)

rect(ax, A2X, 6.25, 4.0, 0.6,
     "Build State Snapshot\n( kW, kVAr, PF, Voltage, Timestamp )",
     fc=C_A2, tc=C_TEXT_DK, fontsize=8.5)
arrow_v(ax, A2X, 5.95, 5.65)

diamond_shape(ax, A2X, 5.15, 3.8, 0.75,
              "Decision Engine\n( equations + thresholds\nfrom site config )",
              fc=C_POOL, fontsize=8)
arrow_v(ax, A2X, 4.77, 4.47)

rect(ax, A2X, 4.1, 4.0, 0.6,
     "Issue Action:\nINJECT_KVAR  /  HOLD  /  REDUCE",
     fc=C_A2, tc=C_TEXT_DK, fontsize=8.5)
arrow_v(ax, A2X, 3.8, 3.5)

rect(ax, A2X, 3.15, 4.0, 0.6,
     "Log SAR Triplet\n( State → Action → Reward )",
     fc=C_A2, tc=C_TEXT_DK, fontsize=8.5)

# -- Shared DB (centre bottom) --
rect(ax, 11, 2.5, 5.2, 0.85,
     "Per-Site SQLite Database\ndata/{ site_id }.db",
     fc=C_DB, fontsize=9, bold=True)

# Agent 1 write → DB
arrow_custom(ax, A1X, 3.85, 8.9, 2.5,
             connectionstyle="arc3,rad=-0.25", color=C_DB, lw=1.5)
label(ax, 7.5, 3.1, "write\nCLEAN rows", fontsize=7.5,
      color=C_DB, bold=False)

# DB → Agent 2 read
arrow_custom(ax, 13.1, 2.5, A2X, 6.9,
             connectionstyle="arc3,rad=-0.25", color=C_DB, lw=1.5)
label(ax, 15.5, 4.5, "read\nsnapshots", fontsize=7.5,
      color=C_DB, bold=False)

# Agent 2 SAR log → DB
arrow_custom(ax, A2X, 2.85, 13.5, 2.5,
             connectionstyle="arc3,rad=0.2", color=C_DB, lw=1.5)
label(ax, 15.7, 2.3, "log SAR", fontsize=7.5, color=C_DB)

# -- Staggered Retrain --
arrow_v(ax, 11, 2.07, 1.55)
rect(ax, 11, 1.15, 6.0, 0.7,
     "Staggered Nightly Retrain\n( offset per site  ·  ~5 sec each  ·  model.pkl updated )",
     fc=C_RETRAIN, fontsize=8.5, bold=True)

# -- Side legend --
legend_items = [
    (C_COORD,   "Coordinator / Scheduler"),
    (C_POOL,    "Process Pool / Agent 2 Decision"),
    (C_SITE,    "Per-Site Worker Process"),
    (C_A1_HDR,  "Agent 1  (Data Validation)"),
    (C_A2_HDR,  "Agent 2  (Analysis)"),
    (C_DIAMOND, "Decision / Validation Step"),
    (C_DB,      "Per-Site SQLite Database"),
    (C_RETRAIN, "Nightly Retrain"),
]
lx, ly = 0.65, 7.6
ax.text(lx, ly + 0.35, "Legend", fontsize=8, fontweight="bold",
        color=C_TEXT_DK, va="center")
for i, (color, name) in enumerate(legend_items):
    yy = ly - 0.4 - i * 0.42
    patch = FancyBboxPatch((lx, yy - 0.14), 0.45, 0.28,
                           boxstyle="round,pad=0",
                           facecolor=color, edgecolor="none", zorder=4)
    ax.add_patch(patch)
    ax.text(lx + 0.6, yy, name, fontsize=7.5, va="center",
            color=C_TEXT_DK, zorder=5)

# Footer
ax.text(11, 0.12,
        "Advancer Smart Technology Pte Ltd  ·  HyESys Department  ·  May 2026  ·  Confidential",
        ha="center", va="center", fontsize=7.5,
        color="#888888", style="italic")

plt.tight_layout(pad=0.3)
plt.savefig(DST, dpi=180, bbox_inches="tight", facecolor="#FAFAFA")
plt.close()
print(f"Saved: {DST}")
