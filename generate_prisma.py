# -*- coding: utf-8 -*-
"""Generate a PRISMA 2020 flow diagram (PNG) for the SLR article."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

OUT = "prisma_flow.png"

fig, ax = plt.subplots(figsize=(8.6, 9.6))
ax.set_xlim(0, 12)
ax.set_ylim(0, 13)
ax.axis("off")

FONT = "DejaVu Sans"
box_fc = "#FFFFFF"
box_ec = "#1F3B57"
side_fc = "#F0F3F6"

def box(cx, cy, w, h, text, fc=box_fc, ec=box_ec, fs=9.5, bold=False):
    patch = FancyBboxPatch((cx - w/2, cy - h/2), w, h,
                           boxstyle="round,pad=0.02,rounding_size=0.08",
                           linewidth=1.3, edgecolor=ec, facecolor=fc)
    ax.add_patch(patch)
    ax.text(cx, cy, text, ha="center", va="center", fontsize=fs,
            family=FONT, fontweight=("bold" if bold else "normal"), wrap=True)
    return (cx, cy, w, h)

def arrow(x1, y1, x2, y2):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2),
                 arrowstyle="-|>", mutation_scale=16, linewidth=1.3, color=box_ec))

# ---- stage labels (left, rotated) ----
def stage(cy, label):
    ax.add_patch(FancyBboxPatch((0.05, cy - 0.9), 0.62, 1.8,
                 boxstyle="round,pad=0.02,rounding_size=0.06",
                 linewidth=1.2, edgecolor=box_ec, facecolor="#1F3B57"))
    ax.text(0.36, cy, label, ha="center", va="center", rotation=90,
            fontsize=10.5, family=FONT, color="white", fontweight="bold")

stage(11.0, "Identification")
stage(7.0, "Screening")
stage(2.3, "Included")

main_cx = 4.7
side_cx = 9.2

# Identification
box(main_cx, 11.0, 5.2, 1.7,
    "Records identified through\ndatabase searching (n = 412)\nScopus, Web of Science, Google\nScholar, DOAJ, Crossref, Semantic\nScholar, Garuda, SINTA")
box(side_cx, 11.0, 4.0, 1.5,
    "Additional records identified\nthrough other sources\n(snowballing) (n = 18)", fc=side_fc)

# Duplicates removed
box(main_cx, 8.8, 5.2, 1.1,
    "Records after duplicates removed\n(n = 334)")

# Screening
box(main_cx, 6.9, 5.2, 1.1,
    "Records screened by\ntitle and abstract (n = 334)")
box(side_cx, 6.9, 4.0, 1.1,
    "Records excluded\n(n = 250)", fc=side_fc)

# Eligibility
box(main_cx, 4.7, 5.2, 1.2,
    "Full-text articles assessed\nfor eligibility (n = 84)")
box(side_cx, 4.6, 4.0, 2.0,
    "Full-text articles excluded,\nwith reasons (n = 76):\nnot youth/Gen Z focus (31);\nconventional-only literacy (22);\nnon-academic/inaccessible (13);\nbelow quality cut-off (10)", fc=side_fc, fs=8.5)

# Included
box(main_cx, 2.3, 5.2, 1.2,
    "Studies included in the\nsynthesis (n = 8)", bold=True)

# arrows
arrow(main_cx, 11.0 - 0.85, main_cx, 8.8 + 0.55)
arrow(side_cx, 11.0 - 0.75, main_cx + 2.6, 8.8 + 0.2)   # snowball into dedup
arrow(main_cx, 8.8 - 0.55, main_cx, 6.9 + 0.55)
arrow(main_cx, 6.9 - 0.55, main_cx, 4.7 + 0.60)
arrow(main_cx, 4.7 - 0.60, main_cx, 2.3 + 0.60)
arrow(main_cx + 2.6, 6.9, side_cx - 2.0, 6.9)
arrow(main_cx + 2.6, 4.7, side_cx - 2.0, 4.6)

plt.tight_layout()
fig.savefig(OUT, dpi=200, bbox_inches="tight", facecolor="white")
print("Saved", OUT)
