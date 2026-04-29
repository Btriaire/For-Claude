import io
import os
import tempfile
from datetime import datetime
from typing import Dict, List, Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


def _rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _format_value(value: float, fmt: str) -> str:
    if fmt == "currency":
        if value >= 1_000_000:
            return f"${value / 1_000_000:.1f}M"
        if value >= 1_000:
            return f"${value / 1_000:.1f}K"
        return f"${value:,.0f}"
    if fmt == "percentage":
        return f"{value:.1f}%"
    if value >= 1_000_000:
        return f"{value / 1_000_000:.1f}M"
    if value >= 1_000:
        return f"{value / 1_000:.1f}K"
    return f"{value:,.0f}"


CHART_PALETTE = [
    "#5B2D8E", "#C9A227", "#6DBF8B", "#E8735A",
    "#9B59B6", "#D4A017", "#52BE80", "#E07B5A",
]

CARD_COLORS = ["#5B2D8E", "#C9A227", "#6DBF8B", "#E8735A",
               "#7D3C98", "#B7860B", "#58D68D", "#D35400"]


class PPTXExporter:
    W = Inches(13.33)
    H = Inches(7.5)

    def __init__(self, config: Dict):
        self.config = config
        colors = config.get("colors", {})
        self.purple = colors.get("primary", "#5B2D8E")
        self.yellow = colors.get("secondary", "#C9A227")
        self.green = colors.get("tertiary", "#6DBF8B")
        self.orange = colors.get("quaternary", "#E8735A")
        self.palette = [self.purple, self.yellow, self.green, self.orange,
                        "#9B59B6", "#D4A017", "#52BE80", "#E07B5A"]

    def generate(self, data: List[Dict], period: str,
                 comments: str, key_events: str) -> str:
        prs = Presentation()
        prs.slide_width = self.W
        prs.slide_height = self.H

        self._title_slide(prs, period)
        if data:
            self._kpi_overview_slide(prs, data, period)

        chart_kpis = [d for d in data if d["breakdown"]]
        for i in range(0, len(chart_kpis), 2):
            self._charts_slide(prs, chart_kpis[i: i + 2], period)

        if key_events or comments:
            self._events_slide(prs, key_events, comments, period)

        out = tempfile.mktemp(suffix=".pptx")
        prs.save(out)
        return out

    # ── helpers ──────────────────────────────────────────────────────────────

    def _blank_slide(self, prs: Presentation):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = _rgb("#FFFFFF")
        return slide

    def _rect(self, slide, x, y, w, h, fill_hex, line=False):
        shape = slide.shapes.add_shape(1, x, y, w, h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill_hex)
        if line:
            shape.line.color.rgb = _rgb(fill_hex)
        else:
            shape.line.fill.background()
        return shape

    def _textbox(self, slide, x, y, w, h, text, size, bold=False,
                 color="#1A1A1A", align=PP_ALIGN.LEFT, font="Source Sans Pro"):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = _rgb(color)
        return tb

    def _header(self, slide, prs, title: str, period: str):
        self._rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.9), self.purple)
        self._rect(slide, Inches(0), Inches(0.9), prs.slide_width, Inches(0.06), self.yellow)
        self._textbox(slide, Inches(0.35), Inches(0.1), Inches(9), Inches(0.75),
                      title, 22, bold=True, color="#FFFFFF")
        self._textbox(slide, Inches(9.8), Inches(0.2), Inches(3.2), Inches(0.5),
                      period, 14, color="#FFFFFF", align=PP_ALIGN.RIGHT)

    # ── slides ────────────────────────────────────────────────────────────────

    def _title_slide(self, prs: Presentation, period: str):
        slide = self._blank_slide(prs)
        self._rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.15), self.purple)
        self._rect(slide, Inches(0), prs.slide_height - Inches(0.15), prs.slide_width, Inches(0.15), self.purple)
        self._rect(slide, Inches(0.5), Inches(2.3), Inches(0.12), Inches(3.0), self.purple)
        self._rect(slide, Inches(0.5), Inches(5.3), Inches(6), Inches(0.06), self.yellow)

        self._textbox(slide, Inches(0.8), Inches(2.5), Inches(10), Inches(1.3),
                      "Monthly Performance Dashboard", 40, bold=True, color="#1A1A1A")
        self._textbox(slide, Inches(0.8), Inches(3.85), Inches(8), Inches(0.65),
                      period, 26, color=self.purple)
        self._textbox(slide, Inches(0.8), Inches(4.7), Inches(8), Inches(0.45),
                      f"Generated on {datetime.now().strftime('%B %d, %Y')}", 13, color="#888888")

    def _kpi_overview_slide(self, prs: Presentation, data: List[Dict], period: str):
        slide = self._blank_slide(prs)
        self._header(slide, prs, "KPI Overview", period)

        card_w = Inches(2.8)
        card_h = Inches(1.55)
        gap = Inches(0.28)
        start_x = Inches(0.45)
        start_y = Inches(1.25)

        for i, kpi in enumerate(data[:8]):
            row, col = divmod(i, 4)
            x = start_x + col * (card_w + gap)
            y = start_y + row * (card_h + gap)
            color = self.palette[i % len(self.palette)]

            self._rect(slide, x, y, card_w, card_h, color)
            self._textbox(slide, x + Inches(0.15), y + Inches(0.12),
                          card_w - Inches(0.3), Inches(0.38),
                          kpi["label"].upper(), 10, bold=True, color="#FFFFFF")
            self._textbox(slide, x + Inches(0.15), y + Inches(0.52),
                          card_w - Inches(0.3), Inches(0.85),
                          _format_value(kpi["total"], kpi["format"]),
                          28, bold=True, color="#FFFFFF")

    def _charts_slide(self, prs: Presentation, kpis: List[Dict], period: str):
        slide = self._blank_slide(prs)
        title = "  ·  ".join(k["label"] for k in kpis)
        self._header(slide, prs, title, period)

        positions = [(Inches(0.35), Inches(1.05)), (Inches(6.85), Inches(1.05))]
        for i, kpi in enumerate(kpis[:2]):
            x, y = positions[i]
            img = self._donut_image(kpi)
            if img:
                slide.shapes.add_picture(img, x, y, Inches(6.0), Inches(5.9))

    def _events_slide(self, prs: Presentation, key_events: str, comments: str, period: str):
        slide = self._blank_slide(prs)
        self._header(slide, prs, "Key Events & Commentary", period)

        y = Inches(1.15)
        if key_events:
            self._textbox(slide, Inches(0.5), y, Inches(12.3), Inches(0.4),
                          "KEY EVENTS", 13, bold=True, color=self.purple)
            y += Inches(0.45)
            self._rect(slide, Inches(0.5), y, Inches(0.06), Inches(2.0), self.yellow)
            self._textbox(slide, Inches(0.75), y, Inches(12.0), Inches(2.0),
                          key_events, 13, color="#1A1A1A")
            y += Inches(2.3)

        if comments:
            self._textbox(slide, Inches(0.5), y, Inches(12.3), Inches(0.4),
                          "COMMENTARY", 13, bold=True, color=self.purple)
            y += Inches(0.45)
            self._rect(slide, Inches(0.5), y, Inches(0.06), Inches(1.8), self.green)
            self._textbox(slide, Inches(0.75), y, Inches(12.0), Inches(1.8),
                          comments, 13, color="#1A1A1A")

    # ── chart image ───────────────────────────────────────────────────────────

    def _donut_image(self, kpi: Dict) -> Optional[io.BytesIO]:
        breakdown = kpi["breakdown"]
        if not breakdown:
            return None

        labels = [b["label"] for b in breakdown]
        values = [b["value"] for b in breakdown]
        total = sum(values)
        colors_hex = self.palette[: len(values)]

        fig, (ax_d, ax_l) = plt.subplots(
            1, 2, figsize=(9.6, 5.5),
            gridspec_kw={"width_ratios": [1, 0.8]},
        )
        fig.patch.set_facecolor("white")

        wedges, _ = ax_d.pie(
            values,
            colors=colors_hex,
            startangle=90,
            wedgeprops=dict(width=0.55, edgecolor="white", linewidth=2.5),
        )
        ax_d.set_aspect("equal")
        ax_d.axis("off")

        # Center text
        ax_d.text(0, 0.10, _format_value(total, kpi.get("format", "number")),
                  ha="center", va="center", fontsize=17, fontweight="bold",
                  color="#1A1A1A", fontfamily="sans-serif")
        ax_d.text(0, -0.18, kpi["label"],
                  ha="center", va="center", fontsize=10.5,
                  color="#666666", fontfamily="sans-serif")

        # Legend panel
        ax_l.axis("off")
        legend_y = 1.0
        line_h = min(0.145, 0.9 / len(labels))

        for label, value, color in zip(labels, values, colors_hex):
            pct = value / total * 100 if total else 0
            rect = mpatches.FancyBboxPatch(
                (0.0, legend_y - 0.055), 0.065, 0.06,
                boxstyle="round,pad=0.005",
                facecolor=color,
                transform=ax_l.transAxes,
                clip_on=False,
            )
            ax_l.add_patch(rect)
            ax_l.text(0.10, legend_y - 0.016, label,
                      transform=ax_l.transAxes, fontsize=9,
                      color="#1A1A1A", va="center", fontfamily="sans-serif")
            ax_l.text(0.10, legend_y - 0.075,
                      f"{_format_value(value, kpi.get('format', 'number'))}  ({pct:.1f}%)",
                      transform=ax_l.transAxes, fontsize=7.5,
                      color="#666666", va="center", fontfamily="sans-serif")
            legend_y -= line_h * 1.8

        plt.subplots_adjust(left=0, right=1, top=1, bottom=0, wspace=0.02)

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                    facecolor="white", edgecolor="none")
        buf.seek(0)
        plt.close(fig)
        return buf
