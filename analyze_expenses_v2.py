#!/usr/bin/env python3
"""
Alex expense analysis (KES) + merchant/location enrichment.

Usage:
  python analyze_expenses_v2.py --input "Alex expences.xlsx" --outdir "alex_expense_outputs_v2"

Outputs:
  - charts (*.png)
  - summary JSON (alex_expense_summary_v2.json)
  - PDF report (Alex_Expense_Report_MerchantLocation_v2.pdf)

Notes:
  - Merchant/Area are heuristically derived from the Location text.
  - Amounts are treated as Kenyan Shillings (KSh).
"""
import argparse, os, re, json
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

from reportlab.lib.pagesizes import LETTER
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch


def parse_location(loc):
    """Return (merchant, area) derived from a free-text Location string."""
    if pd.isna(loc):
        return ("Unknown", "Unknown")
    s = re.sub(r"\s+", " ", str(loc).strip())
    lower = s.lower()

    # Common brand rules
    if lower.startswith("shell"):
        merchant = "Shell"
        rest = s[5:].strip(" ,-_")
        return merchant, (rest if rest else "Unknown")

    if lower.startswith("total"):
        merchant = "Total"
        rest = s[5:].strip(" ,-_")
        return merchant, (rest if rest else "Unknown")

    if lower.startswith("home"):
        merchant = "Home"
        parts = re.split(r"[_-]", s, maxsplit=1)
        area = parts[1].strip() if len(parts) > 1 else s[4:].strip(" ,-_")
        return merchant, (area if area else "Unknown")

    if lower.startswith("love dale butchery") or lower.startswith("love dale"):
        merchant = "Love Dale Butchery"
        if "_" in s:
            area = s.split("_", 1)[1].strip()
        else:
            area = s[len("Love Dale Butchery"):].strip(" ,-_")
        return merchant, (area if area else "Unknown")

    if lower.startswith("dupoint") or lower.startswith("dupont"):
        merchant = "Dupoint Lounge"
        m = re.match(r"(?i)(dupoint|dupont)\s+lounge\s*(.*)", s)
        area = (m.group(2).strip() if m else "") or "Unknown"
        return merchant, area

    if lower.startswith("greenview"):
        merchant = "Greenview Restaurant"
        area = s[len("Greenview"):].strip(" ,-_") or "Unknown"
        return merchant, area

    if lower.startswith("fish pit hub"):
        merchant = "Fish Pit Hub"
        area = s[len("Fish pit hub"):].strip(" ,-_") or "Unknown"
        return merchant, area

    if lower.startswith("junction pizza inn"):
        return "Pizza Inn", "Junction Mall"

    if lower.startswith("junction mall"):
        return "Junction Mall", "Junction Mall"

    if lower.startswith("leofresh"):
        merchant = "LeoFresh"
        area = s[len("LeoFresh"):].strip(" ,-_") or "Unknown"
        return merchant, area

    if lower.startswith("nairobi chapel"):
        merchant = "Nairobi Chapel"
        area = s[len("Nairobi Chapel"):].strip(" ,-_") or "Unknown"
        return merchant, area

    if lower.startswith("karura forest"):
        return "Karura Forest", "Karura"

    if lower.startswith("rockwell"):
        merchant = "Rockwell Service Station"
        area = s[len("Rockwell"):].strip(" ,-_") or "Unknown"
        return merchant, area

    if lower.startswith("kisii"):
        return "Kisii Contribution", "Kisii"

    if lower.startswith("naivasha road"):
        return "Naivasha Road", "Naivasha Road"

    # Generic fallbacks
    if "_" in s:
        m, a = s.split("_", 1)
        return m.strip() or "Unknown", a.strip() or "Unknown"

    words = s.split()
    if len(words) <= 2:
        return s, "Unknown"

    return " ".join(words[:2]), " ".join(words[2:]) or "Unknown"


def fmt_kes(x):
    return f"KSh {x:,.0f}"


def save_chart(fig, outpath, dpi=200):
    fig.tight_layout()
    fig.savefig(outpath, dpi=dpi)
    plt.close(fig)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Path to Excel file")
    ap.add_argument("--outdir", default="alex_expense_outputs_v2", help="Output directory for charts")
    ap.add_argument("--summary_json", default="alex_expense_summary_v2.json", help="Output JSON summary file")
    ap.add_argument("--report_pdf", default="Alex_Expense_Report_MerchantLocation_v2.pdf", help="Output PDF report file")
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)

    df = pd.read_excel(args.input)
    df["Date"] = pd.to_datetime(df["Date"])
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df[["Merchant", "Area"]] = df["Location"].apply(lambda x: pd.Series(parse_location(x)))
    df["Month"] = df["Date"].dt.to_period("M").astype(str)

    total = float(df["Amount"].sum())
    n = int(len(df))
    avg = float(df["Amount"].mean())
    median = float(df["Amount"].median())
    date_min = str(df["Date"].min().date())
    date_max = str(df["Date"].max().date())

    monthly = df.groupby("Month")["Amount"].sum().sort_index()
    type_tot = df.groupby("Type")["Amount"].sum().sort_values(ascending=False)
    pay_tot = df.groupby("Payment type")["Amount"].sum().sort_values(ascending=False)
    merch_tot = df.groupby("Merchant")["Amount"].sum().sort_values(ascending=False)
    area_tot = df.groupby("Area")["Amount"].sum().sort_values(ascending=False)

    formatter = FuncFormatter(lambda x, pos: f"{x:,.0f}")

    # 01 Monthly spend
    fig = plt.figure(figsize=(8, 4.5))
    plt.plot(monthly.index, monthly.values, marker="o")
    plt.gca().yaxis.set_major_formatter(formatter)
    plt.title("Monthly spend (KES)")
    plt.xlabel("Month")
    plt.ylabel("Amount (KSh)")
    plt.xticks(rotation=45, ha="right")
    p1 = os.path.join(args.outdir, "01_monthly_spend.png")
    save_chart(fig, p1)

    # 02 Spend by type
    fig = plt.figure(figsize=(8, 4.8))
    plt.bar(type_tot.index.astype(str), type_tot.values)
    plt.gca().yaxis.set_major_formatter(formatter)
    plt.title("Spend by category/type (KES)")
    plt.xlabel("Type")
    plt.ylabel("Amount (KSh)")
    plt.xticks(rotation=30, ha="right")
    p2 = os.path.join(args.outdir, "02_spend_by_type.png")
    save_chart(fig, p2)

    # 03 Payment method
    fig = plt.figure(figsize=(6.8, 4.5))
    plt.bar(pay_tot.index.astype(str), pay_tot.values)
    plt.gca().yaxis.set_major_formatter(formatter)
    plt.title("Spend by payment method (KES)")
    plt.xlabel("Payment method")
    plt.ylabel("Amount (KSh)")
    p3 = os.path.join(args.outdir, "03_payment_method.png")
    save_chart(fig, p3)

    # 04 Top merchants
    top_merch = merch_tot.head(10)[::-1]
    fig = plt.figure(figsize=(8, 5.2))
    plt.barh(top_merch.index.astype(str), top_merch.values)
    plt.gca().xaxis.set_major_formatter(formatter)
    plt.title("Top merchants by spend (Top 10, KES)")
    plt.xlabel("Amount (KSh)")
    plt.ylabel("Merchant")
    p4 = os.path.join(args.outdir, "04_top_merchants.png")
    save_chart(fig, p4)

    # 05 Top areas
    top_area = area_tot.head(10)[::-1]
    fig = plt.figure(figsize=(8, 5.2))
    plt.barh(top_area.index.astype(str), top_area.values)
    plt.gca().xaxis.set_major_formatter(formatter)
    plt.title("Top locations/areas by spend (Top 10, KES)")
    plt.xlabel("Amount (KSh)")
    plt.ylabel("Area")
    p5 = os.path.join(args.outdir, "05_top_areas.png")
    save_chart(fig, p5)

    # 06 Heatmap: top 8 areas vs types
    top_areas_list = area_tot.head(8).index.tolist()
    pivot = df[df["Area"].isin(top_areas_list)].pivot_table(
        index="Area", columns="Type", values="Amount", aggfunc="sum", fill_value=0
    )
    pivot = pivot.loc[top_areas_list]

    fig = plt.figure(figsize=(9, 5.2))
    ax = plt.gca()
    im = ax.imshow(pivot.values, aspect="auto")
    ax.set_xticks(np.arange(pivot.shape[1]))
    ax.set_xticklabels(pivot.columns.astype(str), rotation=30, ha="right")
    ax.set_yticks(np.arange(pivot.shape[0]))
    ax.set_yticklabels(pivot.index.astype(str))
    ax.set_title("Spend mix by top areas vs type (KES)")
    for i in range(pivot.shape[0]):
        for j in range(pivot.shape[1]):
            val = pivot.values[i, j]
            if val > 0:
                ax.text(j, i, f"{val:,.0f}", ha="center", va="center", fontsize=7)
    fig.colorbar(im, ax=ax, fraction=0.03, pad=0.02)
    p6 = os.path.join(args.outdir, "06_area_type_heatmap.png")
    save_chart(fig, p6)

    # Shell station detail
    shell_area = df[df["Merchant"] == "Shell"].groupby("Area")["Amount"].sum().sort_values(ascending=False)

    # Summary JSON (for the PPT script)
    type_share = (type_tot / total * 100).round(1)
    pay_share = (pay_tot / total * 100).round(1)

    summary = {
        "currency": "KES",
        "date_range": {"start": date_min, "end": date_max},
        "kpis": {
            "total_spend": total,
            "transactions": n,
            "avg_txn": avg,
            "median_txn": median,
            "highest_month": {"month": monthly.idxmax(), "amount": float(monthly.max())},
            "lowest_month": {"month": monthly.idxmin(), "amount": float(monthly.min())},
        },
        "by_type": [{"type": k, "amount": float(v), "share_pct": float(type_share.loc[k])} for k, v in type_tot.items()],
        "by_payment": [{"payment_type": k, "amount": float(v), "share_pct": float(pay_share.loc[k])} for k, v in pay_tot.items()],
        "top_merchants": [{"merchant": k, "amount": float(v)} for k, v in merch_tot.head(10).items()],
        "top_areas": [{"area": k, "amount": float(v)} for k, v in area_tot.head(10).items()],
        "shell_by_station": [{"station": k, "amount": float(v)} for k, v in shell_area.items()],
        "monthly_totals": [{"month": k, "amount": float(v)} for k, v in monthly.items()],
    }

    with open(args.summary_json, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2)

    # PDF Report (ReportLab)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Title"], fontSize=22, spaceAfter=10)
    h_style = ParagraphStyle("Heading", parent=styles["Heading2"], fontSize=14, spaceAfter=6)
    body = ParagraphStyle("Body", parent=styles["BodyText"], fontSize=10, leading=13)
    small = ParagraphStyle("Small", parent=styles["BodyText"], fontSize=8, leading=10, textColor=colors.grey)

    def kpi_table():
        kpi_data = [
            ["Total spend", fmt_kes(total)],
            ["Transactions", f"{n:,}"],
            ["Average transaction", fmt_kes(avg)],
            ["Median transaction", fmt_kes(median)],
            ["Highest month", f"{monthly.idxmax()} ({fmt_kes(monthly.max())})"],
            ["Lowest month", f"{monthly.idxmin()} ({fmt_kes(monthly.min())})"],
        ]
        t = Table(kpi_data, colWidths=[2.2 * inch, 4.8 * inch])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONT", (0, 0), (-1, -1), "Helvetica", 10),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke),
        ]))
        return t

    # Shell table
    shell_tbl_data = [["Shell station (from Location)", "Spend (KSh)"]]
    for k, v in shell_area.head(6).items():
        shell_tbl_data.append([k, fmt_kes(v)])
    shell_tbl = Table(shell_tbl_data, colWidths=[4.8 * inch, 2.2 * inch])
    shell_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9),
        ("FONT", (0, 1), (-1, -1), "Helvetica", 9),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
    ]))

    # Area table
    area_stats = df.groupby("Area")["Amount"].agg(["sum", "count"]).sort_values("sum", ascending=False).head(8)
    area_tbl_data = [["Area", "Spend (KSh)", "Txn count"]]
    for area, row in area_stats.iterrows():
        area_tbl_data.append([area, fmt_kes(row["sum"]), f"{int(row['count'])}"])
    area_tbl = Table(area_tbl_data, colWidths=[3.6 * inch, 2.0 * inch, 1.4 * inch])
    area_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9),
        ("FONT", (0, 1), (-1, -1), "Helvetica", 9),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
        ("ALIGN", (2, 1), (2, -1), "CENTER"),
    ]))

    # Insights bullets
    top_type = type_tot.index[0]
    top_pay = pay_tot.index[0]
    top_merchant = merch_tot.index[0]
    top_area = area_tot.index[0]
    insights = [
        f"Spending is concentrated: top category is <b>{top_type}</b> at {(type_tot.iloc[0] / total * 100):.1f}% of total.",
        f"Payments are mostly via <b>{top_pay}</b> ({(pay_tot.iloc[0] / total * 100):.1f}%).",
        f"Top merchant is <b>{top_merchant}</b> at {fmt_kes(merch_tot.iloc[0])} ({(merch_tot.iloc[0] / total * 100):.1f}%).",
        f"Top area/location is <b>{top_area}</b> at {fmt_kes(area_tot.iloc[0])} ({(area_tot.iloc[0] / total * 100):.1f}%).",
    ]
    bullet_html = "<br/>".join([f"&bull; {x}" for x in insights])

    rec = [
        "Track a monthly budget vs actual for the largest categories (Rent, Tithe, Car).",
        "For fuel/transport, monitor spend by station (Shell locations) to spot price/volume changes.",
        "If possible, separate 'Merchant' from 'Location' in the data capture to improve accuracy and automation.",
    ]

    doc = SimpleDocTemplate(args.report_pdf, pagesize=LETTER, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    story.append(Paragraph("Alex Monthly Spending Analysis (Kenyan Shillings)", title_style))
    story.append(Paragraph(f"Period covered: {date_min} to {date_max} &nbsp;&nbsp;|&nbsp;&nbsp; Data source: {os.path.basename(args.input)}", body))
    story.append(Spacer(1, 12))

    story.append(Paragraph("Executive summary", h_style))
    story.append(kpi_table())
    story.append(Spacer(1, 10))
    story.append(Paragraph(bullet_html, body))
    story.append(Spacer(1, 12))
    story.append(RLImage(p1, width=6.8 * inch, height=3.6 * inch))
    story.append(PageBreak())

    story.append(Paragraph("Category and payment breakdown", h_style))
    story.append(RLImage(p2, width=6.8 * inch, height=3.8 * inch))
    story.append(Spacer(1, 10))
    story.append(RLImage(p3, width=6.4 * inch, height=4.0 * inch))
    story.append(PageBreak())

    story.append(Paragraph("Merchant analysis", h_style))
    story.append(Paragraph("Merchants are derived from the Location text (e.g., 'Shell Dagoretti' -> merchant 'Shell').", body))
    story.append(Spacer(1, 10))
    story.append(RLImage(p4, width=6.8 * inch, height=4.0 * inch))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Shell spending by station (detail)", body))
    story.append(shell_tbl)
    story.append(PageBreak())

    story.append(Paragraph("Location/area analysis", h_style))
    story.append(Paragraph("Areas are extracted from the Location text (e.g., 'Home _Kikuyu Road' -> area 'Kikuyu Road').", body))
    story.append(Spacer(1, 10))
    story.append(RLImage(p5, width=6.8 * inch, height=4.0 * inch))
    story.append(Spacer(1, 8))
    story.append(area_tbl)
    story.append(PageBreak())

    story.append(Paragraph("Spend mix by area vs category", h_style))
    story.append(Paragraph("Heatmap of spend amounts across top areas and spending types (values shown are totals in KSh).", body))
    story.append(Spacer(1, 10))
    story.append(RLImage(p6, width=7.1 * inch, height=4.3 * inch))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Practical recommendations", h_style))
    story.append(Paragraph("<br/>".join([f"&bull; {x}" for x in rec]), body))
    story.append(Spacer(1, 6))
    story.append(Paragraph("All amounts are Kenyan Shillings (KSh). Merchant/area extraction is heuristic-based from the Location text.", small))

    doc.build(story)

    print("Done.")
    print("Charts:", args.outdir)
    print("Summary JSON:", args.summary_json)
    print("PDF report:", args.report_pdf)


if __name__ == "__main__":
    main()
