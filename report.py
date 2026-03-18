from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os

os.chdir(r'D:\Portfolio projects\Upwork\sales_dashboard')

# ── Setup ─────────────────────────────────────────────
doc = SimpleDocTemplate(
    "Superstore_Sales_Analysis.pdf",
    pagesize=A4,
    rightMargin=0.75*inch,
    leftMargin=0.75*inch,
    topMargin=0.75*inch,
    bottomMargin=0.75*inch
)

styles = getSampleStyleSheet()
story = []

# ── Custom Styles ──────────────────────────────────────
title_style = ParagraphStyle(
    'CustomTitle',
    parent=styles['Title'],
    fontSize=24,
    textColor=colors.HexColor('#2c3e50'),
    spaceAfter=6,
    alignment=TA_CENTER,
    fontName='Helvetica-Bold'
)

subtitle_style = ParagraphStyle(
    'Subtitle',
    parent=styles['Normal'],
    fontSize=12,
    textColor=colors.HexColor('#7f8c8d'),
    spaceAfter=4,
    alignment=TA_CENTER
)

section_style = ParagraphStyle(
    'Section',
    parent=styles['Heading1'],
    fontSize=14,
    textColor=colors.HexColor('#2980b9'),
    spaceBefore=16,
    spaceAfter=8,
    fontName='Helvetica-Bold'
)

body_style = ParagraphStyle(
    'Body',
    parent=styles['Normal'],
    fontSize=10,
    textColor=colors.HexColor('#2c3e50'),
    spaceAfter=6,
    leading=16
)

# ── Title Page ─────────────────────────────────────────
story.append(Spacer(1, 0.5*inch))
story.append(Paragraph("Superstore Sales Analysis", title_style))
story.append(Paragraph("Business Intelligence Report — 2014 to 2017", subtitle_style))
story.append(Paragraph("Prepared by: Muhammad Ammar Yameen Shizgar", subtitle_style))
story.append(Paragraph("Data Analyst | Python | SQL | Data Visualization", subtitle_style))
story.append(Spacer(1, 0.2*inch))
story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#2980b9')))
story.append(Spacer(1, 0.2*inch))

# ── Executive Summary ──────────────────────────────────
story.append(Paragraph("Executive Summary", section_style))
story.append(Paragraph(
    "This report presents a comprehensive analysis of Superstore sales data spanning four years "
    "(2014–2017). The analysis covers revenue performance, profitability by region and category, "
    "monthly sales trends, and the impact of discounting on profit margins. Key findings reveal "
    "a total revenue of $2,297,200.86 with a net profit margin of 12.5%, with the West region "
    "and Technology category driving the strongest returns.",
    body_style
))

# ── Key Metrics Table ──────────────────────────────────
story.append(Paragraph("Key Business Metrics", section_style))

metrics_data = [
    ['Metric', 'Value'],
    ['Total Revenue', '$2,297,200.86'],
    ['Total Profit', '$286,397.02'],
    ['Profit Margin', '12.5%'],
    ['Total Orders', '5,009'],
    ['Total Customers', '793'],
    ['Best Performing Region', 'West'],
    ['Best Performing Category', 'Technology'],
]

metrics_table = Table(metrics_data, colWidths=[3*inch, 3*inch])
metrics_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2980b9')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, 0), 11),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#ecf0f1'), colors.white]),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#bdc3c7')),
    ('ROWHEIGHT', (0, 0), (-1, -1), 20),
    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
    ('FONTSIZE', (0, 1), (-1, -1), 10),
]))
story.append(metrics_table)

# ── Chart 1 ───────────────────────────────────────────
story.append(Paragraph("Sales Performance by Category", section_style))
story.append(Paragraph(
    "Technology leads all categories in total sales, followed by Furniture and Office Supplies. "
    "This suggests the business should prioritize Technology inventory and marketing efforts "
    "to maximize revenue potential.",
    body_style
))
story.append(Image('chart1_sales_by_category.png', width=6*inch, height=3*inch))

# ── Chart 2 ───────────────────────────────────────────
story.append(Paragraph("Monthly Sales Trend (2014–2017)", section_style))
story.append(Paragraph(
    "Sales show a consistent upward trend year over year, with notable spikes in Q4 each year "
    "suggesting strong seasonal demand. This pattern should inform inventory planning and "
    "staffing decisions ahead of peak periods.",
    body_style
))
story.append(Image('Chart2_monthly_trend.png', width=6*inch, height=3*inch))

# ── Chart 3 ───────────────────────────────────────────
story.append(Paragraph("Profitability by Region", section_style))
story.append(Paragraph(
    "The West region generates the highest profit, while the Central region underperforms "
    "relative to its sales volume. A deeper review of Central region pricing and discount "
    "strategies is recommended to improve profitability.",
    body_style
))
story.append(Image('Chart3_Profit_by_region.png', width=6*inch, height=3*inch))

# ── Chart 4 ───────────────────────────────────────────
story.append(Paragraph("Top 10 Sub-Categories by Sales", section_style))
story.append(Paragraph(
    "Phones and Chairs dominate sub-category sales, together accounting for a significant "
    "portion of total revenue. Storage and Tables follow, indicating strong demand across "
    "both Technology and Furniture segments.",
    body_style
))
story.append(Image('Chart4_Top_SC.png', width=6*inch, height=3*inch))

# ── Chart 5 ───────────────────────────────────────────
story.append(Paragraph("Discount Impact on Profit", section_style))
story.append(Paragraph(
    "A clear negative correlation exists between discount rate and profit. Orders with discounts "
    "above 40% consistently result in losses. It is strongly recommended to cap discounts at "
    "20% to protect profit margins across all categories.",
    body_style
))
story.append(Image('Chart5_discount_vs_proft.png', width=6*inch, height=3*inch))

# ── Recommendations ────────────────────────────────────
story.append(Paragraph("Strategic Recommendations", section_style))

recommendations = [
    "1.  <b>Cap discounts at 20%</b> — discounts above 40% consistently generate losses across all categories.",
    "2.  <b>Double down on Technology</b> — highest revenue and profit category, deserves increased investment.",
    "3.  <b>Review Central region strategy</b> — underperforming on profit despite reasonable sales volume.",
    "4.  <b>Prepare for Q4 demand spikes</b> — consistent seasonal pattern requires proactive inventory planning.",
    "5.  <b>Focus on Phones and Chairs</b> — top sub-categories by sales, strong candidates for upselling.",
]

for rec in recommendations:
    story.append(Paragraph(rec, body_style))

# ── Footer ─────────────────────────────────────────────
story.append(Spacer(1, 0.3*inch))
story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#bdc3c7')))
story.append(Spacer(1, 0.1*inch))
story.append(Paragraph(
    "Report prepared using Python (Pandas, Matplotlib, Seaborn, ReportLab) | "
    "Muhammad Ammar Yameen Shizgar | shizgar2015@gmail.com",
    subtitle_style
))

# ── Build PDF ──────────────────────────────────────────
doc.build(story)
print("✅ PDF report generated: Superstore_Sales_Analysis.pdf")