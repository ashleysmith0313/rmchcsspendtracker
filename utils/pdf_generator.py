import os
import io
from datetime import datetime
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ── Brand colors ──────────────────────────────────────────────────────────────
NAVY      = colors.HexColor('#0f1724')
BLUE      = colors.HexColor('#1a3a5c')
LIGHTBLUE = colors.HexColor('#3b82f6')
SLATE     = colors.HexColor('#64748b')
LIGHT     = colors.HexColor('#f1f5f9')
WHITE     = colors.white
BORDER    = colors.HexColor('#e2e8f0')

REPORTS_DIR = os.path.join(os.path.dirname(__file__), '..', 'reports')

def generate_pdf_report(
    df: pd.DataFrame,
    week_ending: str,
    title: str,
    prepared_by: str = "Vista Staffing Solutions",
    include_detail: bool = True,
    include_notes: bool = True
) -> str:
    """Generate a PDF spend report and return the file path."""
    os.makedirs(REPORTS_DIR, exist_ok=True)
    filename = f"RMCHCS_SpendReport_{week_ending}.pdf"
    filepath = os.path.join(REPORTS_DIR, filename)

    doc = SimpleDocTemplate(
        filepath,
        pagesize=letter,
        rightMargin=0.65 * inch,
        leftMargin=0.65 * inch,
        topMargin=0.65 * inch,
        bottomMargin=0.65 * inch
    )

    styles = getSampleStyleSheet()

    # Custom styles
    header_style = ParagraphStyle('header', fontName='Helvetica-Bold', fontSize=18,
                                   textColor=WHITE, leading=22, alignment=TA_LEFT)
    sub_style = ParagraphStyle('sub', fontName='Helvetica', fontSize=9,
                                textColor=colors.HexColor('#94a3b8'), leading=13, alignment=TA_LEFT)
    section_style = ParagraphStyle('section', fontName='Helvetica-Bold', fontSize=11,
                                    textColor=NAVY, leading=14, spaceBefore=14, spaceAfter=6)
    body_style = ParagraphStyle('body', fontName='Helvetica', fontSize=9,
                                 textColor=SLATE, leading=13)
    kpi_label_style = ParagraphStyle('kpi_label', fontName='Helvetica', fontSize=7,
                                      textColor=SLATE, leading=10, alignment=TA_CENTER)
    kpi_value_style = ParagraphStyle('kpi_value', fontName='Helvetica-Bold', fontSize=16,
                                      textColor=NAVY, leading=20, alignment=TA_CENTER)
    small_style = ParagraphStyle('small', fontName='Helvetica', fontSize=8,
                                  textColor=SLATE, leading=11)
    right_style = ParagraphStyle('right', fontName='Helvetica', fontSize=8,
                                  textColor=SLATE, leading=11, alignment=TA_RIGHT)

    story = []

    # ── Header banner ──
    header_data = [[
        Paragraph(f"<b>RMCHCS Weekly Spend Report</b>", header_style),
        Paragraph(f"Week Ending: {week_ending}<br/>Generated: {datetime.now().strftime('%B %d, %Y')}<br/>Prepared by: {prepared_by}", sub_style)
    ]]
    header_table = Table(header_data, colWidths=[4.2 * inch, 2.8 * inch])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), NAVY),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (0, -1), 18),
        ('RIGHTPADDING', (-1, 0), (-1, -1), 16),
        ('TOPPADDING', (0, 0), (-1, -1), 16),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 16),
        ('ROUNDEDCORNERS', [8, 8, 8, 8]),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 0.2 * inch))

    # ── KPI summary row ──
    total_spend = df['total_spend'].sum()
    total_providers = df['provider_name'].nunique()
    total_specialties = df['specialty'].nunique()
    total_days = df['days_worked'].sum()

    kpi_items = [
        ("TOTAL SPEND", f"${total_spend:,.2f}"),
        ("PROVIDERS", str(total_providers)),
        ("SPECIALTIES", str(total_specialties)),
        ("TOTAL DAYS", f"{total_days:.1f}"),
    ]
    kpi_cells = []
    for label, value in kpi_items:
        kpi_cells.append([
            Paragraph(label, kpi_label_style),
            Paragraph(value, kpi_value_style)
        ])

    kpi_table = Table(
        [[cell[0] for cell in kpi_cells], [cell[1] for cell in kpi_cells]],
        colWidths=[1.74 * inch] * 4
    )
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), LIGHT),
        ('BOX', (0, 0), (-1, -1), 1, BORDER),
        ('LINEAFTER', (0, 0), (2, -1), 0.5, BORDER),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROUNDEDCORNERS', [6, 6, 6, 6]),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 0.18 * inch))

    # ── Specialty summary ──
    story.append(Paragraph("Spend by Specialty", section_style))
    spec_summary = df.groupby('specialty').agg(
        Providers=('provider_name', 'nunique'),
        Days=('days_worked', 'sum'),
        Spend=('total_spend', 'sum')
    ).reset_index().sort_values('Spend', ascending=False)

    spec_header = [['Specialty', 'Providers', 'Days Worked', 'Total Spend', '% of Week']]
    spec_rows = []
    for _, row in spec_summary.iterrows():
        pct = (row['Spend'] / total_spend * 100) if total_spend > 0 else 0
        spec_rows.append([
            row['specialty'],
            str(int(row['Providers'])),
            f"{row['Days']:.1f}",
            f"${row['Spend']:,.2f}",
            f"{pct:.1f}%"
        ])
    spec_rows.append(['TOTAL', str(total_providers), f"{total_days:.1f}", f"${total_spend:,.2f}", "100%"])

    spec_table = Table(spec_header + spec_rows, colWidths=[2.4*inch, 1.0*inch, 1.0*inch, 1.3*inch, 0.95*inch])
    spec_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('TEXTCOLOR', (0, 1), (-1, -1), NAVY),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [WHITE, LIGHT]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#dbeafe')),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, -1), (-1, -1), NAVY),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (0, -1), 10),
        ('GRID', (0, 0), (-1, -1), 0.4, BORDER),
    ])
    spec_table.setStyle(spec_style)
    story.append(spec_table)
    story.append(Spacer(1, 0.15 * inch))

    # ── Service line summary ──
    story.append(Paragraph("Spend by Service Line", section_style))
    sl_summary = df.groupby('service_line').agg(
        Providers=('provider_name', 'nunique'),
        Spend=('total_spend', 'sum')
    ).reset_index().sort_values('Spend', ascending=False)

    sl_header = [['Service Line', 'Providers', 'Total Spend', '% of Week']]
    sl_rows = []
    for _, row in sl_summary.iterrows():
        pct = (row['Spend'] / total_spend * 100) if total_spend > 0 else 0
        sl_rows.append([
            row['service_line'],
            str(int(row['Providers'])),
            f"${row['Spend']:,.2f}",
            f"{pct:.1f}%"
        ])

    sl_table = Table(sl_header + sl_rows, colWidths=[2.4*inch, 1.1*inch, 1.5*inch, 1.1*inch])
    sl_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('TEXTCOLOR', (0, 1), (-1, -1), NAVY),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, LIGHT]),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (0, -1), 10),
        ('GRID', (0, 0), (-1, -1), 0.4, BORDER),
    ]))
    story.append(sl_table)

    # ── Provider detail table ──
    if include_detail:
        story.append(Spacer(1, 0.1 * inch))
        story.append(Paragraph("Provider Detail", section_style))

        detail_cols = ['provider_name', 'provider_type', 'specialty', 'service_line', 'days_worked', 'daily_rate', 'total_spend']
        if include_notes:
            detail_cols.append('notes')

        header_row = [['Provider', 'Type', 'Specialty', 'Svc Line', 'Days', 'Rate/Day', 'Total']]
        if include_notes:
            header_row[0].append('Notes')

        detail_rows = []
        for _, row in df.sort_values('specialty').iterrows():
            r = [
                row['provider_name'],
                row.get('provider_type', ''),
                row['specialty'],
                row['service_line'],
                f"{row['days_worked']:.1f}",
                f"${row['daily_rate']:,.2f}",
                f"${row['total_spend']:,.2f}",
            ]
            if include_notes:
                r.append(str(row.get('notes', '') or ''))
            detail_rows.append(r)

        if include_notes:
            col_widths = [1.4*inch, 0.7*inch, 1.1*inch, 0.75*inch, 0.45*inch, 0.75*inch, 0.8*inch, 1.2*inch]
        else:
            col_widths = [1.7*inch, 0.8*inch, 1.4*inch, 0.9*inch, 0.5*inch, 0.85*inch, 1.0*inch]

        detail_table = Table(header_row + detail_rows, colWidths=col_widths)
        detail_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 7.5),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7.5),
            ('TEXTCOLOR', (0, 1), (-1, -1), NAVY),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, LIGHT]),
            ('ALIGN', (4, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (6, 0), (6, -1), 'RIGHT'),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING', (0, 0), (0, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.3, BORDER),
        ]))
        story.append(detail_table)

    # ── Footer ──
    story.append(Spacer(1, 0.25 * inch))
    story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER))
    story.append(Spacer(1, 0.08 * inch))
    footer_data = [[
        Paragraph(f"Confidential | {prepared_by}", small_style),
        Paragraph(f"Report generated {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", right_style)
    ]]
    footer_table = Table(footer_data, colWidths=[3.5 * inch, 3.5 * inch])
    footer_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(footer_table)

    doc.build(story)
    return filepath
