import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from datetime import datetime


def add_border_to_cell(cell, border_type="bottom", color="808080", size=2):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcBorders = OxmlElement('w:tcBorders')

    if border_type == "bottom":
        border = OxmlElement('w:bottom')
    else:
        border = OxmlElement('w:top')

    border.set(qn('w:val'), 'single')
    border.set(qn('w:sz'), str(size * 4))
    border.set(qn('w:color'), color)

    tcBorders.append(border)
    tcPr.append(tcBorders)


def compact_paragraph(paragraph):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1


def set_heading2_grey(doc):
    style = doc.styles['Heading 2']
    font = style.font
    font.color.rgb = RGBColor(89, 89, 89)


def create_compact_discography_table():

    try:
        df = pd.read_csv(
            'richard_niles_discography.csv',
            sep='\t',
            encoding='cp1252'
        )
    except FileNotFoundError:
        print("Error: richard_niles_discography.csv not found")
        return

    df = df.fillna('')
    df['producer'] = df['producer'].astype(str) == 'True'
    df['arranger'] = df['arranger'].astype(str) == 'True'
    df['composer'] = df['composer'].astype(str) == 'True'

    df['year'] = pd.to_numeric(df['year'], errors='coerce')
    df = df.dropna(subset=['year'])
    df['year'] = df['year'].astype(int)

    df_sorted = df.sort_values(['year', 'artist', 'album', 'track_title'])

    total_tracks = len(df_sorted)
    total_years = df_sorted['year'].nunique()
    year_range = f"{df_sorted['year'].min()}-{df_sorted['year'].max()}"

    doc = Document()
    set_heading2_grey(doc)

    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Header
    header = doc.sections[0].header
    p = header.paragraphs[0]
    p.text = "Richard Niles Discography by Year - Compact Format"
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    compact_paragraph(p)
    p.runs[0].font.size = Pt(9)

    # Footer
    footer = doc.sections[0].footer
    p = footer.paragraphs[0]
    p.text = f"{total_tracks} Tracks • {total_years} Years ({year_range}) • Generated {datetime.now().strftime('%B %d, %Y')}"
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    compact_paragraph(p)
    p.runs[0].font.size = Pt(8)

    # Title
    title = doc.add_heading('Richard Niles Discography by Year', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    compact_paragraph(title)
    title.runs[0].font.size = Pt(18)

    # Subtitle
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    compact_paragraph(subtitle)
    run = subtitle.add_run(f'Compact Chronological Format • {total_tracks} Tracks • {year_range}')
    run.font.size = Pt(11)
    run.italic = True

    # Table
    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    table.columns[0].width = Inches(0.65)
    table.columns[1].width = Inches(1.46)
    table.columns[2].width = Inches(1.46)
    table.columns[3].width = Inches(1.46)
    table.columns[4].width = Inches(0.84)

    headers = ['Year', 'Artist', 'Album', 'Track Title', 'Role']

    for i, text in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = text

        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            compact_paragraph(p)
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)

        add_border_to_cell(cell, "bottom", "808080", 4)

    current_year = None

    for _, row in df_sorted.iterrows():

        roles = []
        if row['producer']: roles.append('P')
        if row['arranger']: roles.append('A')
        if row['composer']: roles.append('C')
        if row['other_roles']: roles.append(str(row['other_roles']))

        role_text = ', '.join(roles)

        table_row = table.add_row()
        cells = table_row.cells

        is_first = current_year != row['year']
        if is_first:
            current_year = row['year']

        # Year
        cell = cells[0]
        cell.text = str(row['year']) if is_first else ""
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            compact_paragraph(p)
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)
        if is_first:
            add_border_to_cell(cell, "top")

        # Artist
        cell = cells[1]
        cell.text = str(row['artist'])
        for p in cell.paragraphs:
            compact_paragraph(p)
            for run in p.runs:
                run.font.size = Pt(9)
        if is_first:
            add_border_to_cell(cell, "top")

        # Album
        cell = cells[2]
        cell.text = str(row['album'])
        for p in cell.paragraphs:
            compact_paragraph(p)
            for run in p.runs:
                run.font.size = Pt(9)
                run.font.italic = True
        if is_first:
            add_border_to_cell(cell, "top")

        # Track
        cell = cells[3]
        cell.text = str(row['track_title'])
        for p in cell.paragraphs:
            compact_paragraph(p)
            for run in p.runs:
                run.font.size = Pt(9)
        if is_first:
            add_border_to_cell(cell, "top")

        # Role
        cell = cells[4]
        cell.text = role_text
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            compact_paragraph(p)
            for run in p.runs:
                run.font.size = Pt(8)
                run.font.bold = True
        if is_first:
            add_border_to_cell(cell, "top")

        table_row.height = Pt(12)

    # Legend
    doc.add_heading('Role Legend', 2)
    p = doc.add_paragraph()
    compact_paragraph(p)
    p.add_run('P').bold = True
    p.add_run(' = Producer • ')
    p.add_run('A').bold = True
    p.add_run(' = Arranger • ')
    p.add_run('C').bold = True
    p.add_run(' = Composer • Other roles listed in full')

    # Notes
    doc.add_heading('Format Notes', 2)
    p = doc.add_paragraph()
    compact_paragraph(p)
    p.add_run('This compact format presents all ')
    p.add_run(f'{total_tracks} tracks').bold = True
    p.add_run(' in a space-efficient layout with aligned columns and minimal spacing.')

    # Save
    doc.save('richard_niles_discography_compact.docx')

    print("Document generated successfully.")
    print(f"Tracks: {total_tracks}")
    print(f"Years: {year_range}")


if __name__ == "__main__":
    create_compact_discography_table()