import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.shared import OxmlElement, qn

# ----------------------------
# CONFIG
# ----------------------------
BOOK_PATH = 'HITS AND HAPPINESS FINAL Format.docx'
OUTPUT_FILE = 'Hits And Happiness Final Discog.docx'


# ----------------------------
# Helpers
# ----------------------------
def add_border_to_cell(cell, border_type="bottom", color="808080", size=2):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcBorders = OxmlElement('w:tcBorders')
    border = OxmlElement(f'w:{border_type}')
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
    pf.keep_together = True


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))


# ----------------------------
# Main
# ----------------------------
def create_discography():

    df = pd.read_csv(
        'richard_niles_discography.csv',
        sep='\t',
        encoding='cp1252'
    )

    df = df.fillna('')
    df['year'] = pd.to_numeric(df['year'], errors='coerce')
    df = df.dropna(subset=['year'])
    df['year'] = df['year'].astype(int)

    df_sorted = df.sort_values(['year', 'artist', 'album', 'track_title'])

    total_tracks = len(df_sorted)
    total_years = df_sorted['year'].nunique()
    year_range = f"{df_sorted['year'].min()}-{df_sorted['year'].max()}"

    # ----------------------------
    # Open book template
    # ----------------------------
    doc = Document(BOOK_PATH)
    doc.add_section(WD_SECTION.NEW_PAGE)

    # ----------------------------
    # TITLE (ALL CAPS, Georgia 16)
    # ----------------------------
    title = doc.add_paragraph("RICHARD NILES DISCOGRAPHY BY YEAR")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    pf = title.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1

    for run in title.runs:
        run.bold = True
        run.font.size = Pt(16)
        run.font.name = "Georgia"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Georgia')

    # ----------------------------
    # SUMMARY
    # ----------------------------
    summary = doc.add_paragraph(
        f"Discography of {total_tracks} tracks ({year_range}), documenting Richard Niles’ work as producer (P), arranger (A), and composer (C)."
    )
    summary.alignment = WD_ALIGN_PARAGRAPH.CENTER

    pf = summary.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1

    for run in summary.runs:
        run.font.size = Pt(9)
        run.italic = True

    # ----------------------------
    # ROLE LEGEND LINE
    # ----------------------------
    legend = doc.add_paragraph(
        "P = Producer | A = Arranger | C = Composer"
    )
    legend.alignment = WD_ALIGN_PARAGRAPH.CENTER

    pf = legend.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1

    for run in legend.runs:
        run.font.size = Pt(8)

    # ----------------------------
    # TABLE
    # ----------------------------
    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    widths = [0.65, 1.46, 1.46, 1.46, 0.84]

    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)

    headers = ['Year', 'Artist', 'Album', 'Track Title', 'Role']

    for i, text in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = text

        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            compact_paragraph(p)
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(10)

        add_border_to_cell(cell, "bottom", "808080", 4)

    current_year = None

    for _, row in df_sorted.iterrows():

        table_row = table.add_row()
        prevent_row_split(table_row)

        cells = table_row.cells

        for i, w in enumerate(widths):
            cells[i].width = Inches(w)

        is_first = current_year != row['year']
        if is_first:
            current_year = row['year']

        values = [
            str(row['year']) if is_first else "",
            str(row['artist']).strip(),
            str(row['album']).strip(),
            str(row['track_title']).strip(),
            ""
        ]

        roles = []
        if str(row.get('producer')) == 'True': roles.append('P')
        if str(row.get('arranger')) == 'True': roles.append('A')
        if str(row.get('composer')) == 'True': roles.append('C')

        values[4] = ', '.join(roles)

        for i, val in enumerate(values):
            cell = cells[i]
            cell.text = val

            for p in cell.paragraphs:
                compact_paragraph(p)

                if i == 0 or i == 4:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                for run in p.runs:
                    run.font.size = Pt(9)

        if is_first:
            for cell in cells:
                add_border_to_cell(cell, "top", "808080", 2)

        table_row.height = Pt(12)

    doc.save(OUTPUT_FILE)

    print("\n--- Discography Generation Complete ---")
    print(f"File: {OUTPUT_FILE}")
    print(f"Tracks: {total_tracks}")
    print(f"Years: {year_range}")
    print(f"Total years: {total_years}")
    print("--------------------------------------\n")


if __name__ == "__main__":
    create_discography()