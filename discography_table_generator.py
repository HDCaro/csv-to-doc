import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.shared import OxmlElement, qn

BOOK_PATH = 'HITS AND HAPPINESS FINAL Format.docx'
OUTPUT_FILE = 'Hits And Happiness Final Discog.docx'


def compact_paragraph(p):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))


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
    year_range = f"{df_sorted['year'].min()}-{df_sorted['year'].max()}"

    # ----------------------------
    # GROUP DATA (KEY PART)
    # ----------------------------
    grouped = df_sorted.groupby(['year', 'artist', 'album'])

    # ----------------------------
    # OPEN BOOK
    # ----------------------------
    doc = Document(BOOK_PATH)
    doc.add_section(WD_SECTION.NEW_PAGE)

    # ----------------------------
    # TITLE
    # ----------------------------
    title = doc.add_paragraph("RICHARD NILES DISCOGRAPHY BY YEAR")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

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

    for run in summary.runs:
        run.font.size = Pt(9)
        run.italic = True

    # ----------------------------
    # SIMPLIFIED TABLE
    # ----------------------------
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ['Artist', 'Album', 'Details', 'Role']

    for i, text in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = text

        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(10)

    current_year = None

    for (year, artist, album), group in grouped:

        # YEAR ROW
        if year != current_year:
            current_year = year

            row = table.add_row()
            merged = row.cells[0]
            for i in range(1, 4):
                merged = merged.merge(row.cells[i])

            merged.text = str(year)

            for p in merged.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in p.runs:
                    run.bold = True
                    run.font.size = Pt(11)

        # ROLE (aggregate)
        roles = set()
        for _, r in group.iterrows():
            if str(r.get('producer')) == 'True': roles.add('P')
            if str(r.get('arranger')) == 'True': roles.add('A')
            if str(r.get('composer')) == 'True': roles.add('C')

        role_text = ', '.join(sorted(roles))

        # DETAILS
        if len(group) == 1:
            details = group.iloc[0]['track_title']
            album_label = "Single"
        else:
            details = f"{len(group)} tracks"
            album_label = album

        # ROW
        row = table.add_row()
        prevent_row_split(row)

        row.cells[0].text = artist
        row.cells[1].text = album_label
        row.cells[2].text = details
        row.cells[3].text = role_text

        for i, cell in enumerate(row.cells):
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i < 3 else WD_ALIGN_PARAGRAPH.CENTER
                compact_paragraph(p)
                for run in p.runs:
                    run.font.size = Pt(9)

    doc.save(OUTPUT_FILE)

    print("Book version generated.")


if __name__ == "__main__":
    create_discography()