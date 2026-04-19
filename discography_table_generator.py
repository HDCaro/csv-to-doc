import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.shared import OxmlElement, qn

BOOK_PATH = 'HITS AND HAPPINESS FINAL Format.docx'
OUTPUT_FILE = 'Hits And Happiness Final Discog.docx'


# ----------------------------
# Helpers
# ----------------------------
def set_cell_text(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, hanging=False):
    cell.text = ""

    p = cell.paragraphs[0]
    p.alignment = align

    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1

    if hanging:
        pf.left_indent = Pt(8)
        pf.first_line_indent = Pt(-8)

    run = p.add_run(text)
    run.font.size = Pt(9)
    run.font.name = "Georgia"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Georgia')


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))


def smart_text(text, max_len=40, hard_limit=70):
    text = str(text).strip()
    if len(text) <= max_len:
        return text
    if len(text) <= hard_limit:
        return text
    return text[:hard_limit - 3].rstrip() + "..."


def split_artists(text):
    parts = [p.strip() for p in str(text).split(",")]
    if len(parts) <= 1:
        return text
    return ",\n".join(parts)


# ----------------------------
# Adaptive sizing (balanced)
# ----------------------------
def compute_column_ratios(df_grouped):
    max_artist = 1
    max_album = 1
    max_details = 1

    for (year, artist, album), group in df_grouped:
        max_artist = max(max_artist, len(str(artist)))
        max_album = max(max_album, len(str(album)))

        if len(group) == 1:
            details = str(group.iloc[0]['track_title'])
        else:
            details = f"{len(group)} tracks"

        max_details = max(max_details, len(details))

    role_weight = 6
    total = max_artist + max_album + max_details + role_weight

    # Optical adjustment
    artist = (max_artist / total) * 0.85
    album = (max_album / total)
    details = (max_details / total) * 1.10
    role = (role_weight / total)

    total2 = artist + album + details + role

    return [
        artist / total2,
        album / total2,
        details / total2,
        role / total2
    ]


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
    grouped = df_sorted.groupby(['year', 'artist', 'album'])

    total_tracks = len(df_sorted)
    year_range = f"{df_sorted['year'].min()}-{df_sorted['year'].max()}"

    col_ratios = compute_column_ratios(grouped)

    doc = Document(BOOK_PATH)
    doc.add_section(WD_SECTION.NEW_PAGE)

    # ----------------------------
    # Title
    # ----------------------------
    title = doc.add_paragraph("RICHARD NILES DISCOGRAPHY BY YEAR")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for run in title.runs:
        run.bold = True
        run.font.size = Pt(16)
        run.font.name = "Georgia"

    # ----------------------------
    # Summary
    # ----------------------------
    summary = doc.add_paragraph(
        f"Discography of {total_tracks} tracks ({year_range}), documenting Richard Niles’ work as producer (P), arranger (A), and composer (C)."
    )
    summary.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for run in summary.runs:
        run.font.size = Pt(9)
        run.italic = True

    # ----------------------------
    # Table
    # ----------------------------
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    table.allow_autofit = False

    # 🔥 EXACT WIDTH FIX (no gap, no overflow)
    EMU_PER_INCH = 914400
    section = doc.sections[-1]
    usable_width = section.page_width - section.left_margin - section.right_margin
    usable_inches = usable_width / EMU_PER_INCH

    raw_widths = [usable_inches * r for r in col_ratios]

    # correction to eliminate rounding gap
    correction = usable_inches - sum(raw_widths)
    raw_widths[2] += correction  # apply to DETAILS column

    widths = [Inches(w) for w in raw_widths]

    for i, w in enumerate(widths):
        table.columns[i].width = w

    headers = ['Artist', 'Album', 'Details', 'Role']

    for i, text in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, text, WD_ALIGN_PARAGRAPH.CENTER)
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(10)

    current_year = None

    for (year, artist, album), group in grouped:

        if year != current_year:
            current_year = year

            row = table.add_row()
            merged = row.cells[0]
            for i in range(1, 4):
                merged = merged.merge(row.cells[i])

            p = merged.paragraphs[0]
            p.paragraph_format.space_before = Pt(6)

            run = p.add_run(str(year))
            run.bold = True
            run.font.size = Pt(12)

        roles = set()
        for _, r in group.iterrows():
            if str(r.get('producer')) == 'True': roles.add('P')
            if str(r.get('arranger')) == 'True': roles.add('A')
            if str(r.get('composer')) == 'True': roles.add('C')

        role_text = ', '.join(sorted(roles))

        if len(group) == 1:
            details = smart_text(group.iloc[0]['track_title'])
            album_label = "Single"
        else:
            details = f"{len(group)} tracks"
            album_label = smart_text(album)

        row = table.add_row()
        prevent_row_split(row)

        for i, w in enumerate(widths):
            row.cells[i].width = w

        artist_text = split_artists(smart_text(artist, 60, 120))

        set_cell_text(row.cells[0], artist_text, hanging=True)
        set_cell_text(row.cells[1], album_label, hanging=True)
        set_cell_text(row.cells[2], details, hanging=True)
        set_cell_text(row.cells[3], role_text, WD_ALIGN_PARAGRAPH.CENTER)

    doc.save(OUTPUT_FILE)

    print("\n✔ FINAL PERFECT FIT — no overflow, no gap, margins aligned\n")


if __name__ == "__main__":
    create_discography()