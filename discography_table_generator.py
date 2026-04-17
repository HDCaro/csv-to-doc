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
def compact_paragraph(p):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1
    pf.keep_together = True
    pf.widow_control = True


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

    result = parts[0]
    for p in parts[1:]:
        result += ",\n" + p

    return result


# 🔥 FINAL TEXT FUNCTION (HANGING INDENT SUPPORT)
def set_cell_text(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, hanging=False):
    cell.text = ""

    p = cell.paragraphs[0]
    p.alignment = align

    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1
    pf.keep_together = True
    pf.widow_control = True

    # 🔥 Hanging indent (key fix)
    if hanging:
        pf.left_indent = Pt(8)
        pf.first_line_indent = Pt(-8)

    run = p.add_run(text)
    run.font.size = Pt(9)


# ----------------------------
# Adaptive sizing
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

    return [
        max_artist / total,
        max_album / total,
        max_details / total,
        role_weight / total
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

    total_tracks = len(df_sorted)
    year_range = f"{df_sorted['year'].min()}-{df_sorted['year'].max()}"

    grouped = df_sorted.groupby(['year', 'artist', 'album'])

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
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Georgia')

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
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.allow_autofit = False

    EMU_PER_INCH = 914400
    section = doc.sections[-1]
    usable_width = section.page_width - section.left_margin - section.right_margin
    usable_inches = usable_width / EMU_PER_INCH

    widths = [Inches(usable_inches * r) for r in col_ratios]

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

        # ----------------------------
        # YEAR ROW
        # ----------------------------
        if year != current_year:
            current_year = year

            row = table.add_row()
            merged = row.cells[0]
            for i in range(1, 4):
                merged = merged.merge(row.cells[i])

            merged.text = ""
            p = merged.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(6)

            run = p.add_run(str(year))
            run.bold = True
            run.font.size = Pt(12)

            # separator line
            tc = merged._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')

            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '6')
            bottom.set(qn('w:color'), '808080')

            tcBorders.append(bottom)
            tcPr.append(tcBorders)

        # roles
        roles = set()
        for _, r in group.iterrows():
            if str(r.get('producer')) == 'True': roles.add('P')
            if str(r.get('arranger')) == 'True': roles.add('A')
            if str(r.get('composer')) == 'True': roles.add('C')

        role_text = ', '.join(sorted(roles))

        # details
        if len(group) == 1:
            details = smart_text(group.iloc[0]['track_title'])
            album_label = "Single"
        else:
            details = f"{len(group)} tracks"
            album_label = smart_text(album)

        # row
        row = table.add_row()
        prevent_row_split(row)

        for i, w in enumerate(widths):
            row.cells[i].width = w

        artist_text = split_artists(smart_text(artist, 60, 120))

        # 🔥 APPLY HANGING TO ALL TEXT COLUMNS
        set_cell_text(row.cells[0], artist_text, hanging=True)
        set_cell_text(row.cells[1], album_label, hanging=True)
        set_cell_text(row.cells[2], details, hanging=True)

        # Role (no hanging)
        set_cell_text(row.cells[3], role_text, WD_ALIGN_PARAGRAPH.CENTER, hanging=False)

    doc.save(OUTPUT_FILE)

    print("\n--- Discography Generation Complete ---")
    print(f"File: {OUTPUT_FILE}")
    print(f"Tracks: {total_tracks}")
    print(f"Years: {year_range}")
    print("--------------------------------------\n")


if __name__ == "__main__":
    create_discography()