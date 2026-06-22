from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        val = kwargs.get(edge, 'single')
        sz = kwargs.get(f'{edge}_sz', 4)
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), val)
        tag.set(qn('w:sz'), str(sz))
        tag.set(qn('w:space'), '0')
        tag.set(qn('w:color'), '000000')
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def set_cell_bg(cell, color_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def add_bullet(cell, text, font_size=10):
    p = cell.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.size = Pt(font_size)
    run.font.name = 'Times New Roman'
    return p

def make_activity_table(doc, weeks_data):
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    # Header row
    hdr = table.rows[0].cells
    hdr[0].width = Inches(0.8)
    hdr[1].width = Inches(5.5)
    for cell in hdr:
        set_cell_bg(cell, 'D3D3D3')
    p0 = hdr[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run0 = p0.add_run('Week')
    run0.bold = True
    run0.font.size = Pt(11)
    run0.font.name = 'Times New Roman'
    p1 = hdr[1].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p1.add_run('Projects / Activities')
    run1.bold = True
    run1.font.size = Pt(11)
    run1.font.name = 'Times New Roman'

    for week_num, bullets in weeks_data:
        row = table.add_row()
        row.cells[0].width = Inches(0.8)
        row.cells[1].width = Inches(5.5)
        # Week number cell
        wc = row.cells[0]
        wc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        wp = wc.paragraphs[0]
        wp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        wr = wp.add_run(str(week_num))
        wr.bold = True
        wr.font.size = Pt(11)
        wr.font.name = 'Times New Roman'
        # Activities cell - clear default paragraph first
        ac = row.cells[1]
        ac.paragraphs[0]._element.getparent().remove(ac.paragraphs[0]._element)
        for bullet_text in bullets:
            add_bullet(ac, bullet_text)

    return table

def build_report(doc, month_year, weeks_data, trainee_date, supervisor_date):
    # Header
    for line, bold, size in [
        ('Tunku Abdul Rahman University of', True, 13),
        ('Management and Technology', True, 13),
        ('Faculty of Computing and Information Technology', True, 13),
        ('Industrial Training Progress Report', True, 13),
        ('Activity Log', False, 12),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(line)
        r.bold = bold
        r.font.size = Pt(size)
        r.font.name = 'Times New Roman'

    doc.add_paragraph()

    # Info fields
    for label, value in [
        ('Name of Trainee:', 'OOI YU ZHEN'),
        ('Name of Company:', 'Eng Sheng Sdn. Bhd.'),
        ('Month/Year', month_year),
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        r_label = p.add_run(label)
        r_label.font.size = Pt(11)
        r_label.font.name = 'Times New Roman'
        p.add_run('\t\t')
        r_val = p.add_run(value)
        r_val.font.size = Pt(11)
        r_val.font.name = 'Times New Roman'

    doc.add_paragraph()

    # Activity table
    make_activity_table(doc, weeks_data)

    doc.add_paragraph()

    # Suggestions box
    p = doc.add_paragraph()
    r = p.add_run('Suggestions / Comments / Additional information (if any):')
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = 'Times New Roman'

    sug_table = doc.add_table(rows=1, cols=1)
    sug_table.style = 'Table Grid'
    sug_row = sug_table.rows[0]
    sug_row.height = Cm(2)
    sug_row.cells[0].text = ''

    doc.add_paragraph()

    # Leave section
    p = doc.add_paragraph()
    r = p.add_run('Leave Application / Leave Taken')
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = 'Times New Roman'

    for line in [
        '1.From (dd/mm/yyyy): _____________ to (dd/mm/yyyy) _____________ (  day(s))',
        '2. Reasons for taking leave: ________________________________________________________________',
        '3. Total number of days taken: ________________________________________________________________',
    ]:
        p = doc.add_paragraph(line)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        for run in p.runs:
            run.font.size = Pt(11)
            run.font.name = 'Times New Roman'

    doc.add_paragraph()

    p = doc.add_paragraph()
    r = p.add_run('I hereby declare that the information given above is correct.')
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = 'Times New Roman'

    doc.add_paragraph()

    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.style = 'Table Grid'
    sig_table.autofit = False
    lc = sig_table.rows[0].cells[0]
    rc = sig_table.rows[0].cells[1]
    lc.width = Inches(3)
    rc.width = Inches(3)
    lp = lc.paragraphs[0]
    lr = lp.add_run('Signature: _______________________')
    lr.font.size = Pt(11)
    lr.font.name = 'Times New Roman'
    rp = rc.paragraphs[0]
    rr = rp.add_run(f'Date (dd/mm/yyyy):   {trainee_date}')
    rr.font.size = Pt(11)
    rr.font.name = 'Times New Roman'
    # Remove borders
    for cell in [lc, rc]:
        set_cell_border(cell, top='none', left='none', bottom='none', right='none')

    doc.add_paragraph()
    doc.add_paragraph()

    # Supervisor section
    p = doc.add_paragraph()
    r = p.add_run('Endorsement by the Company Supervisor:')
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = 'Times New Roman'

    doc.add_paragraph()

    p = doc.add_paragraph()
    r = p.add_run('The above is a true record of activities taken by the trainee in the captioned week.')
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = 'Times New Roman'

    doc.add_paragraph()

    sup_table = doc.add_table(rows=5, cols=2)
    sup_table.style = 'Table Grid'
    sup_data = [
        ('Signature of Supervisor:', ''),
        ('Name of Supervisor:', 'Ardyawati Binti Zainon'),
        ('Date (dd/mm/yyyy):', supervisor_date),
        ('Email:', 'ardyawati@engsheng.com'),
        ('Mobile / Office Contact No.:', '016-5227316'),
    ]
    for i, (label, val) in enumerate(sup_data):
        lc = sup_table.rows[i].cells[0]
        rc = sup_table.rows[i].cells[1]
        lc.width = Inches(2.5)
        rc.width = Inches(3.5)
        if i == 0:
            lc.paragraphs[0].paragraph_format.space_before = Pt(20)
            lc.paragraphs[0].paragraph_format.space_after = Pt(20)
            rc.paragraphs[0].paragraph_format.space_before = Pt(20)
            rc.paragraphs[0].paragraph_format.space_after = Pt(20)
        lp = lc.paragraphs[0]
        lr = lp.add_run(label)
        lr.font.size = Pt(11)
        lr.font.name = 'Times New Roman'
        rp = rc.paragraphs[0]
        rr = rp.add_run(val)
        rr.font.size = Pt(11)
        rr.font.name = 'Times New Roman'

    doc.add_paragraph()

    p = doc.add_paragraph()
    r = p.add_run('Company Stamp:')
    r.font.size = Pt(11)
    r.font.name = 'Times New Roman'


# ---- JUNE DATA ----
june_weeks = [
    (1, [
        "Monitored the deployed WhatsApp-based lorry assignment system during its initial live operation period. Verified end-to-end message flow, file exchange, and assignment output correctness under real daily usage conditions.",
        "Investigated and identified root causes for post-deployment defects surfaced during live usage, including incorrect date-based filtering in the remarks parser and session state persistence issues in the daily assignment log.",
        "Reviewed the full assignment rule set against live operational data and documented a prioritised list of rule inconsistencies and edge cases for the next development iteration.",
    ]),
    (2, [
        "Implemented a route-direction guard across all assignment and overflow passes to prevent geographically opposing delivery orders from being co-assigned to the same vehicle, resolving incorrect multi-directional grouping.",
        "Fixed the capacity utilisation threshold logic to correctly enforce an 80% load target before triggering a vehicle upgrade split, and corrected the fill pass to consolidate partial loads up to the threshold before allocating overflow to a second unit.",
        "Added earliest-date-first ordering across all assignment passes to ensure older outstanding orders are prioritised over newer ones when fleet capacity is limited.",
        "Tightened the urban route guard to enforce corridor-group-only merging, blocking cross-corridor co-assignment even between routes of the same classification.",
        "Fixed the date selection flow to correctly include Saturday as a working day while excluding Sunday only.",
    ]),
    (3, [
        "Refactored all assignment business-rule constants into a centralised configuration module (assignment_config.py) and created a companion rules reference file (ASSIGNMENT_RULES.md) loaded at system startup, separating operational parameters from engine logic to enable rule updates without code changes.",
        "Replaced hardcoded lorry-to-route mappings with a data-driven preferred lorry lookup table, and added a free-text tonnage cap parser to the remarks field supporting variable natural-language phrasings.",
        "Refactored the DO merge logic to apply a three-level priority: route match first, same-city consolidation second, and same-state nearest-longitude grouping third. Added a pre-merge pass for same-route-prefix sub-buckets to resolve incorrect route splitting.",
        "Integrated OSRM road-network routing as the primary stop-sequencing method with haversine as fallback, and replaced the latitude-sweep ordering with a greedy nearest-neighbour algorithm to minimise total trip distance in the generated manifests.",
        "Added a return-to-depot row and round-trip ETA summary to trip manifests, and fixed stop-ordering to use a principal-axis geographic sweep instead of single-dimension coordinate sorting.",
        "Resolved a state-compatibility exclusion bug causing eligible lorries to be incorrectly filtered on multi-state route buckets, and corrected preferred-owner logic to treat owner assignment as a hard constraint with fallback only on full capacity.",
        "Wrote and published a comprehensive operator and maintenance manual covering system architecture, assignment rules, data file formats, deployment configuration, restart procedures, and troubleshooting reference.",
    ]),
    (4, [
        "Corrected route-specific preferred lorry configurations where incorrect fallback lorries were being selected, ensuring strict assignment rules are enforced per route.",
        "Fixed two vehicle-contamination bugs where lorries already committed to one route class were incorrectly considered available for a different class during both primary and overflow assignment passes.",
        "Implemented automatic proportional splitting of oversized single delivery orders that exceed the maximum capacity of any available vehicle, replacing the previous unresolvable NO_LORRY outcome with a multi-vehicle assignment.",
    ]),
]

# ---- JULY DATA ----
july_weeks = [
    (1, [
        "Conducted user acceptance testing (UAT) by running the system against historical daily data to validate assignment accuracy and manifest correctness. Documented and categorised output discrepancies by root cause for targeted rule refinement.",
        "Analysed recurring NO_LORRY and NO_ELIGIBLE_LORRY failure patterns from live and UAT data to identify weight and route combinations outside the current fleet capacity model, and began implementing an extended fallback resolution strategy.",
        "Extended the remarks parser to handle a wider set of free-text delivery restriction phrases, including bilingual entries, and improved negation detection to correctly scope compound restriction clauses.",
    ]),
    (2, [
        "Developed an end-of-session summary generation feature that produces a structured text report of total orders assigned, vehicles utilised, average load percentage, and unresolved items, delivered via WhatsApp at session close.",
        "Prepared system handover documentation covering codebase structure, server deployment configuration, tunnel management, API token renewal, and common failure recovery procedures to support independent operation after the internship period.",
    ]),
]

doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)

# ---- JUNE REPORT ----
build_report(doc, 'JUNE/2026', june_weeks, '30/6/2026', '30/06/2026')

doc.add_page_break()

# ---- JULY REPORT ----
build_report(doc, 'JULY/2026', july_weeks, '14/7/2026', '14/07/2026')

doc.save('/home/user/DO_bot/June_July_2026_REPORT_OOI_YU_ZHEN.docx')
print("Done.")
