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
        "Monitored the WhatsApp-based lorry assignment system during its initial live operation period following deployment. Tracked end-to-end message flow across the webhook server, Meta Cloud API, and file exchange pipeline to verify that inbound delivery order files were being processed correctly and that assignment outputs and trip manifests were being returned to users as expected under real daily workload conditions.",
        "Investigated post-deployment defects surfaced during live usage. Identified a root cause in the remarks field parser where date-based delivery restrictions were being incorrectly evaluated against the wrong reference date, causing valid orders to be filtered out before assignment. Also diagnosed a session state persistence issue where the daily assignment log was not being correctly retained between bot reconnection events, leading to duplicate assignments.",
        "Reviewed the full set of assignment rules against live operational data and cross-referenced system outputs with expected assignment results. Compiled a prioritised list of rule inconsistencies and unhandled edge cases across route boundary enforcement, capacity utilisation thresholds, and preferred-vehicle override behaviour to be addressed in the next development cycle.",
    ]),
    (2, [
        "Implemented a route-direction guard across all delivery order consolidation and overflow assignment passes. The guard uses bearing-based geographic classification to prevent orders destined for geographically opposing directions from being co-assigned to the same vehicle, resolving an engine defect that was producing multi-directional trip manifests with excessive backtracking.",
        "Fixed the vehicle capacity utilisation threshold logic to correctly enforce an 80% load fill target before triggering a capacity upgrade split. Corrected the fill-to-threshold pass so that partial loads from multiple orders are accumulated and consolidated up to the utilisation target before any overflow is allocated to a second vehicle, preventing unnecessary under-loaded dispatches.",
        "Added earliest-date-first ordering to all assignment passes so that delivery orders with earlier scheduled dates are always evaluated and assigned before more recent ones. This prevents newly uploaded orders from displacing older outstanding ones when total order weight approaches fleet capacity limits.",
        "Tightened the urban route merging guard to enforce corridor-group-only consolidation, so that orders belonging to different geographic corridor groups cannot be merged onto the same vehicle even when both are classified under the same broad route category. This closed a gap where orders from opposite ends of the urban delivery zone were being incorrectly co-assigned.",
        "Corrected the date selection flow in the WhatsApp button interface to include Saturday as a valid selectable working day while continuing to exclude Sunday, aligning the system's date logic with the actual six-day working week.",
    ]),
    (3, [
        "Refactored all assignment engine business-rule constants out of the core logic modules and into a dedicated centralised configuration file (assignment_config.py). Created a companion human-readable rules reference (ASSIGNMENT_RULES.md) that is loaded at every system startup, enabling operational rule parameters to be reviewed and updated independently of application code without requiring code changes or redeployment.",
        "Replaced hardcoded vehicle-to-route mapping tables with a data-driven preferred vehicle lookup that reads configuration from a master data file at runtime. Extended the delivery order remarks parser to extract free-text vehicle size cap instructions and normalise them to a standard tonnage constraint, handling variable natural-language phrasings through token classification rather than fixed pattern matching.",
        "Refactored the delivery order merge and consolidation logic to apply a three-level priority hierarchy: exact route match first, same-city grouping second, and same-state nearest-coordinate grouping third. Added a pre-merge pass that consolidates same-route-prefix sub-buckets before cross-route grouping is attempted, resolving a defect where orders sharing a common route prefix were being incorrectly split across separate vehicles.",
        "Integrated OSRM road-network distance routing using the OpenStreetMap dataset as the primary stop-sequencing method, with haversine great-circle distance as an automatic fallback when the routing service is unavailable. Replaced the previous single-axis coordinate sweep with a greedy nearest-neighbour algorithm that sequences trip manifest stops from the depot outward along actual road paths, reducing total estimated trip distance.",
        "Extended the trip manifest output to include a return-to-depot entry at the end of each vehicle's stop list and a round-trip estimated time summary. Fixed the stop-ordering algorithm to use a principal-axis geographic sweep that correctly sequences stops in the actual direction of travel from the depot, replacing a defective single-dimension coordinate sort that produced zigzag routing patterns.",
        "Resolved a state-compatibility exclusion defect in the no-eligible-vehicle resolution path where multi-state route buckets were incorrectly excluding compatible vehicles based on a misapplied state filter. Corrected the preferred-owner vehicle logic to enforce owner assignment as a hard constraint, with fallback to non-owner vehicles triggered only when all owner-assigned vehicles are confirmed at full capacity.",
        "Wrote and published a comprehensive operator and maintenance manual documenting the complete system architecture, WhatsApp interaction flow, assignment rule set, master data file formats and update procedures, external API dependencies, environment configuration, service restart commands, maintenance workflows, troubleshooting procedures, and quick-reference cards for both daily operators and the system maintainer.",
    ]),
    (4, [
        "Corrected route-specific preferred vehicle configuration entries where incorrect fallback vehicles were being selected during assignment. Enforced strict per-route vehicle reservation rules and validated the corrected configuration against historical assignment data to confirm that the expected vehicles are being selected consistently across different load scenarios.",
        "Fixed two vehicle-contamination defects that were causing vehicles already committed to one route class to be incorrectly considered available for a different route class. The first defect occurred during the primary assignment pass when route prefix overlap caused an outstation-committed vehicle to match an urban route. The second defect occurred during the overflow reassignment pass where already-committed vehicles were not being excluded from the available pool.",
        "Implemented automatic proportional splitting for oversized single delivery orders whose total weight exceeds the capacity of any single available vehicle. The engine now distributes such orders across multiple vehicles in proportion to their available remaining capacity, replacing the previous behaviour where such orders were left permanently unresolved in the output with a no-vehicle flag.",
    ]),
]

# ---- JULY DATA ----
july_weeks = [
    (1, [
        "Conducted user acceptance testing (UAT) by executing the system against a set of historical daily delivery order files to validate end-to-end assignment correctness and trip manifest accuracy. Documented all cases where system output diverged from expected results, categorised each discrepancy by root cause, and produced a structured defect list for targeted rule and engine refinement.",
        "Analysed recurring unresolved assignment failure patterns collected from both UAT runs and live operational data to identify the specific order weight and route combinations that fall outside the current fleet capacity model. Investigated the engine's fallback resolution paths and began implementing an extended fallback strategy to handle these cases through alternative vehicle selection logic rather than leaving orders unassigned.",
        "Extended the delivery remarks parser to correctly handle a wider range of free-text restriction phrase formats encountered in real data, including bilingual entries mixing multiple languages within the same field. Improved the negation detection logic to correctly identify and scope compound restriction clauses, preventing misclassification of negated delivery day restrictions.",
    ]),
    (2, [
        "Developed an end-of-session summary generation feature that compiles a structured report at the close of each assignment session, covering total orders processed, vehicles assigned, load utilisation percentage per vehicle, and any unresolved items remaining. The summary is formatted and delivered via the existing WhatsApp messaging interface, providing an immediate overview of the day's assignment results without requiring the user to open exported files.",
        "Prepared system handover documentation covering the full codebase structure, production server deployment configuration, tunnel service management, external API credential renewal procedures, and step-by-step recovery instructions for common failure scenarios. The documentation is written to enable routine system maintenance and restart to be carried out independently by non-developer staff following the end of the internship period.",
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
