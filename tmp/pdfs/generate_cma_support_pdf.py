from reportlab.lib.pagesizes import LETTER
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, ListFlowable, ListItem

OUT_PATH = r"output/pdf/CMA_File_Watcher_Support_Guide.pdf"

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='TitleCustom', parent=styles['Title'], fontSize=20, leading=24, spaceAfter=14))
styles.add(ParagraphStyle(name='H1', parent=styles['Heading1'], fontSize=14, leading=18, spaceBefore=10, spaceAfter=8))
styles.add(ParagraphStyle(name='H2', parent=styles['Heading2'], fontSize=11.5, leading=14, spaceBefore=8, spaceAfter=6))
styles.add(ParagraphStyle(name='Body', parent=styles['BodyText'], fontSize=9.5, leading=13, spaceAfter=6))
styles.add(ParagraphStyle(name='Small', parent=styles['BodyText'], fontSize=8.5, leading=11, textColor=colors.HexColor('#444444')))


def p(text, style='Body'):
    return Paragraph(text, styles[style])

cell_style = ParagraphStyle(name='Cell', parent=styles['BodyText'], fontSize=8.1, leading=10)
head_style = ParagraphStyle(name='HeadCell', parent=styles['BodyText'], fontSize=8.2, leading=10, textColor=colors.white)

def _as_cell(val, header=False):
    if isinstance(val, Paragraph):
        return val
    s = '' if val is None else str(val)
    return Paragraph(s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'), head_style if header else cell_style)


def bullet(items):
    return ListFlowable(
        [ListItem(Paragraph(i, styles['Body']), leftIndent=12) for i in items],
        bulletType='bullet',
        start='circle',
        leftIndent=18,
        bulletFontSize=8,
    )


def table(data, col_widths):
    cooked = []
    for i, row in enumerate(data):
        cooked.append([_as_cell(c, header=(i == 0)) for c in row])
    t = Table(cooked, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4e78')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#8ea9c1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f4f8fb')]),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    return t


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont('Helvetica', 8)
    canvas.setFillColor(colors.HexColor('#555555'))
    canvas.drawString(0.65 * inch, 0.45 * inch, 'CMA File Watcher Support Guide')
    canvas.drawRightString(7.85 * inch, 0.45 * inch, f'Page {doc.page}')
    canvas.restoreState()


doc = SimpleDocTemplate(
    OUT_PATH,
    pagesize=LETTER,
    leftMargin=0.65 * inch,
    rightMargin=0.65 * inch,
    topMargin=0.7 * inch,
    bottomMargin=0.65 * inch,
)

story = []

story.append(p('CMA File Watcher Service - Support Guide', 'TitleCustom'))
story.append(p('Audience: New SalesOps or IT support staff responsible for CMA intake and PCF creation support.', 'Small'))
story.append(p('Scope: How the service works, how to install/deploy it, and what inbound file structure is required for successful processing.', 'Small'))
story.append(Spacer(1, 8))

story.append(p('1) What This Program Does', 'H1'))
story.append(p('The Windows service watches the inbound CMA share for new Excel files. For each file, it reads key header fields and item rows, writes records to BAT_App.dbo.Chap_CmaItems, validates business rules, runs stored procedures to create PCFs, emails the submitter, archives the workbook, and removes the original from the inbound folder.', 'Body'))

story.append(p('Main flow from import to PCF creation', 'H2'))
story.append(bullet([
    'User submits CMA workbook to \\ciiedi01\\SendDocs\\CMAInbound (file extension .xlsx).',
    'FileSystemWatcher detects file creation and enqueues it (single file processing queue).',
    'Service waits until the file is unlocked and non-empty.',
    'Service reads worksheet data (prefers sheet name "CMA Template", then "CMA", then first sheet).',
    'Service inserts one Chap_CmaItems row per item line in the workbook.',
    'CMAValidator checks required fields, date rules, duplicate conflicts, and item validity.',
    'If valid, service runs CreatePcfFromChapCmaItems_sp (and sp_ArchiveReplacedPCF for replacements).',
    'If PCF number exists after procedure, success email is sent and file is archived to Processed.',
    'If validation or creation fails, failure email is sent (validation failures) and file is archived to Rejected.',
    'Original inbound workbook is deleted only after archive copy exists in Processed or Rejected.'
]))

story.append(Spacer(1, 6))
story.append(p('Service dependencies and integration points', 'H2'))
story.append(table([
    ['Component', 'Used For'],
    ['Windows Service (CmaFileWatcherService.exe)', 'Continuous file monitoring and orchestration'],
    ['Inbound Share (CMAInbound)', 'Receives CMA workbook files'],
    ['SQL Server (ciisql10 BAT_App)', 'Staging records in Chap_CmaItems and email dispatch via Database Mail'],
    ['Linked server (CiiSQL01 + PCF DB)', 'Existing PCF conflict checks against ProgControl/pcitems'],
    ['Stored Procedures', 'CreatePcfFromChapCmaItems_sp and sp_ArchiveReplacedPCF drive PCF lifecycle'],
    ['Database Mail profile SalesOps', 'Sends success/failure notifications to submitter email'],
], [2.25 * inch, 4.85 * inch]))

story.append(PageBreak())

story.append(p('2) Installation and Deployment', 'H1'))
story.append(p('Build prerequisites', 'H2'))
story.append(bullet([
    'Windows host capable of running a .NET Framework 4.7.2 service.',
    'Visual Studio/MSBuild with NuGet restore support.',
    'Access to referenced package folder structure (../packages) from solution path.',
    'Network access and permissions for inbound share and SQL Server ciisql10.'
]))

story.append(p('Build steps (from repository root)', 'H2'))
story.append(bullet([
    'Restore NuGet packages (if not already restored).',
    'Build solution CmaFileWatcherSvc.sln (Debug or Release).',
    'Verify output contains CmaFileWatcherService.exe and dependent DLLs in CmaFileWatcherService/bin/<Config>.'
]))

story.append(p('Deploy to service server (example from code comments)', 'H2'))
story.append(bullet([
    'Copy compiled output to target folder (example: C:\\CmaFileWatcher).',
    'Stop existing service: sc stop CmaFileWatcherService',
    'Delete existing registration: sc delete CmaFileWatcherService',
    'Create service: sc create CmaFileWatcherService binPath= "C:\\CmaFileWatcher\\CmaFileWatcherService.exe" start= auto',
    'Start service: sc start CmaFileWatcherService',
    'Check status: sc query CmaFileWatcherService'
]))

story.append(p('Runtime configuration and controls', 'H2'))
story.append(table([
    ['Setting/Control', 'Behavior'],
    ['config file path', '\\ciiedi01\\SendDocs\\CMAInbound\\config.ini'],
    ['PcfDatabase (INI)', 'Used by validator when checking linked server PCF conflicts'],
    ['Watch folder', 'Hardcoded to \\ciiedi01\\SendDocs\\CMAInbound in service code'],
    ['Watched file type', '.xlsx only'],
    ['Debug logging switch', 'Set environment variable DebugCMAWatcher=true'],
    ['Logs', 'CmaFileWatcherService.log always; debug.log only when debug switch enabled'],
], [2.1 * inch, 5.0 * inch]))

story.append(p('Security and operational notes', 'H2'))
story.append(bullet([
    'Current code contains direct SQL credentials in connection strings; treat binaries and source as sensitive.',
    'Service account must have read/write/create/delete rights on inbound, Processed, and Rejected folders.',
    'Service account must execute required SQL queries/stored procedures and send dbmail via profile SalesOps.',
    'If file remains locked longer than retry window, it is skipped and logged as inaccessible.'
]))

story.append(PageBreak())

story.append(p('3) Required Inbound Workbook Structure (what support should validate)', 'H1'))
story.append(p('Important clarification: this service processes Excel workbooks (.xlsx), not PDF files directly. In practice, users may generate this workbook from their CMA process/template, but the watched input is an Excel file.', 'Body'))

story.append(p('Expected file naming', 'H2'))
story.append(bullet([
    'Extension must be .xlsx.',
    'File name should include "_sentby_<email>" so submitter email can be extracted.',
    'Accepted sender domains in filename parser: @chapinmfg.com or @chapinusa.com.',
    'Example: CUSTOMER123_sentby_user@chapinmfg.com.xlsx'
]))

story.append(p('Worksheet selection', 'H2'))
story.append(bullet([
    'First choice: sheet named "CMA Template".',
    'Second choice: sheet named "CMA".',
    'Fallback: first worksheet in workbook.'
]))

story.append(p('Header fields the service reads', 'H2'))
story.append(table([
    ['Field', 'Cell Logic Used by Service', 'Purpose'],
    ['Existing PCF (replacement check)', 'A2', 'If numeric, request is treated as replacement CMA'],
    ['PCF Type heading row', 'Search B1:B10 for "PCF Type"', 'Determines placement of promo fields'],
    ['Start date anchor row', 'Search P1:P10 for first parseable date', 'Base row for customer/date offsets'],
    ['Promo Terms', 'F5 (or F2 when start row is 1)', 'Required for PD/PW types'],
    ['Promo Freight Terms', 'C5 (or C2 when start row is 1)', 'Required for PD/PW types'],
    ['Promo Freight Minimums', 'D5 (or D2 when start row is 1)', 'Required for PD/PW types'],
    ['Promo Freight Other Amount', 'D6 (or D3 when start row is 1)', 'Optional value captured to database'],
    ['Customer/Corp fields', 'B row offsets from detected start row', 'Used to derive Cust_num, Corp_flag, archive name'],
    ['Buying Group', 'K(startRow+2)', 'If value is #N/A it is converted to blank'],
    ['Submitted By', 'M(startRow+2)', 'Stored in Chap_CmaItems'],
    ['Start Date', 'P(startRow)', 'Stored in Chap_CmaItems.StartDate'],
    ['End Date', 'P(startRow+2)', 'Stored in Chap_CmaItems.EndDate'],
    ['General Notes', 'If A1 is "General Notes", uses B1', 'Stored as GenNotes'],
], [1.95 * inch, 2.15 * inch, 2.95 * inch]))

story.append(Spacer(1, 6))
story.append(p('Item detail rows', 'H2'))
story.append(bullet([
    'Item parsing starts at row (startRow + 5).',
    'Loop continues while column B (Item #) is non-empty.',
    'Per row mapping: A=Description, B=Item #, D=Sell Price (non-numeric defaults to 0).',
    'One database row is inserted per item into Chap_CmaItems with shared header values.'
]))

story.append(p('Example from sample workbook in repository', 'H2'))
story.append(bullet([
    'Row 10 contains customer number and P-column label for End Date in the sample.',
    'Row 12 begins item grid headers (Description, Item #, cost/sell columns).',
    'Actual item records start at row 13 and continue down column B.'
]))

story.append(PageBreak())

story.append(p('4) Validation Rules and Failure Behavior', 'H1'))
story.append(p('Validation checks executed by CMAValidator', 'H2'))
story.append(bullet([
    'Required checks on first record: PcfTypeText, StartDate, EndDate.',
    'For PcfType starting PD or PW: Promo Terms, Promo Freight Terms, and Promo Freight Minimums are required.',
    'Date integrity: StartDate cannot be after EndDate.',
    'Conflict check: same customer, same start date, same item cannot overlap an active PCF record (status 3, not 98), excluding the replaced PCF when applicable.',
    'Item check: each item must exist in item_mst and not be obsolete (stat != O).'
]))

story.append(p('On validation failure', 'H2'))
story.append(bullet([
    'Rows for the CMA are marked Status = F when PCFNumber is null and status was N.',
    'Failure email sent to parsed submitter email with list of validation errors.',
    'Workbook copy saved under Rejected folder with generated archive name.'
]))

story.append(p('On success', 'H2'))
story.append(bullet([
    'Optional replacement path archives old PCF first (sp_ArchiveReplacedPCF).',
    'PCF creation procedure executed (CreatePcfFromChapCmaItems_sp).',
    'If PCF number is populated, success email is sent and workbook saved under Processed.',
    'Service writes returned PCF number into cell A2 in archived workbook copy.'
]))

story.append(p('Archive naming behavior', 'H2'))
story.append(bullet([
    'archiveName = <corpOrCust>_<yyyyMMddHHmmss>_<baseFilename>',
    'baseFilename removes "_sentby_<email>" part but preserves extension.'
]))

story.append(p('5) Support Runbook: Fast Triage Checklist', 'H1'))
story.append(table([
    ['Symptom', 'What to check first'],
    ['No processing started', 'Service status, watch path availability, file extension is .xlsx, file not locked.'],
    ['File processed but rejected', 'CmaFileWatcherService.log + failure email details + template required fields.'],
    ['No success/failure email', 'Filename contains parseable sender email/domain and dbmail profile SalesOps health.'],
    ['PCF not created', 'Stored procedure execution rights, SQL errors, Chap_CmaItems rows and validation state.'],
    ['Original file still in inbound', 'Whether archive copy exists; deletion happens only after Processed/Rejected write.'],
], [2.1 * inch, 5.0 * inch]))

story.append(Spacer(1, 6))
story.append(p('Key SQL objects referenced by this service', 'H2'))
story.append(bullet([
    'Tables: Chap_CmaItems, item_mst, linked ProgControl, linked pcitems.',
    'Procedures: CreatePcfFromChapCmaItems_sp, sp_ArchiveReplacedPCF.',
    'Mail proc: msdb.dbo.sp_send_dbmail (profile SalesOps).'
]))

story.append(Spacer(1, 6))
story.append(p('Document generated from source review of CmaFileWatcherService project on 2026-03-05.', 'Small'))


doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
print(OUT_PATH)
