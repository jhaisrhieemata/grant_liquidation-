// === Code.gs ===
const TEMPLATE_ID = '1hdx0luh-MR9p4hccGcMCeMWn_02eQhgTWZIrEzlqnUE'; // Converted Google Doc ID 
const FOLDER_ID = '1-2plkS_X6qIjwISXAMU38lY31z1Qp_7d'; // optional: Google Drive folder for PDFs
const SHEET_FILE_ID = '1zHSS4R47kz7qXNIr7nlPn5l4AapzKzb15lvDkywfjqc';

function formatDateDMY(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}


function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Grant Liquidation Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saveAndCreatePdf(data) {
  if (!data || typeof data !== 'object') throw new Error('Invalid data received from form.');
  if (!Array.isArray(data.items)) data.items = [];
  if (!data.branch) throw new Error('Name/Branch');

  // Open spreadsheet
  let ss;
  try {
    ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  } catch (err) {
    ss = SpreadsheetApp.create('Grant Liquidation Database');
  }

  // Use the branch name as the sheet name
  const branchSheetName = data.branch.trim().toUpperCase();
  let sheet = ss.getSheetByName(branchSheetName);

  // If sheet for that branch doesn’t exist → create it
  if (!sheet) {
    sheet = ss.insertSheet(branchSheetName);
    sheet.appendRow([
      'Timestamp', 'Name/Branch', 'Date Created', 'Purpose for Budget', 'Amount of Withdrawal',
      'Receipt date', 'Description', 'Total Amount',
      'Total', 'Excess', 'PDF URL'
    ]);
  }

  const ts = new Date();
  const pdfUrl = createPdfFromTemplate(data, ts);

  // Append one row per item
  (data.items || []).forEach(it => {
    sheet.appendRow([
      ts,
      data.branch || '',
      formatDateDMY(data.date) || '',
      data.budget || '',
      data.withdrawal || '',
      formatDateDMY(it.tdate) || '',
      it.description || '',
      it.amount || '',
      data.total || '',
      data.excess || '',
      pdfUrl || ''
    ]);
  });

  return pdfUrl;
}


function createPdfFromTemplate(data, timestamp) {
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const copy = templateFile.makeCopy(
  `Liquidation ${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')}`
);

  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // === Replace placeholders ===
  body.replaceText('{{BRANCH}}', data.branch || '');
  body.replaceText('{{DATE}}', formatDateDMY(data.date) || '');
  body.replaceText('{{BUDGET}}', data.budget || '');
  body.replaceText('{{WITHDRAWAL}}', data.withdrawal || '');
  body.replaceText('{{TOTAL AMOUNT}}', data.total || '');
  body.replaceText('{{EXCESS}}', data.excess || '');

  // === Replace {{ITEMS}} placeholder with table ===
  const search = body.findText('{{ITEMS}}');
  if (search) {
    const element = search.getElement();
    const parent = element.getParent();

    // Build table data
    const tableData = [['Date', 'Description', 'Amount']];
    (data.items || []).forEach(it => {
      tableData.push([
         formatDateDMY(it.tdate) || '',
        it.description || '',
        it.amount || ''
      ]);
    });
    //tableData.push(['', '', '', 'Total', data.total || '']);

    // Insert table
    const table = body.insertTable(body.getChildIndex(parent) + 1, tableData);
    table.setBorderWidth(1);

    // === Style each cell safely ===
    for (let i = 0; i < table.getNumRows(); i++) {
  const row = table.getRow(i);
  for (let j = 0; j < row.getNumCells(); j++) {
    const cell = row.getCell(j);
    cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(5).setPaddingRight(5);

    // Set width
    if (j === 0) cell.setWidth(70);
    if (j === 1) cell.setWidth(400);
    if (j === 2) cell.setWidth(70);

    const text = cell.getChild(0);
    if (text && text.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = text.asParagraph();

      // NEW LOGIC: Priority alignment for the first column
      if (j === 0) {
        para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      } else if (i === 0) { 
        // Header row for other columns
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      } else if (j === 2) {
        // Specifically center the 3rd column (index 2) for data rows
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      } else {
        // Default for remaining data cells
        para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      }

      // Keep header styling separate
      if (i === 0) {
        para.setBold(false);
        cell.setBackgroundColor('#f0f0f0');
      }
    }
  }
}


    // Remove the placeholder
    element.asText().setText('');
  }

  doc.saveAndClose();

  // === Convert to PDF ===
  const pdfBlob = copy.getBlob().getAs('application/pdf');
  const pdfFile = FOLDER_ID
    ? DriveApp.getFolderById(FOLDER_ID).createFile(pdfBlob).setName(copy.getName() + '.pdf')
    : DriveApp.createFile(pdfBlob).setName(copy.getName() + '.pdf');

  try {
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {}

  // Remove temporary Google Doc
  DriveApp.getFileById(copy.getId()).setTrashed(true);

  return pdfFile.getUrl();
}
