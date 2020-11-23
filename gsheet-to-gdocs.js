function onOpen() {
  const menuEntries = [
    {
      name: "Render Template for Sheet",
      functionName: "RenderGdocTemplate",
    },
  ];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Templating", menuEntries);
}

function RenderGdocTemplate() {
  const documentTemplateId = "1RCaTCZqIPAnM2vrW1z5PfAD_4M8CHN3SPIcQModKcP4";
  const outputFolderId = "12wA6fsRFeTxKKifRYXNd9xIUmWSripAH";
  const startRow = 1;
  const endRow = undefined;

  return renderGdocTemplate(documentTemplateId, outputFolderId, startRow, endRow);
}

function renderGdocTemplate(documentTemplateId, outputFolderId, startRow, lastRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const folder = DriveApp.getFolderById(outputFolderId);

  const templateDoc = DocumentApp.openById(documentTemplateId);
  const destinationFile = DriveApp.getFileById(documentTemplateId).makeCopy();
  destinationFile.moveTo(folder);
  const doc = DocumentApp.openById(destinationFile.getId());

  // Assumes field names are in the first row.
  const lastRow = endRow || sheet.getLastRow()-1;
  const sheetData = sheet.getRange(startRow, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = sheetData[0];

  //var section = doc.getActiveSection();
  const tmpl = templateDoc.getBody();

  const body = doc.getBody();
  body.clear();

  for (const i in sheetData) {
    const newContent = tmpl.copy();
    const text = newContent.editAsText();

    const context = {};
    headers.forEach(function (h, j) {
      context[h] = sheetData[i][j];
      text.replaceText("{{" + h + "}}", sheetData[i][j]);
    });

    const nChildren = newContent.getNumChildren();
    for (let i = 0; i < nChildren; i++) {
      const child = newContent.getChild(i).copy();
      switch(child.getType()) {
        case DocumentApp.ElementType.PARAGRAPH:
          body.appendParagraph(child.asParagraph());
          break;
        case DocumentApp.ElementType.TABLE:
          body.appendTable(child.asTable());
          break;
        case DocumentApp.ElementType.LIST_ITEM:
          body.appendListItem(child.asListItem());
          break;
        case DocumentApp.ElementType.INLINE_IMAGE:
          body.appendImage(child.asInlineImage());
          break
        default:
          body.appendParagraph(child.getText());
          break
      }
    }

    // body.appendPageBreak();
    body.appendHorizontalRule();
  }
}
