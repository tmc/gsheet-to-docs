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
  const documentTemplateId = "1W2PHmWl6kI1GnzkVusZ5ihbn5gZRA5IePqPCfmK9fes";
  const outputFolderId = "1_WHSd9IX1ysYuw6sI2R2eQOJAo7k8jvT";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const folder = DriveApp.getFolderById(outputFolderId);

  const templateDoc = DocumentApp.openById(documentTemplateId);
  const destinationFile = DriveApp.getFileById(documentTemplateId).makeCopy();
  destinationFile.moveTo(folder);
  const doc = DocumentApp.openById(destinationFile.getId());

  const sheetData = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getValues();

  const headers = sheetData[0];

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

    newContent.getParagraphs().forEach(function (p) {
      body.appendParagraph(p.copy());
    });

    // body.appendPageBreak();
    body.appendHorizontalRule();
  }
}
