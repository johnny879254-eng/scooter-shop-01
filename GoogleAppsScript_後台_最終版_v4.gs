const SHEET_NAME = "代步車借出紀錄";
const PDF_FOLDER_NAME = "代步車借出完整存證PDF";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["系統收件時間","登記日期","原車車牌","代步車車牌","取車店家","條款確認","送出時間","完整PDF存證連結"]);
      sheet.setFrozenRows(1);
    }

    const folder = getOrCreateFolder_(PDF_FOLDER_NAME);
    const stamp = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmmss");
    const plate = safe_(data["原車車牌"] || "未填車牌");
    const pdfFile = createCompletePdf_(folder, data, stamp + "_" + plate + "_代步車借出登記存證.pdf");

    sheet.appendRow([new Date(),data["登記日期"]||"",data["原車車牌"]||"",data["代步車車牌"]||"",data["取車店家"]||"",data["條款確認"]||"",data["送出時間"]||"",pdfFile.getUrl()]);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function createCompletePdf_(folder, data, pdfName) {
  const doc = DocumentApp.create("TEMP_" + new Date().getTime());
  const body = doc.getBody();
  body.clear();

  let p = body.appendParagraph("代步車借出登記存證單");
  p.setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const info = [
    ["登記日期", data["登記日期"] || ""],
    ["原車車牌", data["原車車牌"] || ""],
    ["代步車車牌", data["代步車車牌"] || ""],
    ["取車店家", data["取車店家"] || ""],
    ["條款確認", data["條款確認"] || ""],
    ["送出時間", data["送出時間"] || ""]
  ];
  const infoTable = body.appendTable(info);
  for (let r=0;r<infoTable.getNumRows();r++) {
    infoTable.getRow(r).getCell(0).setBackgroundColor("#f3f4f6");
    infoTable.getRow(r).getCell(0).editAsText().setBold(true);
  }

  body.appendParagraph("");
  body.appendParagraph("借用規範確認內容").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  (data["條款內容"] || []).forEach(function(t, i) {
    body.appendParagraph("☑ " + (i+1) + ". " + t);
  });

  body.appendPageBreak();
  body.appendParagraph("駕照正面照片").setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  appendDataUrlImage_(body, data["駕照照片"], 430, 300);

  body.appendParagraph("");
  body.appendParagraph("本人簽名").setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  appendDataUrlImage_(body, data["本人簽名"], 430, 180);

  body.appendParagraph("");
  body.appendParagraph("本文件由代步車借出登記系統自動產生，作為店家內部借用存證使用。")
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  doc.saveAndClose();
  Utilities.sleep(1000);

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs(MimeType.PDF).setName(pdfName);
  const pdfFile = folder.createFile(pdfBlob);
  docFile.setTrashed(true);
  return pdfFile;
}

function appendDataUrlImage_(body, dataUrl, maxW, maxH) {
  if (!dataUrl) {
    body.appendParagraph("未提供圖片");
    return;
  }
  const parts = dataUrl.split(",");
  const mime = (parts[0].match(/data:(.*);base64/) || [,"image/png"])[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, "image");
  const img = body.appendImage(blob);
  const w = img.getWidth(), h = img.getHeight();
  const scale = Math.min(maxW / w, maxH / h, 1);
  img.setWidth(Math.round(w * scale)).setHeight(Math.round(h * scale));
  img.getParent().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function safe_(s) {
  return String(s).replace(/[\\/:*?"<>|]/g, "_");
}