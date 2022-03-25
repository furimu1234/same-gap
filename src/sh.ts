import SpreadsheetApp = GoogleAppsScript.Spreadsheet.SpreadsheetApp

import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet

interface PostEvent {
  queryString?: string;
  parameter?: { [index: string]: string; };
  parameters?: { [index: string]: [string]; };
  contentLenth: number;
  postData: {
    length: number;
    type: string;
    contents: string;
    name: string;
  };
}


interface Json {
  type: string,
  gmail?: string,
  bookid?: string;
  filename?: string;
  sheetname: string;
  range: string;
  value: string
}

const mainbook = SpreadsheetApp.openById("1qUtA-TNdsDcST7kb17VqOrWbiEOuSUeRMaq4iYZvKtM");

function doGet() {
  let workbook = SpreadsheetApp.openById("754680802578792498");
  let sheet = workbook.getSheetByName("基本設定") || workbook.insertSheet();
  let values = sheet.getRange("B5:B6").getValues();

  return ContentService.createTextOutput(values);
}


function doPost(data: PostEvent) {
  const json: Json = JSON.parse(data.postData.contents);

  if (json.type === "create") {
    const workbook = mainbook.copy(json.filename || "");
    workbook.addEditor(json.gmail || "");

    return ContentService.createTextOutput(`${workbook.getId()}:${workbook.getUrl()}`);
  }

  const book = getWorkBook(data);

  if (json.type == "invite") {
    book.addEditor(json.gmail || "");
    return ContentService.createTextOutput(`${book.getUrl()}`);
  }

  const sheet = getSheet(data, book);


  if (json.type === "post") {
    sheet.getRange(json.range).setValue(json.value)

    return ContentService.createTextOutput("ok");
  }

  else if (json.type === "get") {
    const value = sheet.getRange(json.range).getValues();

    return ContentService.createTextOutput(value);
  }

}

function getWorkBook(data: PostEvent) {
  const json: Json = JSON.parse(data.postData.contents);
  let workbook: GoogleAppsScript.Spreadsheet.Spreadsheet;

  try {
    workbook = SpreadsheetApp.openById(json.bookid || "");
  }
  catch (e) {
    let now = new Date();

    let month: number = now.getMonth();
    let day: number = now.getDay();
    let hour: number = now.getHours();
    let minute: number = now.getMinutes();
    let second: number = now.getSeconds();

    let error_date = `${month}/${day} ${hour}:${minute}:${second}`;

    workbook = mainbook.copy(json.filename || `error_date ${error_date}`);

  }
  return workbook;
}


function getSheet(data: PostEvent, book: Spreadsheet) {
  const json: Json = JSON.parse(data.postData.contents);
  let sheet = book.getSheetByName(json.sheetname);

  if (!sheet) {
    sheet = book.insertSheet();
    sheet.setName(json.sheetname);
  }
  return sheet;
}