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


interface Json{
  type: string,
  gmail?: string,
  bookid?: string;  
  filename?: string;
  sheetname: string;
  row: number,
  col: number,
  value: string
}

const mainbook = SpreadsheetApp.openById("1HtGLfZMu7TqxOQUIrGRfk8WS6saq-vMNhQg3tKrEQZk");

function doPost(data: PostEvent) {
  const json: Json = JSON.parse(data.postData.contents);

  if (json.type === "create"){
    const workbook = mainbook.copy(json.filename || "");
    workbook.addEditor(json.gmail || "");

    return ContentService.createTextOutput(`${workbook.getId()}:${workbook.getUrl()}`);  
  }

  const book = getWorkBook(data);
  const sheet = getSheet(data, book);

  if (json.type === "post"){
    sheet.getRange(json.row, json.col).setValue(json.value)
  
    return ContentService.createTextOutput("ok");  
  }

  else if (json.type === "get"){
    const value = sheet.getRange(json.row, json.col).getValue();

    return ContentService.createTextOutput(value);
  }

}


function getWorkBook(data: PostEvent){
  const json: Json = JSON.parse(data.postData.contents);
  let workbook: GoogleAppsScript.Spreadsheet.Spreadsheet;

  try{
      workbook = SpreadsheetApp.openById(json.bookid || "");
  }
  catch(e){
    workbook = mainbook.copy(json.filename || "");
    
  }
  return workbook;
}


function getSheet(data: PostEvent, book: Spreadsheet){
  const json: Json = JSON.parse(data.postData.contents);
  let sheet = book.getSheetByName(json.sheetname);

  if(!sheet){
      sheet = book.insertSheet();
      sheet.setName(json.sheetname);
  }
  return sheet;
}