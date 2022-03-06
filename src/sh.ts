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

interface GetEvent {
  queryString?: string;
  parameter?: { [index: string]: string; };
  parameters?: { [index: string]: [string]; };
  contentLenth: number;
  getData: {
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

interface getJson{
  bookid?: string;  
  filename?: string;
  sheetname: string;
}


function doPost(data: PostEvent) {
  var json: Json = JSON.parse(data.postData.contents);

  if (json.type === "create"){
    const mainbook = SpreadsheetApp.openById("1HtGLfZMu7TqxOQUIrGRfk8WS6saq-vMNhQg3tKrEQZk");
    let workbook = mainbook.copy(json.filename || "");
    workbook.addEditor(json.gmail || "");

    return ContentService.createTextOutput(`${workbook.getId()}:${workbook.getUrl()}`);  
  }


  else if (json.type === "post"){
    var book = getWorkBook(data);

    var sheet = getSheet(data, book);
  
    sheet.getRange(json.row, json.col).setValue(json.value)
  
    return ContentService.createTextOutput("ok");  
  }

  else if (json.type === "get"){

    var book = getWorkBook(data);

    var sheet = getSheet(data, book);

    var value = sheet.getRange(json.row, json.col).getValue();

    return ContentService.createTextOutput(value);
  }

}


function getWorkBook(data: PostEvent){
  var json: Json = JSON.parse(data.postData.contents);
  var workbook: GoogleAppsScript.Spreadsheet.Spreadsheet;
  try{
      workbook = SpreadsheetApp.openById(json.bookid || "");
  }
  catch(e){
    const mainbook = SpreadsheetApp.openById("1HtGLfZMu7TqxOQUIrGRfk8WS6saq-vMNhQg3tKrEQZk");
    workbook = mainbook.copy(json.filename || "");
    
  }
  return workbook;
}


function getSheet(data: PostEvent, book: Spreadsheet){
  var json: Json = JSON.parse(data.postData.contents);

  var sheet = book.getSheetByName(json.sheetname);

  if(!sheet){
      sheet = book.insertSheet();
      sheet.setName(json.sheetname);
  }
  return sheet;
}