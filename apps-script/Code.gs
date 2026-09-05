/**
 * ベビーシッター/ファミサポ料金シミュレーター用の依頼状況スプレッドシート連携API。
 * スプレッドシートの「拡張機能 > Apps Script」に貼り付けてデプロイしてください。
 */

const SHEET_NAME = "依頼状況";
const HEADERS = ["id","year","month","day","dow","label","assignment","assignmentName","hours","rate","confirmed"];

function getSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if(sheet.getLastRow() === 0){
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function readAllEntries_(){
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row[headers.indexOf("id")] !== "")
    .map(row=>{
      const obj = {};
      headers.forEach((h,i)=>{ obj[h] = row[i]; });
      obj.confirmed = obj.confirmed === true || obj.confirmed === "TRUE" || obj.confirmed === "true";
      return obj;
    });
}

function findRowIndexById_(sheet, id){
  const data = sheet.getDataRange().getValues();
  const idCol = HEADERS.indexOf("id");
  for(let i=1;i<data.length;i++){
    if(String(data[i][idCol]) === String(id)) return i+1;
  }
  return -1;
}

function doGet(e){
  try{
    return jsonOut_({ ok:true, entries: readAllEntries_() });
  }catch(err){
    return jsonOut_({ ok:false, error: String(err) });
  }
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const sheet = getSheet_();

    if(action === "upsert"){
      const entry = body.entry;
      const rowValues = HEADERS.map(h => entry[h] !== undefined ? entry[h] : "");
      const rowIndex = findRowIndexById_(sheet, entry.id);
      if(rowIndex === -1){
        sheet.appendRow(rowValues);
      } else {
        sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([rowValues]);
      }
      return jsonOut_({ ok:true });
    }

    if(action === "delete"){
      const rowIndex = findRowIndexById_(sheet, body.id);
      if(rowIndex !== -1) sheet.deleteRow(rowIndex);
      return jsonOut_({ ok:true });
    }

    if(action === "replaceAll"){
      const entries = body.entries || [];
      sheet.clearContents();
      sheet.appendRow(HEADERS);
      entries.forEach(entry=>{
        sheet.appendRow(HEADERS.map(h => entry[h] !== undefined ? entry[h] : ""));
      });
      return jsonOut_({ ok:true });
    }

    return jsonOut_({ ok:false, error: "unknown action: " + action });
  }catch(err){
    return jsonOut_({ ok:false, error: String(err) });
  }
}
