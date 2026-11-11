// 光子艦隊：在目標 Google 試算表的「擴充功能 → Apps Script」貼上全部程式。
const SHEET_NAME = '光電效應';
const HEADERS = ['日期','開始時間','結束時間','遊戲時長','分數','姓名','班別','學號','重新開始次數','答對題數','原始分數','時間係數','紀錄ID','完成狀態'];
function setup() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) throw new Error('請從目標試算表開啟 Apps Script');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID',ss.getId());
  let sh=ss.getSheetByName(SHEET_NAME)||ss.insertSheet(SHEET_NAME);
  if(sh.getLastRow()===0) sh.appendRow(HEADERS);
  else if(sh.getRange(1,1,1,HEADERS.length).getValues()[0].join('|')!==HEADERS.join('|')) throw new Error('同名分頁欄位不同，請改 SHEET_NAME 後重試');
  sh.setFrozenRows(1); sh.getRange(1,1,1,HEADERS.length).setBackground('#14213a').setFontColor('#ffe476').setFontWeight('bold');
}
function getSheet_(){
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if(!id)throw new Error('請先手動執行 setup 並授權');
  return SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME);
}
function find_(sh,id){return sh.getLastRow()<2?null:sh.getRange(2,13,sh.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();}
function text_(s){s=String(s||'').trim().slice(0,100);return /^[=+@-]/.test(s)?"'"+s:s;}
function doPost(e){
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const d=JSON.parse(e.postData.contents);
    if(!/^[a-zA-Z0-9-]{10,100}$/.test(d.id)||!d.name||!d.className||!d.studentNo)throw new Error('登記資料不完整');
    const start=new Date(d.start),end=new Date(d.end),sec=Number(d.duration),a=d.answers;
    if(!Number.isFinite(start.getTime())||!Number.isFinite(end.getTime())||end<start||!Number.isFinite(sec)||sec<0||!Array.isArray(a)||a.length!==10||a.some(x=>!Number.isInteger(x)||x<0||x>3))throw new Error('成績資料錯誤');
    const key=[1,0,2,1,0,2,3,0,2,1],correct=a.reduce((n,x,i)=>n+(x===key[i]?1:0),0);
    const factor=.6+.4*Math.max(0,Math.min(1,(300-sec)/200));
    const score=Math.round(correct*10*factor*10)/10;
    const sh=getSheet_();if(find_(sh,d.id))return ContentService.createTextOutput('duplicate');
    const tz='Asia/Macau',fmt=(d,f)=>Utilities.formatDate(d,tz,f);
    sh.appendRow([fmt(start,'yyyy-MM-dd'),fmt(start,'yyyy-MM-dd HH:mm:ss'),fmt(end,'yyyy-MM-dd HH:mm:ss'),Math.round(sec*10)/10,score,text_(d.name),text_(d.className),"'"+String(d.studentNo).slice(0,30),Math.max(0,Math.floor(Number(d.restarts)||0)),correct,correct*10,Math.round(factor*10000)/10000,d.id,'完成']);
    return ContentService.createTextOutput('ok');
  }finally{lock.releaseLock();}
}
function doGet(e){
  const cb=String(e.parameter.callback||'');
  if(!/^photon_cb_[a-zA-Z0-9_]+$/.test(cb))return ContentService.createTextOutput('光子艦隊接收器已啟動');
  let out={ok:false};
  try{const id=String(e.parameter.id||'');if(/^[a-zA-Z0-9-]{10,100}$/.test(id)){const sh=getSheet_(),cell=find_(sh,id);out={ok:!!cell};}}catch(err){out={ok:false,error:'請檢查 setup、分頁與部署設定'};}
  return ContentService.createTextOutput(cb+'('+JSON.stringify(out)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
