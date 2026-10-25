/** Dooley's Adventure School — Google Apps Script backend.
 * Deploy as a Web app: Execute as Me; access Anyone.
 * Set Script Property SHEET_ID to the existing spreadsheet ID.
 * No shared secret is required.
 */
const HEADERS = ['id','date','environment','yanks','transitionsGood','transitionsTotal','success','attempts','rating','notes','savedAt'];
function sheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID') || props.getProperty('DOOLEY_SHEET_ID');
  if (!id) throw Error('Set SHEET_ID in Script Properties to your spreadsheet ID.');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName('Walks');
  if (!sh) sh = ss.insertSheet('Walks');
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}
function clean_(v,limit) {return String(v==null?'':v).slice(0,limit).replace(/^[=+@\-\t\r]/,'\'$&');}
function number_(v,max) { if(v===''||v==null)return ''; const n=Number(v);if(!Number.isInteger(n)||n<0||n>max)throw Error('Invalid count');return n; }
function write_(data) {
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try {
    const sh=sheet_(), id=clean_(data.id,100);
    if(!/^[a-zA-Z0-9_-]{8,100}$/.test(id))throw Error('Invalid record ID');
    const last=sh.getLastRow();
    if(last>1 && sh.getRange(2,1,last-1,1).getValues().some(r=>r[0]===id))return {ok:true,id,duplicate:true};
    const date=String(data.date||'');if(!/^\d{4}-\d{2}-\d{2}$/.test(date))throw Error('Invalid date');
    const good=number_(data.transitionsGood,100), total=number_(data.transitionsTotal,100),success=number_(data.success,100),attempts=number_(data.attempts,100);
    if(good!==''&&total!==''&&good>total)throw Error('Controlled transitions exceed attempts');
    if(success!==''&&attempts!==''&&success>attempts)throw Error('Cue successes exceed attempts');
    sh.appendRow([id,date,clean_(data.environment,50),number_(data.yanks,100),good,total,success,attempts,number_(data.rating,5),clean_(data.notes,1000),new Date().toISOString()]);
    return {ok:true,id};
  } finally {lock.releaseLock()}
}
function records_() {
  const sh=sheet_(),last=sh.getLastRow();
  if(last<2)return {ok:true,records:[]};
  const rows=sh.getRange(2,1,Math.min(last-1,1000),HEADERS.length).getDisplayValues();
  return {ok:true,records:rows.map(r=>Object.fromEntries(HEADERS.map((h,i)=>[h,r[i]])))};
}
function doGet(e) {
  const p=e.parameter||{};
  let result;try {result=records_()}catch(err){result={ok:false,error:String(err.message||err)}}
  const callback=String(p.callback||'');
  if(!/^[a-zA-Z_$][\w$]*$/.test(callback))return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  return ContentService.createTextOutput(callback+'('+JSON.stringify(result).replace(/</g,'\\u003c')+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function doPost(e) {
  let result;try {const data=JSON.parse(e.postData.contents);result=write_(data)}catch(err){result={ok:false,error:String(err.message||err)}}
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
