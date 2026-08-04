/*
Google Apps Script backend for Attendance app
Deploy: create a new Google Apps Script project, paste this file, set SPREADSHEET_ID, then Publish -> Deploy as web app (execute as: Me, who has access: Anyone)

Sheets expected in the spreadsheet (create these with headers):
- Teachers: id,name,subject,pin,targetArrival
- Records: id,teacherId,date,checkIn,checkOut,distance
- (optional) Config: key,value  OR geofence stored in PropertiesService

Replace SPREADSHEET_ID with your sheet id.
*/

const SPREADSHEET_ID = 'REPLACE_WITH_SPREADSHEET_ID';

function doGet(e){
  try{
    const action = (e.parameter.action || '').toString();
    switch(action){
      case 'getTeachers': return jsonOk({teachers: getTeachers()});
      case 'addTeacher': return jsonOk(addTeacher(e.parameter));
      case 'getRecords': return jsonOk({records: getRecords(e.parameter.month)});
      case 'scan': return jsonOk(scan(e.parameter));
      case 'editRecord': return jsonOk(editRecord(e.parameter));
      case 'deleteRecord': return jsonOk(deleteRecord(e.parameter.recordId));
      case 'getGeofence': return jsonOk({geofence: getGeofence()});
      case 'setGeofence': return jsonOk(setGeofence(e.parameter));
      case 'getTeacherByPin': return jsonOk(getTeacherByPin(e.parameter.pin));
      default: return jsonError('Unknown action: '+action);
    }
  }catch(err){
    return jsonError(err.message || String(err));
  }
}

function jsonOk(obj){
  const out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
function jsonError(msg){
  const out = ContentService.createTextOutput(JSON.stringify({error: msg}));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

function ss(){
  if (!SPREADSHEET_ID || SPREADSHEET_ID==='REPLACE_WITH_SPREADSHEET_ID') throw new Error('Set SPREADSHEET_ID in the script');
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function sheetRows(sheetName){
  const s = ss().getSheetByName(sheetName);
  if(!s) return [];
  const vals = s.getDataRange().getValues();
  if(vals.length<=1) return [];
  const headers = vals[0];
  return vals.slice(1).map(r=>{
    const obj = {};
    headers.forEach((h,i)=> obj[String(h).trim()] = r[i]);
    return obj;
  });
}

function getTeachers(){
  const rows = sheetRows('Teachers');
  return rows.map(r=>({id:String(r.id), name:String(r.name||''), subject: String(r.subject||''), pin: String(r.pin||''), targetArrival: String(r.targetArrival||'')}));
}

function addTeacher(params){
  const name = (params.name||'').toString();
  const subject = (params.subject||'').toString();
  const pin = (params.pin||'').toString();
  const target = (params.targetArrival||'').toString();
  if(!name) throw new Error('Name required');
  const id = Utilities.getUuid();
  const s = ss().getSheetByName('Teachers') || ss().insertSheet('Teachers');
  if(s.getLastRow()===0){ s.appendRow(['id','name','subject','pin','targetArrival']); }
  s.appendRow([id, name, subject, pin, target]);
  return {id};
}

function getRecords(monthPrefix){
  // monthPrefix like '2026-08' or undefined -> return all
  const rows = sheetRows('Records');
  const out = rows.map(r=>({id:String(r.id), teacherId:String(r.teacherId), date:String(r.date||''), checkIn:String(r.checkIn||''), checkOut:String(r.checkOut||''), distance: r.distance||''}));
  if(monthPrefix) return out.filter(x=>x.date && x.date.indexOf(monthPrefix)===0);
  return out;
}

function scan(params){
  // params: pin, date (YYYY-MM-DD), time (HH:MM), distance (meters optional)
  const pin = (params.pin||'').toString();
  const date = (params.date||'').toString();
  const time = (params.time||'').toString();
  const distance = params.distance ? Number(params.distance) : '';
  if(!pin) throw new Error('PIN required');
  const teachers = getTeachers();
  const teacher = teachers.find(t=>t.pin===pin);
  if(!teacher) throw new Error('Teacher with that PIN not found');

  // find an open record for this teacher (same date, checkOut empty)
  const s = ss().getSheetByName('Records') || ss().insertSheet('Records');
  if(s.getLastRow()===0) s.appendRow(['id','teacherId','date','checkIn','checkOut','distance']);
  const data = s.getDataRange().getValues();
  const headers = data[0];
  let openRow = null;
  for(let i=1;i<data.length;i++){
    const row = data[i];
    const rowObj = {};
    headers.forEach((h,idx)=> rowObj[String(h).trim()] = row[idx]);
    if(String(rowObj.teacherId)===String(teacher.id) && String(rowObj.date)===date && !rowObj.checkOut){ openRow = {index:i+1, row: rowObj}; break; }
  }
  if(openRow){
    // check-out
    const rowIndex = openRow.index;
    const ci = openRow.row.checkIn || '';
    s.getRange(rowIndex, headers.indexOf('checkOut')+1).setValue(time);
    if(distance!=='') s.getRange(rowIndex, headers.indexOf('distance')+1).setValue(distance);
    return {action:'checkOut', recordId: String(openRow.row.id), teacher: teacher.name, date: date, time: time};
  } else {
    // create check-in
    const id = Utilities.getUuid();
    s.appendRow([id, teacher.id, date, time, '', distance]);
    return {action:'checkIn', recordId: id, teacher: teacher.name, date: date, time: time};
  }
}

function editRecord(params){
  const recordId = (params.recordId||'').toString();
  const checkIn = (params.checkIn||'').toString();
  const checkOut = (params.checkOut||'').toString();
  if(!recordId) throw new Error('recordId required');
  const s = ss().getSheetByName('Records');
  if(!s) throw new Error('No Records sheet');
  const data = s.getDataRange().getValues();
  const headers = data[0];
  for(let i=1;i<data.length;i++){
    if(String(data[i][headers.indexOf('id')])===recordId){
      if(checkIn!=='') s.getRange(i+1, headers.indexOf('checkIn')+1).setValue(checkIn);
      if(checkOut!=='') s.getRange(i+1, headers.indexOf('checkOut')+1).setValue(checkOut);
      return {recordId};
    }
  }
  throw new Error('Record not found');
}

function deleteRecord(recordId){
  if(!recordId) throw new Error('recordId required');
  const s = ss().getSheetByName('Records');
  if(!s) throw new Error('No Records sheet');
  const data = s.getDataRange().getValues();
  const headers = data[0];
  for(let i=1;i<data.length;i++){
    if(String(data[i][headers.indexOf('id')])===recordId){
      s.deleteRow(i+1);
      return {recordId};
    }
  }
  throw new Error('Record not found');
}

function getGeofence(){
  // prefer PropertiesService
  const p = PropertiesService.getScriptProperties().getProperty('geofence');
  if(p) return JSON.parse(p);
  // fallback to Config sheet
  const rows = sheetRows('Config');
  const g = rows.find(r=>r.key==='geofence');
  return g ? JSON.parse(g.value) : null;
}

function setGeofence(params){
  const lat = Number(params.lat);
  const lng = Number(params.lng);
  const radius = Number(params.radius || params.radiusMeters || params.radius_meters || 0);
  if(isNaN(lat) || isNaN(lng) || isNaN(radius)) throw new Error('Invalid geofence params');
  const obj = {lat:lat, lng:lng, radiusMeters: radius, enforce: true};
  PropertiesService.getScriptProperties().setProperty('geofence', JSON.stringify(obj));
  return obj;
}

function getTeacherByPin(pin){
  if(!pin) throw new Error('pin required');
  const t = getTeachers().find(x=>String(x.pin)===String(pin));
  if(!t) throw new Error('Not found');
  return t;
}
