const SPREADSHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('OnTime HR')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() {
  if (SPREADSHEET_ID === 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE')
    throw new Error('Set SPREADSHEET_ID in Code.gs first.');
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function setup() {
  const ss=ss_();
  const defs={
    Employees:['EmployeeID','Name','Department','Role','PasswordHash','Status','CreatedAt'],
    Attendance:['RecordID','EmployeeID','Date','CheckIn','CheckOut','Status','Notes','CreatedAt','UpdatedAt'],
    Leaves:['LeaveID','EmployeeID','FromDate','ToDate','Type','Reason','Status','CreatedAt','UpdatedAt'],
    Settings:['Key','Value']
  };
  Object.keys(defs).forEach(n=>{
    let sh=ss.getSheetByName(n)||ss.insertSheet(n);
    if(sh.getLastRow()===0) sh.appendRow(defs[n]);
  });
  const e=ss.getSheetByName('Employees');
  if(!e.getDataRange().getValues().slice(1).some(r=>String(r[0])==='ADMIN001'))
    e.appendRow(['ADMIN001','Administrator','HR','Admin',hash_('admin123'),'Active',new Date()]);
  return 'OnTime HR setup complete';
}

function hash_(p){
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,p,Utilities.Charset.UTF_8)
    .map(b=>('0'+(b<0?b+256:b).toString(16)).slice(-2)).join('');
}
function safeUser_(r){return {id:String(r[0]),name:String(r[1]),department:String(r[2]||''),role:String(r[3]||'Employee'),status:String(r[5]||'Active')};}
function session_(token){
  const x=CacheService.getScriptCache().get('s_'+token);
  if(!x) throw new Error('Session expired. Please login again.');
  return JSON.parse(x);
}
function login(id,password){
  const rows=ss_().getSheetByName('Employees').getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0]).toLowerCase()===String(id).trim().toLowerCase()){
      if(rows[i][5]!=='Active') throw new Error('Account is inactive.');
      if(rows[i][4]!==hash_(password)) throw new Error('Invalid User ID or password.');
      const u=safeUser_(rows[i]), t=Utilities.getUuid();
      CacheService.getScriptCache().put('s_'+t,JSON.stringify(u),21600);
      return {token:t,user:u};
    }
  }
  throw new Error('Invalid User ID or password.');
}
function logout(t){if(t)CacheService.getScriptCache().remove('s_'+t);return true;}
function admin_(t){const u=session_(t);if(u.role!=='Admin')throw new Error('Admin permission required.');return u;}

function today_(d){return Utilities.formatDate(d||new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');}
function attObj_(r){
 return {recordId:String(r[0]),employeeId:String(r[1]),date:today_(new Date(r[2])),
 checkIn:r[3]?Utilities.formatDate(new Date(r[3]),Session.getScriptTimeZone(),'HH:mm'):'',
 checkOut:r[4]?Utilities.formatDate(new Date(r[4]),Session.getScriptTimeZone(),'HH:mm'):'',
 status:String(r[5]||''),notes:String(r[6]||'')};
}
function getTodayStatus(t){
 const u=session_(t),rows=ss_().getSheetByName('Attendance').getDataRange().getValues(),k=today_();
 for(let i=1;i<rows.length;i++)if(String(rows[i][1])===u.id&&today_(new Date(rows[i][2]))===k)return attObj_(rows[i]);
 return null;
}
function markAttendance(t,action){
 const u=session_(t), sh=ss_().getSheetByName('Attendance'), rows=sh.getDataRange().getValues(), k=today_();
 for(let i=1;i<rows.length;i++){
   if(String(rows[i][1])===u.id&&today_(new Date(rows[i][2]))===k){
     if(action==='checkin'){if(rows[i][3])throw new Error('Already checked in today.');sh.getRange(i+1,4).setValue(new Date());}
     else {if(!rows[i][3])throw new Error('Check in first.');if(rows[i][4])throw new Error('Already checked out today.');sh.getRange(i+1,5).setValue(new Date());}
     sh.getRange(i+1,6).setValue('Present');sh.getRange(i+1,9).setValue(new Date());return true;
   }
 }
 if(action==='checkout')throw new Error('No check-in found for today.');
 sh.appendRow([Utilities.getUuid(),u.id,new Date(),new Date(),'','Present','',new Date(),new Date()]);return true;
}
function getMyAttendance(t){
 const u=session_(t),rows=ss_().getSheetByName('Attendance').getDataRange().getValues().slice(1);
 return rows.filter(r=>String(r[1])===u.id).map(attObj_);
}
function getAllAttendance(t){
 admin_(t);return ss_().getSheetByName('Attendance').getDataRange().getValues().slice(1).map(attObj_);
}
function listEmployees(t){admin_(t);return ss_().getSheetByName('Employees').getDataRange().getValues().slice(1).map(safeUser_);}
function createEmployee(t,d){
 admin_(t);const sh=ss_().getSheetByName('Employees'),rows=sh.getDataRange().getValues();
 if(!d.id||!d.name||!d.password)throw new Error('Name, ID and password are required.');
 if(rows.slice(1).some(r=>String(r[0]).toLowerCase()===String(d.id).toLowerCase()))throw new Error('User ID already exists.');
 sh.appendRow([d.id,d.name,d.department||'','Employee',hash_(d.password),'Active',new Date()]);
 return true;
}
function deactivateEmployee(t,id){
 admin_(t);if(id==='ADMIN001')throw new Error('Main admin cannot be deactivated.');
 const sh=ss_().getSheetByName('Employees'),rows=sh.getDataRange().getValues();
 for(let i=1;i<rows.length;i++)if(String(rows[i][0])===String(id)){sh.getRange(i+1,6).setValue('Inactive');return true;}
 throw new Error('Employee not found.');
}
function applyLeave(t,d){
 const u=session_(t);if(!d.fromDate||!d.toDate||!d.type)throw new Error('Complete the leave form.');
 ss_().getSheetByName('Leaves').appendRow([Utilities.getUuid(),u.id,new Date(d.fromDate),new Date(d.toDate),d.type,d.reason||'','Pending',new Date(),new Date()]);
 return true;
}
function myLeaves(t){
 const u=session_(t);return ss_().getSheetByName('Leaves').getDataRange().getValues().slice(1)
  .filter(r=>String(r[1])===u.id).map(r=>({id:r[0],from:r[2],to:r[3],type:r[4],reason:r[5],status:r[6]}));
}
function leaveAdmin(t){
 admin_(t);return ss_().getSheetByName('Leaves').getDataRange().getValues().slice(1)
  .map(r=>({id:r[0],employeeId:r[1],from:r[2],to:r[3],type:r[4],reason:r[5],status:r[6]}));
}
function updateLeave(t,id,status){
 admin_(t);const sh=ss_().getSheetByName('Leaves'),rows=sh.getDataRange().getValues();
 for(let i=1;i<rows.length;i++)if(String(rows[i][0])===String(id)){sh.getRange(i+1,7).setValue(status);sh.getRange(i+1,9).setValue(new Date());return true;}
 throw new Error('Leave not found.');
}
