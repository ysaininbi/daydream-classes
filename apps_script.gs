const LEADS_SHEET="Leads";
const VISITS_SHEET="Visits";

function getSheet_(name, headers){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name); sh.appendRow(headers);}
  return sh;
}

function doPost(e){
  const data=JSON.parse(e.postData.contents||"{}");
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const now=new Date(), tz=Session.getScriptTimeZone()||"Asia/Kolkata";
  const date=Utilities.formatDate(now,tz,"dd-MM-yyyy");
  const time=Utilities.formatDate(now,tz,"hh:mm:ss a");
  const type=data.type||"lead";

  if(type==="visit"){
    const sh=getSheet_(VISITS_SHEET,["Date","Time","Visitor ID","Page","Referrer","User Agent"]);
    sh.appendRow([date,time,data.visitorId||"",data.page||"",data.referrer||"",data.userAgent||""]);
  } else {
    const sh=getSheet_(LEADS_SHEET,["Date","Time","Student Name","Class","Contact Number","Page","User Agent"]);
    sh.appendRow([date,time,data.name||"",data.class_name||"",data.phone||"",data.page||"",data.userAgent||""]);
  }
  return ContentService.createTextOutput(JSON.stringify({status:"success"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  if(e.parameter.action==="dashboard") return dashboard_();
  return ContentService.createTextOutput("DAYDREAM API OK");
}
function dashboard_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), tz=Session.getScriptTimeZone()||"Asia/Kolkata";
  const today=Utilities.formatDate(new Date(),tz,"dd-MM-yyyy");
  const leadSh=ss.getSheetByName(LEADS_SHEET), visitSh=ss.getSheetByName(VISITS_SHEET);
  const leads=leadSh?leadSh.getDataRange().getValues().slice(1):[];
  const visits=visitSh?visitSh.getDataRange().getValues().slice(1):[];
  const tl=leads.filter(r=>String(r[0])===today);
  const tv=visits.filter(r=>String(r[0])===today);
  const cls=[...new Set(tl.map(r=>String(r[3])).filter(Boolean))];
  const out=leads.slice(-100).reverse().map(r=>({date:r[0],time:r[1],name:r[2],cls:r[3],phone:r[4]}));
  return ContentService.createTextOutput(JSON.stringify({todayVisits:tv.length,todayLeads:tl.length,todayClasses:cls.length,totalLeads:leads.length,leads:out})).setMimeType(ContentService.MimeType.JSON);
}
function setupSheets(){
  getSheet_(LEADS_SHEET,["Date","Time","Student Name","Class","Contact Number","Page","User Agent"]);
  getSheet_(VISITS_SHEET,["Date","Time","Visitor ID","Page","Referrer","User Agent"]);
}
