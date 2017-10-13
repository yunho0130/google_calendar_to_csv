function export_gcal_to_gsheet(){
//Refreshing previous event for checking missing events 
//Default values is 2 
var dayPeriod = 5
var currentEvent = new Date()
var startEvent = new Date(currentEvent.getTime()-dayPeriod*(24*3600*1000))
var endEvent = new Date(currentEvent.getTime()+dayPeriod*(3*3600*1000)) // Plan
//var endEvent = new Date("September 9, 2017 20:00:00 GMT +0900")

// Update detail 
// Missing Location handling
// Fixed address -> Append 
// 
// 

// Reference Websites:
// https://developers.google.com/apps-script/reference/calendar/calendar
// https://developers.google.com/apps-script/reference/calendar/calendar-event
//
var allCal = CalendarApp.getAllCalendars();
var row = 2
j = 0
var sheet = SpreadsheetApp.getActiveSheet();
var dtRange = sheet.getDataRange();
// My category numbers 8
for (var k=0;k<8;k++) {
  var cal = allCal[k];

  var events = cal.getEvents(startEvent, endEvent);
  events = events.sort()
  //var events = cal.getEventsForDay(new Date());
// Loop through all calendar events found and write them out starting on calulated ROW 2 (i+2)
  row=row+events.length;

for (var i=0;i<events.length;i++) {
  row=i+row;
  var event_title = events[i].getTitle()
  // Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
  // NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
  var summary = ''
  var location = ''
  var hpi  = ''
  var ctn = ''
  var cal  = ''
  var tags = ''
  var title = ''
  var imp = ''
  var evt_id = events[i].getId();
  try{ 
    atIdx = event_title.indexOf("@")
    if (atIdx < 1){title = event_title} else{  title = event_title.substring(0, atIdx-1)}
  } catch(e) { }
  try{ summary = /(\[)([가-힣A-Za-z0-9 _-~]+)(\])/.exec(event_title)[0] } catch(e) { }
  try{ location = /(@)([가-힣A-Za-z0-9 _-]+)( )/.exec(event_title)[0] } catch(e) { }
  try{ hpi = /(h|hpi)=([0-9])/.exec(event_title)[2] } catch(e) { }
  try{ ctn = /(c|ctn|cti)=([0-9])/.exec(event_title)[2] } catch(e) { }
  try{ cal = /(k|kal|cal)=([0-9]+)/.exec(event_title)[2] } catch(e) { }
  try{ imp = /(p|imp)=([0-9]+)/.exec(event_title)[2] } catch(e) { }
  try{ 
    sharpIdx = event_title.indexOf("#")
    if (sharpIdx > 3){
      tags = '#'+event_title.substring(sharpIdx+1, event_title.length) 
    }
  } catch(e) { }
  //var stTime = events[i].getStartTime();
  var stTime = Utilities.formatDate(events[i].getStartTime(), "GMT+09", "yyyy/MM/dd HH:mm");
  //var endTime = events[i].getEndTime();
  var endTime = Utilities.formatDate(events[i].getEndTime(), "GMT+09", "yyyy/MM/dd HH:mm");
  var lrIdx = dtRange.getLastRow()-1
  var values = dtRange.getValues()
  var evt_id = events[i].getId();
  var skip = 'N'
  // Recent 100 event for checking # For reduce time complexity
  var numRct = 50*dayPeriod //150
  var initialNum = 0;
  if (lrIdx<numRct){initialNum = 0} else {initialNum = lrIdx-numRct}
  // Duplicate check 
  for (var n=initialNum; n<lrIdx+1; n++){
    if(summary == '' ) { 
      skip = 'Y' 
    } 
    else if (evt_id == values[n][11] && events[i].getLastUpdated().toString() == values[n][12]){
      skip = 'Y' 
    }
    else if (evt_id == values[n][11] && events[i].getLastUpdated().toString() != values[n][12] && skip != 'M') { 
      skip = 'M' 
      Logger.log('modified')
      sheet.deleteRow(n+1);
      Logger.log('# Append, skip: ' + skip)
      var details=[[stTime, endTime , (events[i].getEndTime()-events[i].getStartTime())/60000, summary, title, location, hpi , ctn, cal ,imp, tags, evt_id, events[i].getLastUpdated()]];
      sheet.appendRow(details[0])   
    }
    else { }
  };// n = end of inner for statment 
  if (skip == 'Y' || skip =='M' ) { }
  //  else if(skip == 'M'){
  //    for (j=0; j<length??; j++){
  //    evt_id
  //    }
  //    sheet.deleteRow(n);
  //    var details=[[stTime, endTime , (events[i].getEndTime()-events[i].getStartTime())/3600000, summary, title, location, hpi , ctn, cal ,imp, tags, evt_id, events[i].getLastUpdated()]];
  //    sheet.appendRow(details[0])    
  //  }
  else{
    Logger.log('Append, skip: ' + skip)
    var details=[[stTime, endTime , (events[i].getEndTime()-events[i].getStartTime())/60000, summary, title, location, hpi , ctn, cal ,imp, tags, evt_id, events[i].getLastUpdated().toString()]];
    sheet.appendRow(details[0])
  }
}// i
 }// k = end of main for statement
//  for (k=1;k<modifiedIdx.length; k++){
//    i= modifiedIdx[k-1]
//    Logger.log('modified event trace')
//    Logger.log(i)
//    sheet.deleteRow(i);
//    var details=[[stTime, endTime , (events[i].getEndTime()-events[i].getStartTime())/3600000, summary, title, location, hpi , ctn, cal ,imp, tags, evt_id, events[i].getLastUpdated()]];
//    sheet.appendRow(details[0])
//  }
// ## Chekcing Today event
// header fix
var stRange = SpreadsheetApp.getActiveSheet().getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
stRange.sort(1) // sort by start time
sheet.deleteColumn(14); // delete yesterday marker
sheet.getRange('N1').setValue('Today');
for (var m=0; m<lrIdx+1; m++){
  sheet.getRange(lrIdx-m+1, 14).setValue('true');
  if(/(\[\~결산\])/.test(values[lrIdx-m][4])){  break; }
  else {  
  }
 }
}

// Introduction
function onOpen() {
  Browser.msgBox('App Instructions - Please Read This Message', '1) Click Tools then Script Editor\\n2) Read/update the code with your desired values.\\n3) Then when ready click Run export_gcal_to_gsheet from the script editor.', Browser.Buttons.OK);

}

function regTest(){
  var name = "";
//  var regExp = new RegExp("(hpi)=([0-9])"); 
  //re[Symbol.search](str)
  Browser.msgBox('#'+name.substring(name.indexOf("#")+1, name.length));
  //Browser.msgBox(/(#)([가-힣A-Za-z0-9 _-]+)/.exec(name));
}

function cateTest(){
  var mycal = "yunho0130@gmail.com";
  var cal = CalendarApp.getAllCalendars();
  //var events = cal.getEvents(new Date("September 5, 2017 05:00:00 GMT +09:00"), new Date("September 6, 2017 05:00:00 GMT +09:00"));
  Browser.msgBox(cal[7].getName()) //0~7
}

function check_arr(){
var arr_test = ['a', 'b', 'c']
arr_test.push('d')
arr_test.push('d')
Logger.log(arr_test.pop())
}

