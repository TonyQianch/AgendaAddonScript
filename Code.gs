//function onOpen(){
//  SpreadsheetApp.getUi().createAddonMenu().addItem('Reminder Tool', 'showSidebar').addToUi();
//}
//
//function showSidebar() {
//  var html = HtmlService.createTemplateFromFile("reminderTool").evaluate().setTitle("Reminder Tool");
//  SpreadsheetApp.getUi().showSidebar(html);
//}

function promptDate() {
  var agenda = SpreadsheetApp.getActive().getSheetByName("Agenda");
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt("Date", "Please enter the date of the agenda (MM/DD/YYYY)", ui.ButtonSet.OK);
  var response = prompt.getResponseText();
  if (response.slice(2,3)=="/" && response.slice(5,6)=="/"&&response.length==10) {
    agenda.getRange(1,1).setValue(response);
  }
  else promptBox();
}

function promptCheck(date) {
  var ui = SpreadsheetApp.getUi();
  var message = "Please confirm the date of the agenda: "+date;
  var button = ui.alert("Confirm Date", message, ui.ButtonSet.YES_NO);
  if (button == ui.Button.NO) {promptDate();}
}

function getNRow(sheet, startRow, startCol){
  var n = 0;
  var range;
  while (true){
    range = sheet.getRange(startRow+n,startCol);
    if (range.isBlank()){
      return n;
    }
    else {
      n++;
    }
  }
}

function getEmailCol(sheet,startRow,endRow) {
  var emailList = SpreadsheetApp.getActive().getSheetByName("EmailList");
  var numRow = emailList.getLastRow()-1;
  var data = emailList.getRange(2,1,numRow,4).getValues();
  var i, studentName;
  for (var row=startRow; row<endRow+1; row++) {
    studentName = sheet.getRange(row,4).getValue();
    for ( i=0; i<numRow; i++) {
      if (data[i][0]==studentName) {
        sheet.getRange(row,1).setValue(data[i][1]);
        sheet.getRange(row,2).setValue(data[i][2]);
        sheet.getRange(row,3).setValue(data[i][3]);
        break;
      }
    }
  }
}

function clearPTList(){
  var PT = SpreadsheetApp.getActive().getSheetByName("PT");
  var numRow = PT.getLastRow()-1;
  PT.getRange(2, 1, numRow, PT.getLastColumn()).clear();
}

function completePTList(){
  var PT = SpreadsheetApp.getActive().getSheetByName("PT");
  var endRow = PT.getLastRow();
  PT.getRange(2,1,endRow-1,8).setNumberFormat("@");
  getEmailCol(PT,2,endRow);
  for (var i=0; i<endRow-1; i++) {
    PT.getRange(2+i,7).setValue("Private Tutoring");
  }
}

function clearCCList(){
  var CC = SpreadsheetApp.getActive().getSheetByName("CC");
  var numRow = CC.getLastRow()-1;
  CC.getRange(2, 1, numRow, CC.getLastColumn()).clear();
}

function completeCCList(){
  var CC = SpreadsheetApp.getActive().getSheetByName("CC");
  var endRow = CC.getLastRow();
  CC.getRange(2,1,endRow-1,8).setNumberFormat("@");
  getEmailCol(CC,2,endRow);
  for (var i=0; i<endRow-1; i++) {
    CC.getRange(2+i,7).setValue("Counseling Meeting");
  }
}

function dateToDay(date) {
  var weekday = new Array(7);
  weekday[0] = "Sunday";
  weekday[1] = "Monday";
  weekday[2] = "Tuesday";
  weekday[3] = "Wednesday";
  weekday[4] = "Thursday";
  weekday[5] = "Friday";
  weekday[6] = "Saturday";
  return weekday[date.getDay()];
}

function sendEmail(address,subject,body) {
  var img = DriveApp.getFileById("1czIiXmxD5sHIy7PKTcEH5KbSsSrfeeu3").getBlob();
  body += "<img src=\"cid:sampleImage\" width=\"50%\"/>";
  MailApp.sendEmail({to: address, subject: subject, htmlBody: body, inlineImages: {sampleImage: img}});
}

function sendPTEmail(address,student,time,withWho,meetingType,date){
  var day = dateToDay(date);
  var subject = "[YSI] Private Tutoring Reminder -- "+student+" --"+date.toDateString()+"--"+time;
  var body = "Dear "+student+", <br><br>This is a reminder that you have your "+meetingType+" session with <strong>"+withWho
           +" at "+time+" on "+date.toDateString()+"</strong>. <br><br>Please also note that only parents may reschedule or cancel a student's session "
           +" by calling or emailing AT LEAST ONE YSI BUSINESS DAY IN ADVANCE of the scheduled session. Parents who do not call "
           +"or email at least one business day in advance will still be charged the full rate as we are still required to pay "
           +"our instructors.<br><br>We look forward to seeing you on <strong>"+day+"</strong>!<br><br>Best regards,<br><br>";
  sendEmail(address, subject, body);
}

function sendCCEmail(address,student,time,withWho,meetingType,date){
  var day = dateToDay(date);
  var subject = "[YSI] Counseling Meeting Reminder -- "+student+" --"+date.toDateString()+"--"+time;
  var body = "Dear "+student+", <br><br>This is a reminder that you have your "+meetingType+" session with <strong>"+withWho
           +" at "+time+" on "+date.toDateString()+"</strong>. <br><br>Please also note that only parents may reschedule or cancel a student's session "
           +" by calling or emailing AT LEAST ONE YSI BUSINESS DAY IN ADVANCE of the scheduled session. Parents who do not call "
           +"or email at least one business day in advance will still be charged the full rate as we are still required to pay "
           +"our instructors.<br><br>We look forward to seeing you on <strong>"+day+"</strong>!<br><br>Best regards,<br><br>";
  sendEmail(address, subject, body);
}

function studentReminderAddress(sheet,row) {
  var address;
  if (!sheet.getRange(row,1).isBlank()) {
    if (!sheet.getRange(row,2).isBlank()) {
      if (!sheet.getRange(row,3).isBlank()) {
        address = sheet.getRange(row,1).getValue()+","+sheet.getRange(row,2).getValue()+","+sheet.getRange(row,3).getValue();
      }
      else address = sheet.getRange(row,1).getValue()+","+sheet.getRange(row,2).getValue();
    }
    else address = sheet.getRange(row,1).getValue();
  }
  else {
    if (!sheet.getRange(row,2).isBlank()) {
      if (!sheet.getRange(row,3).isBlank()) {
        address = sheet.getRange(row,2).getValue()+","+sheet.getRange(row,3).getValue();
      }
      else address = sheet.getRange(row,2).getValue();
    }
    else address = '';
  }
  return address;
}

function sendPTReminder(){
  var PT = SpreadsheetApp.getActive().getSheetByName("PT");
  var address, student, time, withWho, meetingType;
  var date = SpreadsheetApp.getActive().getSheetByName("Agenda").getRange('A1').getValue();
  promptCheck(date.toDateString());
  const startRow = 2;
  var numRow = getNRow(PT, startRow, 4);
  //send PT email
  for (var i=0; i<numRow; i++) {
    address = studentReminderAddress(PT,startRow+i);
    student = PT.getRange(startRow+i, 4).getValue();
    time = PT.getRange(startRow+i, 5).getValue();
    withWho = PT.getRange(startRow+i, 6).getValue();
    meetingType = PT.getRange(startRow+i, 7).getValue();
    sendPTEmail(address,student,time,withWho,meetingType,date);
    PT.getRange(startRow+i, 8).setValue('Done');
    SpreadsheetApp.flush();
  }
}

function sendCCReminder(){
  var CC = SpreadsheetApp.getActive().getSheetByName("CC");
  var address, student, time, withWho, meetingType;
  var date = SpreadsheetApp.getActive().getSheetByName("Agenda").getRange('A1').getValue();
  promptCheck(date.toDateString());
  const startRow = 2;
  var numRow = getNRow(CC, startRow, 4);
  //send PT email
  for (var i=0; i<numRow; i++) {
    address = studentReminderAddress(CC,startRow+i)+','+CC.getRange(startRow+i,8).getValue();
    student = CC.getRange(startRow+i, 4).getValue();
    time = CC.getRange(startRow+i, 5).getValue();
    withWho = CC.getRange(startRow+i, 6).getValue();
    meetingType = CC.getRange(startRow+i, 7).getValue();
    sendCCEmail(address,student,time,withWho,meetingType,date);
    CC.getRange(startRow+i, 9).setValue('Done');
    SpreadsheetApp.flush();
  }
}

function ptTable() {
  var PT = SpreadsheetApp.getActive().getSheetByName("PT");
  var numRow = PT.getLastRow();
  var data = PT.getRange(1,1,numRow,8).getValues();
  var row, col;
  var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
  var pttable = '<table ' + TABLEFORMAT +' ">';
  for (row = 0; row<data.length; row++){
    pttable += '<tr>';
    for (col = 0 ;col<data[row].length; col++){
      if (data[row][col] === "" || 0) {pttable += '<td>' + 'None' + '</td>';} 
      else if (row === 0)  {
        pttable += '<th>' + data[row][col] + '</th>';
      }
      else {pttable += '<td>' + data[row][col] + '</td>';}
    }
    pttable += '</tr>';
  }
  pttable += '</table>';
  return pttable;
}

function ccTable() {
  var CC = SpreadsheetApp.getActive().getSheetByName("CC");
  var numRow = CC.getLastRow();
  var data = CC.getRange(1,1,numRow,9).getValues();
  var row, col;
  var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
  var cctable = '<table ' + TABLEFORMAT +' ">';
  for (row = 0; row<data.length; row++){
    cctable += '<tr>';
    for (col = 0 ;col<data[row].length; col++){
      if (data[row][col] === "" || 0) {cctable += '<td>' + 'None' + '</td>';} 
      else if (row === 0)  {
        cctable += '<th>' + data[row][col] + '</th>';
      }
      else {cctable += '<td>' + data[row][col] + '</td>';}
    }
    cctable += '</tr>';
  }
  cctable += '</table>';
  return cctable;
}

function reportDone(){
  var date = SpreadsheetApp.getActive().getSheetByName("Agenda").getRange('A1').getValue();
//  var management = "amywang.ysi@gmail.com,crystallam.ysi@gmail.com,elimzheng.ysi@gmail.com,karenlee1.ysi@gmail.com,"
//                 +"lanyao.ysi@gmail.com,rileyliu.ysi@gmail.com,steveparkmail@gmail.com,yanjunlysi@gmail.com,ysprep@gmail.com";
  var testemail = "tonyqianchenhao@gmail.com";
  var subject = date+" PT and Counseling Reminder Status";
  var message = "The reminders emails for "+date+" has been sent. Please check the table below. Thank you!";
  var pttable = ptTable();
  var cctable = ccTable();
  var body = message+'<BR><BR>'+pttable+'<BR><BR>'+cctable+'<BR><BR>';
  sendEmail(testemail,subject,body);
}


function timeToRow(time) {
  var date = SpreadsheetApp.getActive().getSheetByName("Agenda").getRange(1,1).getValue();
  var day = date.getDay();
  var session = Number(time.slice(0,time.indexOf(":")));
  if (time.slice(time.indexOf(":")+1)=="30") session+=0.5;
  if (day < 6) session = session*2+10;
  else {
    if (session < 8) session = session*2+10;
    else session = session*2-14;
  }
  return session;
}

function timeToDuration(startTime,endTime) {
  var date = SpreadsheetApp.getActive().getSheetByName("Agenda").getRange(1,1).getValue();
  var duration = timeToRow(endTime,date)-timeToRow(startTime,date);
  return duration;
}

function parseTime(timeString) {
  var type = timeString.indexOf("-");
  var row, duration;
  if (type==-1) {
    row = timeToRow(timeString.slice(0,timeString.length-3));
    duration = 2;
  }
  else {
    var start = timeString.slice(0,type);
    var end = timeString.slice(type+1);
    row = timeToRow(start);
    duration = timeToDuration(start, end);
  }
  return [row,duration];
}

function checkRoomAvailable(col,row,duration) {
  var room = SpreadsheetApp.getActive().getSheetByName('Room');
  for (var i=0; i<duration; i++) {
    if (!room.getRange(row+i,col).isBlank()) return false;
  }
  return true;
}

function takeRoom(col, row, duration, student) {
  var room = SpreadsheetApp.getActive().getSheetByName('Room');
  for (var i=0; i<duration; i++ ) {
    room.getRange(row+i,col).setValue(student);
  }
}

function findRoomCol(roomName) {
  var room = SpreadsheetApp.getActive().getSheetByName('Room');
  var roomArray = room.getRange(1,1,1,room.getLastColumn()).getValues();
  return roomArray[0].indexOf(roomName);
}

function arrangeClass(date) {
  var day = dateToDay(date);
  var class = SpreadsheetApp.getActive().getSheetByName('Class');
  var firstCol = class.getRange(1,1,class.getLastRow(),1).getValues().map(function(e){return e[0];});
  var row = firstCol.indexOf(day)+1;
  var numClass = class.getRange(row,2).getValue();
  var className, time, roomName, teacher, schedule, roomCol;
  for (var i=0; i<numClass; i++) {
    className = class.getRange(row+1,2+i).getValue();
    time = class.getRange(row+2,2+i).getValue();
    roomName = class.getRange(row+3,2+i).getValue();
    schedule = parseTime(time);
    roomCol = findRoomCol(roomName);
    takeRoom(roomCol,schedule[0],schedule[1],className);
  }
}

function roomArrange(sheet,counseling) {
  var room = SpreadsheetApp.getActive().getSheetByName('Room');
  var count = sheet.getLastRow()-1;
  if (count > 0) {
    var header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues();
    var studentCol = header[0].indexOf("Student")+1;
    var timeCol = header[0].indexOf("Time")+1;
    var withCol = header[0].indexOf("With")+1;
    var roomCol = header[0].indexOf("Classroom")+1;
    var i, j, time, scheduleTime, student, roomFound, available;
    for (i=0; i<count; i++) {
      if (sheet.getRange(2+i,roomCol).isBlank()){
        roomFound = false;
        student = sheet.getRange(2+i,studentCol).getValue();
        time = sheet.getRange(2+i,timeCol).getValue();
        scheduleTime = parseTime(time);
        for (j=0; j<11;j++) {
          if (counseling) {
            available = checkRoomAvailable(12+j,scheduleTime[0],scheduleTime[1]);
            if (available) {
              takeRoom(12+j,scheduleTime[0],scheduleTime[1],student);
              sheet.getRange(2+i, roomCol).setValue(room.getRange(1,12+j).getValue());
              roomFound = true;
              break;
            }
          }
          else {
            available = checkRoomAvailable(j+2,scheduleTime[0],scheduleTime[1]);
            if (available) {
              takeRoom(j+2,scheduleTime[0],scheduleTime[1],student);
              sheet.getRange(2+i, roomCol).setValue(room.getRange(1,j+2).getValue());
              roomFound = true;
              break;
            }
          }
        }
        if (!roomFound) sheet.getRange(2+i, roomCol).setValue('No Room').setBackground('red');
      }
    }
  }
}

function getRyan() {
  var cc = SpreadsheetApp.getActive().getSheetByName('CC');
  var numRow = cc.getLastRow()-1;
  var student, time, scheduleTime;
  for (var i=0; i<numRow; i++) {
    if (cc.getRange(2+i,6).getValue()=="Ryan") {
      student = cc.getRange(2+i,4).getValue();
      time = cc.getRange(2+i,5).getValue();
      scheduleTime = parseTime(time);
      takeRoom(13,scheduleTime[0],scheduleTime[1],student);
      cc.getRange(2+i,10).setValue("Ryan");
    }
  }
}

function finishAgenda() {
  var pt = SpreadsheetApp.getActive().getSheetByName('PT');
  var cc = SpreadsheetApp.getActive().getSheetByName('CC');
  var room = SpreadsheetApp.getActive().getSheetByName('Room');
  var agenda = SpreadsheetApp.getActive().getSheetByName('Agenda');
  room.getRange(2,2,27,12).clear();
  agenda.getRange(4,1,agenda.getLastRow()-4,8).clear();
  agenda.getRange(4,1,agenda.getLastRow()-4,8).setNumberFormat("@");
  var date = agenda.getRange(1,1).getValue();
  promptCheck(date.toDateString());
  //Arrange Classes
  arrangeClass(date);
  //Arrange CC rooms
  cc.getRange(2,1,cc.getLastRow()-1,cc.getLastColumn()).sort(6);
  getRyan();
  roomArrange(cc,true);
  //Arrange PT rooms
  roomArrange(pt,false);
}








