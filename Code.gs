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
//  for (var i=0; i<endRow-1; i++) {
//    PT.getRange(2+i,7).setValue("Private Tutoring");
//  }
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
//  for (var i=0; i<endRow-1; i++) {
//    CC.getRange(2+i,7).setValue("Counseling Meeting");
//  }
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

function sendPTEmail(address,student,time,withWho,branch,date){
  var day = dateToDay(date);
  var subject = "[YSI] Private Tutoring Reminder -- "+student+" --"+date.toDateString()+"--"+time;
  var body = "Dear "+student+", <br><br>This is a reminder that you have your private tutoring session with <strong>"+withWho+" "+branch
           +" at "+time+" on "+date.toDateString()+"</strong>. <br><br>Please also note that only parents may reschedule or cancel a student's session "
           +" by calling or emailing AT LEAST ONE YSI BUSINESS DAY IN ADVANCE of the scheduled session. Parents who do not call "
           +"or email at least one business day in advance will still be charged the full rate as we are still required to pay "
           +"our instructors.<br><br>We look forward to seeing you on <strong>"+day+"</strong>!<br><br>Best regards,<br><br>";
  sendEmail(address, subject, body);
}

function sendCCEmail(address,student,time,withWho,branch,date){
  var day = dateToDay(date);
  var subject = "[YSI] Counseling Meeting Reminder -- "+student+" --"+date.toDateString()+"--"+time;
  var body = "Dear "+student+", <br><br>This is a reminder that you have your counselling meeting session with <strong>"+withWho+" "+branch
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

function getBranch(sheet,row) {
  var branch;
  if (sheet.getRange(row,7).isBlank()) {
    branch = "at San Marino branch";
  }
  else if (sheet.getRange(row,7).getValue()=="Online") {
    branch = "online";
  }
  else {
    branch = "at Arcadia branch";
  }
  return branch;
}

function sendPTReminder(){
  var PT = SpreadsheetApp.getActive().getSheetByName("PT");
  var address, student, time, withWho, branch;
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
    branch = getBranch(PT,startRow+i);
    sendPTEmail(address,student,time,withWho,branch,date);
    PT.getRange(startRow+i, 8).setValue('Done');
    SpreadsheetApp.flush();
  }
}

function sendCCReminder(){
  var CC = SpreadsheetApp.getActive().getSheetByName("CC");
  var address, student, time, withWho, branch;
  var date = SpreadsheetApp.getActive().getSheetByName("Agenda").getRange('A1').getValue();
  promptCheck(date.toDateString());
  const startRow = 2;
  var numRow = getNRow(CC, startRow, 4);
  //send PT email
  for (var i=0; i<numRow; i++) {
    if (CC.getRange(startRow+i, 9).isBlank) {
    address = studentReminderAddress(CC,startRow+i)+','+CC.getRange(startRow+i,8).getValue();
    student = CC.getRange(startRow+i, 4).getValue();
    time = CC.getRange(startRow+i, 5).getValue();
    withWho = CC.getRange(startRow+i, 6).getValue();
    branch = getBranch(CC,startRow+i);
    sendCCEmail(address,student,time,withWho,branch,date);
    CC.getRange(startRow+i, 9).setValue('Done');
    SpreadsheetApp.flush();
    }
  }
}

function ptTable() {
  var PT = SpreadsheetApp.getActive().getSheetByName("PT");
  var numRow = PT.getLastRow();
  var data = PT.getRange(1,4,numRow,5).getValues();
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
  var data = CC.getRange(1,4,numRow,6).getValues();
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
  var management = "amywang.ysi@gmail.com,crystallam.ysi@gmail.com,elimzheng.ysi@gmail.com,karenlee1.ysi@gmail.com,"
                 +"lanyao.ysi@gmail.com,rileyliu.ysi@gmail.com,steveparkmail@gmail.com,yanjunlysi@gmail.com,ysprep@gmail.com";
//  var testemail = "tonyqianchenhao@gmail.com";
  var subject = date+" PT and Counseling Reminder Status";
  var message = "The reminders emails for "+date+" has been sent. Please check the table below. Thank you!";
  var pttable = ptTable();
  var cctable = ccTable();
  var body = message+'<BR><BR>'+pttable+'<BR><BR>'+cctable+'<BR><BR>';
  sendEmail(management,subject,body);
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

function sortListbyTeacherClassCount() {
  var pt = SpreadsheetApp.getActive().getSheetByName('PT');
  var numRow = pt.getLastRow()-1;
  var lr = pt.getRange(2,1,numRow,9);
  lr.sort(6);
  var teacherArray = pt.getRange(2,6,numRow,1).getValues();
  Logger.log(teacherArray);
  var i, j, count;
  for (i=0; i<numRow-1; i++) {
    count = 1;
    Logger.log(teacherArray[i][0]);
    for (j=i+1; j<numRow; j++) {
      if (teacherArray[i][0] == teacherArray[j][0]) count++;
      else break;
    }
    Logger.log(count);
    for (j=i; j<i+count; j++) {
      teacherArray[j][0] = count+teacherArray[j][0];
    }
    i += count-1;
  }
  if (j==numRow-1) {
    teacherArray[numRow-1][0] = "1"+teacherArray[numRow-1][0];
  }
  pt.getRange(2,6,numRow,1).setValues(teacherArray);
  lr.sort({column: 6, ascending: false});
  Logger.log(pt.getRange(2,6,numRow,1).getValues());
  for (i=0;i<numRow;i++) {
    pt.getRange(2+i,6).setValue(pt.getRange(2+i,6).getValue().slice(1));
  }
  Logger.log(pt.getRange(2,6,numRow,1).getValues());
}

function settleArcadia (sheet,count,roomCol) {
  for (var i=0; i<count; i++) {
    if (sheet.getRange(2+i,7).getValue() == "Arcadia") {
      sheet.getRange(2+i,roomCol).setValue("ARC");
    }
  }
}

function roomArrange(sheet,counseling) {
  //sheet = SpreadsheetApp.getActive().getSheetByName('PT');
  var room = SpreadsheetApp.getActive().getSheetByName('Room');
  var count = sheet.getLastRow()-1;
  if (count > 0) {
    var header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues();
    var studentCol = header[0].indexOf("Student")+1;
    var timeCol = header[0].indexOf("Time")+1;
    var withCol = header[0].indexOf("With")+1;
    var roomCol = header[0].indexOf("Classroom")+1;
    settleArcadia(sheet,count,roomCol);
    var i, j, time, scheduleTime, student, roomFound, available;
    for (i=0; i<count; i++) {
      if (sheet.getRange(2+i,roomCol).isBlank()){
        roomFound = false;
        student = sheet.getRange(2+i,withCol).getValue();
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
      if (cc.getRange(2+i,7).getValue()=="Arcadia") {
        cc.getRange(2+i,10).setValue("ARC");
      }
      else {
        student = cc.getRange(2+i,4).getValue();
        time = cc.getRange(2+i,5).getValue();
        scheduleTime = parseTime(time);
        takeRoom(13,scheduleTime[0],scheduleTime[1],student);
        cc.getRange(2+i,10).setValue("Ryan");
      }
    }
  }
}

function fetchClasses() {
  var class = SpreadsheetApp.getActive().getSheetByName('Class');
  var agenda = SpreadsheetApp.getActive().getSheetByName('Agenda');
  var day = dateToDay(agenda.getRange(1,1).getValue());
  //find correct section by date -> day
  var firstCol = class.getRange(1,1,class.getLastRow(),1).getValues().map(function(e){return e[0];});
  var row = firstCol.indexOf(day)+2;
  var numClass = class.getRange(row-1,2).getValue();
  var currentRow = 7;//row to be written in agenda
  var value
  if (day=="Saturday") currentRow = 4;
  else {
    value = [["Study Hall"],["3:00-7:00"]];
    agenda.getRange(4,1,2,1).setValues(value);
    agenda.getRange(4,1).setFontWeight('bold');
  }
  for (var i=0; i<numClass; i++) {
    value = class.getRange(row,2+i).getValue();
    agenda.getRange(currentRow,1).setValue(value)
                                 .setFontWeight('bold')
                                 .setFontLine('underline');
    value = class.getRange(row+1,2+i).getValue()+" "+class.getRange(row+3,2+i).getValue();
    agenda.getRange(currentRow+1,1).setValue(value);
    agenda.getRange(currentRow+2,1).setValue(class.getRange(row+2,2+i).getValue());
    currentRow += 4;
  }
}

function ptAgenda() {
  var pt = SpreadsheetApp.getActive().getSheetByName('PT');
  var agenda = SpreadsheetApp.getActive().getSheetByName('Agenda');
  var sanMarino = [], other = [];
  var time, student, teacher, room, agendaItem;
  for (var i=0; i<pt.getLastRow()-1; i++) {
    time = pt.getRange(i+2,5).getValue();
    student = pt.getRange(i+2,4).getValue();
    teacher = pt.getRange(i+2,6).getValue();
    room = pt.getRange(i+2,9).getValue();
    agendaItem = time+" "+student+" ("+teacher+") "+room;
    if (pt.getRange(i+2,7).isBlank()) {
      sanMarino.push([agendaItem]);
    }
    else other.push([agendaItem]);
  }
  if (sanMarino.length>0) agenda.getRange(4,2,sanMarino.length,1).setValues(sanMarino);
  if (other.length>0) {
    agenda.getRange(4+sanMarino.length,2).setValue("Arcadia/Online")
                                         .setFontLine("underline")
                                         .setFontWeight("bold");
    agenda.getRange(4+sanMarino.length+1,2,other.length,1).setValues(other);
  }
}

function dateFormat(date) {
  var month = date.getMonth()+1;
  var day = date.getDate();
  return month+"/"+day;
}

function agendaFormat() {
  var pt = SpreadsheetApp.getActive().getSheetByName('PT');
  var cc = SpreadsheetApp.getActive().getSheetByName('CC');
  var agenda = SpreadsheetApp.getActive().getSheetByName('Agenda');
  var date = agenda.getRange(1,1).getValue();
  var rowCounter,formula;
  //set up agenda header
  var header = [["Class","Private Tutor","Counseling Meeting","To Do","Initial Meeting"]];
  agenda.getRange(3,1,1,5).setValues(header)
                          .setHorizontalAlignment("center")
                          .setFontLine("underline");
  //populate all classes of the day
  fetchClasses();
  //populate all PT sessions by formula
  rowCounter = pt.getLastRow()-1;
  if (rowCounter > 0) {
    ptAgenda();
//    formula = '=PT!$E2&" "&PT!$D2&" ("&PT!$F2&") "&PT!$I2';
//    agenda.getRange(4,2).setFormula(formula);
//    agenda.getRange(4,2).copyTo(agenda.getRange(4,2,rowCounter));
  }
  //populate all CC sessions by formula
  rowCounter = cc.getLastRow()-1;
  if (rowCounter > 0) {
    formula = '=CC!$E2&" "&CC!$D2&" ("&CC!$F2&") "&CC!$J2';
    agenda.getRange(4,3).setFormula(formula);
    agenda.getRange(4,3).copyTo(agenda.getRange(4,3,rowCounter));
  }
  //set up column width and borders
  agenda.getRange(2,1,agenda.getLastRow()-1,5).setBorder(true, true, true, true, true, false);
  agenda.getRange(3,1,1,5).setBorder(true, true, true, true, true, true);
  agenda.setColumnWidths(1, 5, 150);
  agenda.autoResizeColumns(1,3);
  var agendaTitle = "SM Daily Agenda "+dateFormat(date)+" "+dateToDay(date);
  agenda.getRange(2,1,1,5).merge().setValue(agendaTitle)
                          .setHorizontalAlignment("center")
                          .setFontSize(16);
}

function finishAgenda() {
  var pt = SpreadsheetApp.getActive().getSheetByName('PT');
  var cc = SpreadsheetApp.getActive().getSheetByName('CC');
  var room = SpreadsheetApp.getActive().getSheetByName('Room');
  var agenda = SpreadsheetApp.getActive().getSheetByName('Agenda');
  room.getRange(2,2,27,12).clear();
  agenda.getRange(2,1,agenda.getMaxRows()-1,5).clear();
  agenda.getRange(2,1,agenda.getMaxRows()-1,5).setNumberFormat("@");
  var date = agenda.getRange(1,1).getValue();
  promptCheck(date.toDateString());
  //Arrange Classes
  sortListbyTeacherClassCount();
  arrangeClass(date);
  //Arrange CC rooms
  cc.getRange(2,1,cc.getLastRow()-1,cc.getLastColumn()).sort(6);
  getRyan();
  roomArrange(cc,true);
  //Arrange PT rooms
  roomArrange(pt,false);
  agendaFormat();
}








