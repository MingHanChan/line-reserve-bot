//critical section 測試範例  避免平行寫入問題
//    var param = e.parameter;
//    var msg = param.msg;  
//    var lock = LockService.getScriptLock();
//    if( lock.tryLock(5000)){ 要開始前先拿到lock 且只有拿到lock才可以進入 才不會有同時檢查儲存格為空而同時又有同時寫入造成無法預期的錯誤 
//      var ssTitle = '{公司單位名稱}\t2019年9月';
//      var spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
//      var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
//      var sheet = spreadsheet.getSheetByName('2');
//      if(sheet.getRange(5, 2).getValue()=='')
//        sheet.getRange(5, 2).setValue(msg);
//      else
//        console.log('write FAILED,  message : '+msg);
//    } 
function doPost(e){
  var urlReply = 'https://api.line.me/v2/bot/message/reply';
  var userId = '';
  var replyToken = '';
  var CHANNEL_ACCESS_TOKEN = '';
  var recieveData= JSON.parse(e.postData.contents);
  if(recieveData.events[0].replyToken == '00000000000000000000000000000000')
    return ;
  if (recieveData.events[0].type == "message"){
    userId = recieveData.events[0].source.userId;
    replyToken = recieveData.events[0].replyToken;
    var receiveMsg = recieveData.events[0].message.text;
    var replyMsg = '';
    switch(receiveMsg){
      case '預約團件':
        if(isRegisted(userId)){
          //已註冊 回傳可預約天數 
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('list');
            var row = sheet.createTextFinder(userId).findNext().getRow();
            sheet.getRange(row, 5).setValue('11');//修改狀態為等待輸入人數(11)
         }
        else{
        replyMsg = '你尚未註冊，請先按以下連結註冊\n line://app/1618221073-AN84GWVX';
          UrlFetchApp.fetch(urlReply, {
            'headers': {
              'Content-Type': 'application/json; charset=UTF-8',
              'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
            },
            'method': 'post',
            'payload': JSON.stringify({
              'replyToken': replyToken,
              'messages': [{
                'type': 'text',
                'text': replyMsg,
              }],
            }),
         });
        }
        break;
      case '取消預約':
        if(isRegisted(userId)){
          //已註冊 回傳可預約天數 
        }
        else{
        replyMsg = '不好意思，沒有您的資料喔';
          UrlFetchApp.fetch(urlReply, {
            'headers': {
              'Content-Type': 'application/json; charset=UTF-8',
              'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
            },
            'method': 'post',
            'payload': JSON.stringify({
              'replyToken': replyToken,
              'messages': [{
                'type': 'text',
                'text': replyMsg,
              }],
            }),
         });
        }
        break;
      case '取消註冊':
        if(isRegisted(userId)){
          //已註冊 確認是否取消(flex msg)
        }
        else{
        replyMsg = '不好意思，沒有您的資料喔';
          UrlFetchApp.fetch(urlReply, {
            'headers': {
              'Content-Type': 'application/json; charset=UTF-8',
              'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
            },
            'method': 'post',
            'payload': JSON.stringify({
              'replyToken': replyToken,
              'messages': [{
                'type': 'text',
                'text': replyMsg,
              }],
            }),
         });
        }
        break;
      case '我要註冊':
        if(isRegisted(userId)){
        replyMsg = '您已經註冊!';
          UrlFetchApp.fetch(urlReply, {
            'headers': {
              'Content-Type': 'application/json; charset=UTF-8',
              'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
            },
            'method': 'post',
            'payload': JSON.stringify({
              'replyToken': replyToken,
              'messages': [{
                'type': 'text',
                'text': replyMsg,
              }],
            }),
         });
        }
        else{
          var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
          var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
          var sheet = spreadsheet.getSheetByName('list');
          sheet.appendRow(['','','',userId,00])//新增未註冊之userId
          replyMsg = '請按以下網址連結註冊\n line://app/1618221073-AN84GWVX';
          UrlFetchApp.fetch(urlReply, {
            'headers': {
              'Content-Type': 'application/json; charset=UTF-8',
              'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
            },
            'method': 'post',
            'payload': JSON.stringify({
              'replyToken': replyToken,
              'messages': [{
                'type': 'text',
                'text': replyMsg,
              }],
            }),
         });
        }
        break;
      default://非格式內的訊息
        var status = status(userId);
        switch(status){
          case false://一般民眾的一般訊息
            //隨機回應
            break;
          case '00'://已申請註冊 尚未填寫註冊訊息
            var registInfo = receiveMsg.split('\n');
            var company = registInfo[0];
            var name = registInfo[1];
            var phone = registInfo[2];
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('list');
            var row = sheet.createTextFinder(userId).findNext().getRow();
            sheet.getRange(row, 1).setValue(company);
            sheet.getRange(row, 2).setValue(name);
            sheet.getRange(row, 3).setValue(phone);
            sheet.getRange(row, 5).setValue('01');
            break;
          case '11'://使用者提出預約請求,使用者輸入預約人數
            var num = Math.ceil(parseInt(receiveMsg)/10);
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('list');
            var row = sheet.createTextFinder(userId).findNext().getRow();
            var company = sheet.getRange(row, 1);
            var avilableDays = checkAvailableDays(num,company);
            //(以flex msg 傳送回復訊息)
            
            sheet.getRange(row, 5).setValue('10');
            break;
          case '10'://輸入完人數 , 接著輸入要預約時間
            if(reserve(receiveMsg)){
             //回傳預約成功訊息 並修改狀態為已註冊(01)
              
            }
            else{
             //回傳預約失敗訊息 修改狀態為已註冊(01)
              
            }
            break;
          case '01'://仲介輸入非格式化的訊息
            
            break;
    }
   }
  }
}
//
function status(userId){
  if(!isRegisted(userId))
    return false;
  var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
  var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
  var sheet = spreadsheet.getSheetByName('list');
  var result = sheet.createTextFinder(userId).findNext();
  var row = result.getRow();
  return sheet.getRange(row, 5);//回傳狀態碼
  
}
function myFunction() {
      var ssTitle = '{公司單位名稱}\t2019年9月';
      var spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
      var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
      var sheet = spreadsheet.getSheetByName('2');
      sheet.appendRow(['','',2,4])
      Logger.log(sheet.createTextFinder('C').findNext());
//      if(sheet.getRange(5, 2).getValue()=='')
//        Logger.log('empty string');
//      if(sheet.getRange(5, 2).getValue()==null)
//        Logger.log('null');
//      
//      var ssTitle = '{公司單位名稱}\t2019年9月';
//      var spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
//      var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
//      var sheet = spreadsheet.getSheetByName('9');
//  var range = sheet.getRange('B33:B37');
//      range.setBorder(true, true, true, true, true, true);
//      var rangestr = range[0].getMergedRanges()[0].getA1Notation();
//      var str = 'B' + rangestr[1] + ':' + 'B' +  rangestr[4];
//      sheet.getRange(str).breakApart();
      //range[0].setValue('');

//      var merge = range[0].getMergedRanges();
//      for (var i = 0; i < merge.length; i++) {
//        Logger.log(merge[i].getA1Notation());
//        Logger.log(merge[i].getDisplayValue());
//      }
     // Logger.log(cancel('2019/9/9 09:00~11:00','恆心'));
}

function cancel(cancelTime,company){//finished success return true;otherwise false.
  //cancelTime = '2019/9/9 09:00~11:00' , company = '要取消的仲介名稱'
  cancelTime =  cancelTime.split(' ')
  var date = cancelTime[0];
  var time = cancelTime[1].split('~');
  var lock = LockService.getScriptLock();
  if( lock.tryLock(3000) ){ 
    var date = cancelTime[0].split("/"); //["2019", "9", "9"]
    var ssTitle = '{公司單位名稱}\t'+ date[0] + '年' + date[1] + '月';
    var spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
    var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
    var sheet = spreadsheet.getSheetByName(date[2]);
    var timeslot = sheet.createTextFinder(company).findAll();
    //刪除預約時段
    for(var index = 0 ; index < timeslot.length ; index++){
      var rangestr = timeslot[index].getMergedRanges()[0].getA1Notation();//合併的儲存格範圍 ex."B6:B9"
      var row1 = rangestr.substr(rangestr.indexOf('B')+1,rangestr.indexOf(':')-1);
      var row2 = rangestr.substr(rangestr.lastIndexOf('B')+1,rangestr.length-1);
      var numRows = row2-row1+1;
      for(var col = 2 ; col <= 5 ; col++){
       sheet.getRange(row1, col).setValue('');
       sheet.getRange(row1, col, numRows).breakApart().setBorder(true, true, true, true, true, true);
      }
      for(var row = row1 ; row <= row2 ; row++){
        sheet.getRange(row, 5).setValue('F'); 
      }
    }
    return true;
  }
  else
    return false;
}
//顯示要取消預約的時間
function showCancelTime(reserveAgency){//reserveAgency = '仲介名稱'
  //finished
  var d = new Date();
  var reserveTime = [];
  for(var i = 1 ; i <= 14 ; i++){
    d.setDate(d.getDate()+1);
    if(0 < d.getDay() & d.getDay() < 6){
      var year =  d.getFullYear();
      var month = d.getMonth()+1;
      var date = d.getDate();
      var ssTitle = '{公司單位名稱}\t'+ year + '年' + month + '月';
      var spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
      var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
      var sheet = spreadsheet.getSheetByName(date);
      var timeslot = sheet.createTextFinder(reserveAgency).findAll();//仲介有預約的時段
      if(timeslot.length==0){
        continue;
      }
      else{
        for(var index = 0 ; index < timeslot.length ; index++){
          var rangestr = timeslot[index].getMergedRanges()[0].getA1Notation();//合併的儲存格範圍 ex."B6:B9"
          var row1 = rangestr.substr(rangestr.indexOf('B')+1,rangestr.indexOf(':')-1);
          var row2 = rangestr.substr(rangestr.lastIndexOf('B')+1,rangestr.length-1);
          var start = sheet.getRange(row1,7).getValue().split('~')[0];
          var end = sheet.getRange(row2,7).getValue().split('~')[1];
          var time = start + '~' + end;
          reserveTime.push(year +'/'+ month +'/'+ date +' '+ time);
        }
      }
    }
  }
  return reserveTime;
}
function regist(){
  //利用表單post spreadsheet，讓使用者可以一次輸入完註冊資訊
  //假日研究LIFF
  
}

function reserve(reserveTime){//finished
  // reserveTime = '2019/9/2 13:00~14:30 B' 成功預約回傳true 否則flase
  //critical section， 預約資訊寫入試算表
  var lock = LockService.getScriptLock();
  if( lock.tryLock(3000) ){ 
    var reserveTime = reserveTime.split(" ");
    var date = reserveTime[0].split("/"); //["2019", "8", "31"]
    var ssTitle = '{公司單位名稱}\t'+ date[0] + '年' + date[1] + '月';
    var spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
    var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
    var sheet = spreadsheet.getSheetByName(date[2]);
    var start = reserveTime[1].split("~")[0];
    var end = reserveTime[1].split("~")[1];
    var startRow,endRow;
    if(reserveTime[2]=='A'){
      startRow = sheet.createTextFinder(start+'~').findAll()[0].getRow();
      endRow = sheet.createTextFinder('~'+end).findAll()[0].getRow();
    }
    else{
      startRow = sheet.createTextFinder(start+'~').findAll()[1].getRow();
      endRow = sheet.createTextFinder('~'+end).findAll()[1].getRow();
    }
    var rownum = endRow-startRow+1;
    //檢查要預約的格子都是空的
    for(var i = 0; i < rownum ; i++){
      if(sheet.getRange(startRow+i, 2).getValue()!='')
        return false;
    }
    //預約單位
    sheet.getRange(startRow,2,rownum).merge().setValue('要預約的仲介');
    //聯絡人
    sheet.getRange(startRow,3,rownum).merge().setValue('聯絡人');
    //連絡電話
    sheet.getRange(startRow,4,rownum).merge().setValue('電話');
    //預約人數
    sheet.getRange(startRow,5,rownum).merge().setValue('預約人數');
    //實到人數(免填)要合併
    sheet.getRange(startRow,6,rownum).merge();
    return true;
  }
  else
    return false;
}
function checkAvailableDays(number,reserveAgency){//finished
  //傳入參數number為需要預約的格數 一格最多預約10人,回傳後兩周所有可預約的時間或者false無可用時間
  //最多單次預約人數為50人
  if(number>5)
    return false;
  //檢查兩周內的可預約日期跟時間
  var date = new Date();
  var availableDay = [];
  var timeSlot = [];
  for(var i=1;i<=14;i++){
    //往後推兩個禮拜週一到週五，需考慮跨月份跟跨年份問題
    date.setDate(date.getDate()+1);
    if(0 < date.getDay() & date.getDay() < 6){
      var ssTitle = '{公司單位名稱}\t'+ date.getFullYear() + '年' + (date.getMonth()+1) + '月';
      var spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
      var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
      var sheet = spreadsheet.getSheetByName(date.getDate());
      var isreserved = sheet.createTextFinder(reserveAgency).findAll();//是否有仲介預約同一天
      if(isreserved.length>0)
        continue;
      
      var range = sheet.createTextFinder('F').findAll();//沒有被預約時段
      //如果時段多餘要預約的人數在找
      if(range.length>=number){
        var index = false;
        for(var j=0;j<range.length;j++){
          var interval=0;
          while((range[j+1]!=undefined) && range[j+1].getRow()-range[j].getRow()==1){
            interval++;
            j++;
          }
          if((number-1)==interval){
            index = j-interval;
            //可預約時段
            var row1 = range[index].getRow();
            var row2 = range[index+interval].getRow()
            var start = sheet.getRange(row1, 7).getDisplayValue().split("~")[0];
            var end =  sheet.getRange(row2, 7).getDisplayValue().split("~")[1];
            var time = '';
            //分為AB兩櫃台時段
            if(row1<30)
              time = date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate() + ' ' + start + "~" + end + ' A';//可用時段&櫃台 ex.2019/8/31 13:30~14:00 A
            else
               time = date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate() + ' ' + start + "~" + end + ' B';//可用時段&櫃台 ex.2019/8/31 13:30~14:00 B   
            timeSlot.push(time);
          }
        }
      }
    }
  }
  if(timeSlot.length)
    return timeSlot;
  else
    return false;
}
function createNextMonth(){//finished
  //從'團件預約格式'複製每個月的格式
  var date = new Date();
  date.setMonth(date.getMonth()+1,1);//設定為下個月1號
  var ssTitle = '{公司單位名稱}\t'+ date.getFullYear() + '年' + (date.getMonth()+1) + '月';
  if(!DriveApp.getFilesByName(ssTitle).hasNext()){
    //還沒有試算表,複製一個新的spreadsheet
    var spreadsheetFile = DriveApp.getFilesByName('團件預約格式').next();
    var formatSpreadsheet = SpreadsheetApp.open(spreadsheetFile);
    var spreadsheet = formatSpreadsheet.copy(ssTitle);
    //creat the all days of month
    var totalDays = new Date(date.getFullYear(),(date.getMonth()+1),0).getDate();
    var day = ['星期一','星期二','星期三','星期四','星期五'];
    for(var i = 1 ; i <= totalDays ; i++){
      if(0 < date.getDay() & date.getDay() < 6){
        var sheetTitle = '{公司單位名稱}\t'+ date.getFullYear() + '年' + (date.getMonth()+1) + '月' + date.getDate() + '日 ' + day[date.getDay()-1];
        spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);
        var sheet = spreadsheet.duplicateActiveSheet().setName(i.toString());//依照日期創建一個sheet
        var range = sheet.getRange('A1');
        var richText = SpreadsheetApp.newRichTextValue()
                     .setText(sheetTitle)
                     .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build())// 設定粗體字
                     .build();
        range.setRichTextValue(richText);  
      }
      date.setUTCDate(i+1);//最後會是再下個月1號
    }
  }
  
}


function isRegisted(userId){//finished
 //check the 'HRAgency' spreadsheet, if registed, return true 
  var files = DriveApp.getFilesByName('Human Resources Agency');//取得雲端硬碟中的試算表
  var spreadsheet = SpreadsheetApp.open(files.next());//開啟試算表
  var sheet = spreadsheet.getSheetByName("Human Resources Agency");//開啟仲介資訊
  var textFinder = sheet.createTextFinder(userId);
  if(textFinder.findNext()==null)
    return false;
  else
    return true;
}
