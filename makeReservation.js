function doPost(e){
  var urlReply = 'https://api.line.me/v2/bot/message/reply';
  var userId = '';
  var replyToken = '';
  var CHANNEL_ACCESS_TOKEN = '8y4CMDlL1IZl8IbISUX0VnCwXabT+7G7xTnTtxnIzwzTSVwVZTPrT+7V1moZZs7aFBXqP34iK3u6QFZt96ty81Rsoe7gp3WSTtOlxH/p2i9W8TohClG2VecELklc6varHnhH9FSGrzK9fg81NGfOogdB04t89/1O/w1cDnyilFU=';
  var registerUrl = 'line://app/1618221073-AN84GWVX';
  var recieveData= JSON.parse(e.postData.contents);
  console.log(recieveData);
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
          //var statused = status(userId);
          //已註冊 詢問要預約件數
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('Human Resources Agency');
            var row = sheet.createTextFinder(userId).findNext().getRow();
            sheet.getRange(row, 5).setValue('\'11');//修改狀態為等待輸入人數(11)
            replyMsg ='請輸入您要預約的件數\n輸入完畢後請稍等幾分鐘，將為您查詢可預約的時間。';
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
          var sheet = spreadsheet.getSheetByName('Human Resources Agency');
          sheet.appendRow(['','','',userId,'none'])//新增未註冊之userId
          replyMsg = '您尚未註冊，請先按以下連結註冊\n' + registerUrl;
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
         //已註冊 回傳已經預約天數
        if(isRegisted(userId)){
          var cancelTime = showCancelTime(userId);//object bubble format or false
          //有找到已預約時段
          if(cancelTime!=false){
          UrlFetchApp.fetch(urlReply, {
              'headers': {
                'Content-Type': 'application/json; charset=UTF-8',
                'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
              },
              'method': 'post',
              'payload': JSON.stringify({
                'replyToken': replyToken,
                'messages': [{
                  'type': 'flex',
                  'altText': "取消預約時間",
                  'contents': cancelTime,
                }],
              }),
          });
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('Human Resources Agency');
            var row = sheet.createTextFinder(userId).findNext().getRow();
            sheet.getRange(row, 5).setValue('\'21');//修改狀態為取消預約(21)
         }
          else{
            replyMsg = '不好意思，沒有您的預約資料喔';
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
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('Human Resources Agency');
            var row = sheet.createTextFinder(userId).findNext().getRow();
            sheet.getRange(row, 5).setValue('\'01');//修改狀態為已註冊(01)
          }
        }
        else{
         //尚未註冊
        replyMsg = '不好意思，沒有您的註冊資料喔\n若要註冊，請按以下網址連結註冊\n' + registerUrl;
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
          if(status(userId) == false){
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('Human Resources Agency');
            sheet.appendRow(['','','',userId,'none'])//新增未註冊之userId
          }
        }
        break;
      case '取消註冊':
        if(status(userId)=='none'){
          var files = DriveApp.getFilesByName('Human Resources Agency').next();//取得雲端硬碟中的試算表
          var spreadsheet = SpreadsheetApp.open(files);//開啟試算表
          var sheet = spreadsheet.getSheetByName("Human Resources Agency");//開啟仲介資訊
          var textFinder = sheet.createTextFinder(userId);
          var row = textFinder.findNext().getRow();
          sheet.deleteRow(row);
        replyMsg = '謝謝您~已結束此次註冊流程';
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
          break;
        }
        if(status(userId) == false){
        replyMsg = '謝謝您~已結束此次註冊流程';
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
          break;
        }
        break;
      case '繼續註冊':
        var statused = status(userId);
        if(statused =='none'){
          replyMsg = '請按以下網址連結繼續註冊\n' + registerUrl;
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
          break;
        }
        else if(statused == false){
          var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
          var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
          var sheet = spreadsheet.getSheetByName('Human Resources Agency');
          sheet.appendRow(['','','',userId,'none'])//新增未註冊之userId
          replyMsg = '您尚未註冊，請先按以下連結註冊\n' + registerUrl;
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
          break;
        }
        else{
          replyMsg = '謝謝您~\n您已經註冊';
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
      case '取消預約流程':
        var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
        var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
        var sheet = spreadsheet.getSheetByName('Human Resources Agency');
        var row = sheet.createTextFinder(userId).findNext().getRow();
        sheet.getRange(row, 5).setValue('\'01');
        replyMsg ='好的，已結束此次預約流程';
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
        break;
      //非格式內的訊息
      default:
        var statused = status(userId);
        switch(statused){
          case false://一般民眾的一般訊息
            var msg = '您好~ \n'+'我們的服務時間為 \n星期一至星期五 08:00-17:00 中午不休息\n'+'電話：03-331-0409 (總機)';
            var msg1 = '平板免費借，每月再享10GB網路\n借用超簡單，宅配到您家！\nhttps://nit.immigration.gov.tw/activitypage/activitypage1.html';
            var msg2 = '新住民培力發展資訊網\n https://ifi.immigration.gov.tw/mp.asp?mp=ifi_zh';
            var msg3 = '新住民免費電腦課程!\n https://nit.immigration.gov.tw/';
            var msg4 = '防止人口販運，捍衛人權\n https://www.immigration.gov.tw/5385/7445/7535/';
            var msg5 = '跨國境媒合婚姻不能廣告\n切勿相信網路廣告!\n https://www.immigration.gov.tw/5385/7445/7621/';
            //更改訊息新增或刪減 下面兩行也要更改
            //更改完後要按"發布"->"部屬為網路應用程式"->專案版本選"新增"->按更新
            var ranMsg = [msg1,msg2,msg3,msg4,msg5];
            var ran = Math.floor(Math.random()*5);
            UrlFetchApp.fetch(urlReply, {
              'headers': {
                'Content-Type': 'application/json; charset=UTF-8',
                'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
              },
              'method':'post',
              'payload': JSON.stringify({
                'replyToken': replyToken,
                'messages': [{
                  type:'text',
                  text:ranMsg[ran]
                }]
              }),
            }); 
            break;
          case 'none'://已申請註冊 尚未填寫註冊訊息            
            var registInfo = receiveMsg.split('\n');
            if(registInfo.length == 3){
              var company = registInfo[0];
              var name = registInfo[1];
              var phone = registInfo[2];
            }
            else{
              var JOSNContent = continueRegisterJSON();
              UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
                'headers': {
                  'Content-Type': 'application/json; charset=UTF-8',
                  'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
                },
                'method': 'post',
                'payload': JSON.stringify({
                  'replyToken': replyToken,
                  'messages': [{
                    'type': 'flex',
                    'altText': "是否繼續註冊?",
                    'contents': JOSNContent,
                  }],
                }),
              });
              break;
            }
            //紀錄註冊資訊，並修改狀態碼為已註冊(01)
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('Human Resources Agency');
            var row = sheet.createTextFinder(userId).findNext().getRow();
            sheet.getRange(row, 1).setValue(company);
            sheet.getRange(row, 2).setValue(name);
            sheet.getRange(row, 3).setValue(phone);
            sheet.getRange(row, 5).setValue('\'01');
            
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
                  'text': '謝謝您~\n已幫您註冊\n下次可用該LINE帳號預約團件囉!',
                }],
              }),
            });
            break;
          case '11'://使用者提出預約請求,使用者輸入預約人數,回傳可預約時間
            //使用者輸入非數字字串
            if(isNaN(receiveMsg)){
              var JOSNContent = continueReserveJSON();
              UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
                'headers': {
                  'Content-Type': 'application/json; charset=UTF-8',
                  'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
                },
                'method': 'post',
                'payload': JSON.stringify({
                  'replyToken': replyToken,
                  'messages': [{
                    'type': 'flex',
                    'altText': "是否繼續預約?",
                    'contents': JOSNContent,
                  }],
                }),
              });
            }
            else{
            var num = Math.ceil(parseInt(receiveMsg)/10);
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('Human Resources Agency');
            var row = sheet.createTextFinder(userId).findNext().getRow();
            var company = sheet.getRange(row, 1);
            var avilableDays = checkAvailableDays(num,company,receiveMsg);
            //有預約時段可預約
            if(avilableDays != false){
              var back = backtrack();
              UrlFetchApp.fetch(urlReply, {
                'headers': {
                  'Content-Type': 'application/json',
                  'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
                },
                'method': 'post',
                'payload': JSON.stringify({
                  'replyToken': replyToken,
                  'messages': [
                    {
                      'type': 'flex',
                      'altText': "預約時間",
                      'contents': avilableDays,
                    },
                    {
                      'type': 'flex',
                      'altText': "上一步",
                      'contents': back,
                    },                  
                  ]
                 }),
              });
              sheet.getRange(row, 5).setValue('\'10');
            }
            //沒有預約時段可預約
            else{
              var back = backtrack();
              UrlFetchApp.fetch(urlReply, {
                'headers': {
                  'Content-Type': 'application/json',
                  'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
                },
                'method': 'post',
                'payload': JSON.stringify({
                  'replyToken': replyToken,
                  'messages': [
                    {
                      'type': 'text',
                      'text': "不好意思，沒有可預約的時間\n您可以預約其他件數或結束此次預約流程，謝謝~",
                    },
                    {
                      'type': 'flex',
                      'altText': "上一步",
                      'contents': back,
                    },                  
                  ]
                 }),
              });
            }
            }
            break;
          case '10'://輸入完人數 , 接著輸入要預約時間
            if(reserve(receiveMsg,userId)){
              var time = receiveMsg.split(' ');
              var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
              var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
              var sheet = spreadsheet.getSheetByName('Human Resources Agency');
              var row = sheet.createTextFinder(userId).findNext().getRow();
              sheet.getRange(row, 5).setValue('\'01');
              //回傳預約成功訊息 並修改狀態為已註冊(01)
              replyMsg = '謝謝您，已幫您預約 ' + time[0] + ' ' + time[1] + '\n若要取消預約，請至少提前一天!';
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
              //回傳預約失敗訊息 修改狀態為已註冊(01)
              replyMsg = '預約失敗，請重新預約';
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
            var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
            var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
            var sheet = spreadsheet.getSheetByName('Human Resources Agency');
            var row = sheet.createTextFinder(userId).findNext().getRow();
            sheet.getRange(row, 5).setValue('\'01');//修改狀態為已註冊(01)
            }
            break;
          //仲介輸入非格式化的訊息
          case '01':
            var name = getAgency(userId);
            replyMsg = name + ' 您好! ' + '請問您要預約團件還是取消預約呢?';
            var JOSNContent = wannaJSON(replyMsg);
              UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
                'headers': {
                  'Content-Type': 'application/json; charset=UTF-8',
                  'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
                },
                'method': 'post',
                'payload': JSON.stringify({
                  'replyToken': replyToken,
                  'messages': [{
                    'type': 'flex',
                    'altText': "請問您要預約團件還是取消預約呢?",
                    'contents': JOSNContent,
                  }],
                }),
              });
            break;
          case '21':
            var agency = getAgency(userId);
            if(cancel(receiveMsg,agency)){
              replyMsg = '您好~\n已幫您取消' + receiveMsg +'的預約!';
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
              var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
              var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
              var sheet = spreadsheet.getSheetByName('Human Resources Agency');
              var row = sheet.createTextFinder(userId).findNext().getRow();
              sheet.getRange(row, 5).setValue('\'01'); 
            }
            else{
              replyMsg = '取消預約失敗，請重新操作';
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
              var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
              var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
              var sheet = spreadsheet.getSheetByName('Human Resources Agency');
              var row = sheet.createTextFinder(userId).findNext().getRow();
              sheet.getRange(row, 5).setValue('\'01');
            }
            break;
		}	
    }
  }
}

function status(userId){
  var spreadsheetFile = DriveApp.getFilesByName('Human Resources Agency').next();
  var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
  var sheet = spreadsheet.getSheetByName('Human Resources Agency');
  var result = sheet.createTextFinder(userId).findNext();
  if(result == null)
    return false;
  else{
    var row = result.getRow();
    return sheet.getRange(row, 5).getValue();//回傳狀態碼
  }
}

function cancel(cancelTime,company){//finished success return true;otherwise false.
  //cancelTime = '2019/9/9 09:00~11:00 B' , company = '要取消的仲介名稱'
  cancelTime =  cancelTime.split(' ')
  //輸入錯誤訊息
  if(cancelTime.length != 3)
    return false;
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
function showCancelTime(userId){//  return object 'bubble'
  //finished
  var reserveAgency = getAgency(userId);
  var headertext = reserveAgency + '您好，您已預約時段如下'
  var d = new Date();
  //var reserveTime = [];
  var bubble = {};
  bubble.type = "bubble";
  bubble.size = "kilo";
  bubble.header = {};
  bubble.header.type = "box";
  bubble.header.layout = "vertical";
  bubble.header.contents = [{type:"text",text:headertext}];
  bubble.body = {};  
  bubble.body.type = "box";
  bubble.body.layout = "vertical";
  bubble.body.spacing = "md";
  bubble.body.contents = [];
      var ssTitle = '{公司單位名稱}\t'+ d.getFullYear() + '年' + (d.getMonth()+1) + '月';
      var spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
      var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
      var nowMonth = d.getMonth();
      var nowYear = d.getFullYear();
  for(var i = 1 ; i <= 14 ; i++){
    d.setDate(d.getDate()+1);
    if(0 < d.getDay() & d.getDay() < 6){
      var year =  d.getFullYear();
      var month = d.getMonth();
      var date = d.getDate();
      if((month-nowMonth)!=0 || (year-nowYear)!=0){
       createNextMonth();
       ssTitle = '{公司單位名稱}\t'+ year + '年' + (month+1) + '月';
       spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
       spreadsheet = SpreadsheetApp.open(spreadsheetFile);
      }
      var sheet = spreadsheet.getSheetByName(date);
      var timeslot = sheet.createTextFinder(reserveAgency).findAll();//仲介有預約的時段
      if(timeslot.length==0){
        continue;
      }
      else{
        var time1 = year + '/' + (month+1) + '/' + date
        for(var index = 0 ; index < timeslot.length ; index++){
          var rangestr = timeslot[index].getMergedRanges()[0].getA1Notation();//合併的儲存格範圍 ex."B6:B9"
          var row1 = rangestr.substr(rangestr.indexOf('B')+1,rangestr.indexOf(':')-1);
          var row2 = rangestr.substr(rangestr.lastIndexOf('B')+1,rangestr.length-1);
          var start = sheet.getRange(row1,7).getValue().split('~')[0];
          var end = sheet.getRange(row2,7).getValue().split('~')[1];
          var body_contents = {
              type: "button",
              style: "primary",
              action: {
                type: "message"
              }
          };
          var time2;
          if(row1 < 30)
            time2 = time1 + ' '+ start + '~' + end +' A';
          else
            time2 = time1 + ' '+ start + '~' + end +' B';
          body_contents.action.label = time2
          body_contents.action.text = time2
          bubble.body.contents.push(body_contents);
          //reserveTime.push(year +'/'+ month +'/'+ date +' '+ time);
        }
      }
    }
  }
  if(bubble.body.contents.length>0)
    return bubble;
  else
    return false;
}
function reserve(reserveTime,userId){//finished
  // reserveTime = '2019/9/2 13:00~14:30 B 25' 成功預約回傳true 否則flase
  //critical section， 預約資訊寫入試算表
  var lock = LockService.getScriptLock();
  if( lock.tryLock(3000) ){ 
    var reserveTime = reserveTime.split(" ");
    if(reserveTime.length != 4)
      return false;
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
    var agency = getAgency(userId);
    var name = getUser(userId);
    var phone = getPhone(userId);
    var num = reserveTime[3];
    //預約單位
    sheet.getRange(startRow,2,rownum).merge().setValue(agency);
    //聯絡人
    sheet.getRange(startRow,3,rownum).merge().setValue(name);
    //連絡電話
    sheet.getRange(startRow,4,rownum).merge().setValue(phone);
    //預約人數
    sheet.getRange(startRow,5,rownum).merge().setValue(num);
    //實到人數(免填)要合併
    sheet.getRange(startRow,6,rownum).merge();
    return true;
  }
  else
    return false;
}
function checkAvailableDays(number,reserveAgency,people){//finished  return object 'carousel' 
  //傳入參數number為需要預約的格數 一格最多預約10人,回傳後兩周所有可預約的時間或者false無可用時間
  //最多單次預約人數為50人
  if(number>5)
    return false;
  //檢查兩周內的可預約日期跟時間
  var date = new Date();
  var availableDay = [];
  var timeSlot = [];
  var carousel = {};
  carousel.type = "carousel";
  carousel.contents = [];
  var ssTitle = '{公司單位名稱}\t'+ date.getFullYear() + '年' + (date.getMonth()+1) + '月';
  var spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
  var spreadsheet = SpreadsheetApp.open(spreadsheetFile);
  var nowMonth = date.getMonth();
  var nowYear = date.getFullYear();
  for(var i=1;i<=14;i++){
    //往後推兩個禮拜週一到週五，需考慮跨月份跟跨年份問題
    date.setDate(date.getDate()+1);
    var day = date.getDay();
    if(0 < day & day < 6){
      if((date.getMonth()-nowMonth)!=0 || (date.getFullYear-nowYear)!=0){
       createNextMonth();
       ssTitle = '{公司單位名稱}\t'+ date.getFullYear() + '年' + (date.getMonth()+1) + '月';
       spreadsheetFile = DriveApp.getFilesByName(ssTitle).next();
       spreadsheet = SpreadsheetApp.open(spreadsheetFile);
      }
      var sheet = spreadsheet.getSheetByName(date.getDate());  
      var isreserved = sheet.createTextFinder(reserveAgency).findAll();//是否有仲介預約同一天
      if(isreserved.length>0)
        continue;
      var bubble = {};
      bubble.type = "bubble";
      bubble.size = "kilo";
      bubble.header = {};
      bubble.header.type = "box";
      bubble.header.layout = "vertical";
      
      bubble.body = {};
      bubble.body.type = "box";
      bubble.body.layout = "vertical";
      bubble.body.spacing = "md";
      bubble.body.contents = [];
      var range = sheet.createTextFinder('F').findAll();//沒有被預約時段
      //如果時段多餘要預約的人數在找
      if(range.length>=number){
        for(var j=0;j<range.length;j++){
          var cells=1;
          var a = j;
          do{
            //確定有預約時段可以預約
            if(cells == number){
            //可預約時段
            var row1 = range[j].getRow();
            var row2 = range[j+cells-1].getRow()
            var start = sheet.getRange(row1, 7).getDisplayValue().split("~")[0];
            var end =  sheet.getRange(row2, 7).getDisplayValue().split("~")[1];
            var time = '';
            
            var body_contents = {
              type: "button",
              style: "primary",
              action: {
                type: "message"
              }
            };
            time = date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate();
            bubble.header.contents = [{type:"text",text:time}];
            //分為AB兩櫃台時段
            if(row1<30){
              body_contents.action.label = start + "~" + end + ' A';
              body_contents.action.text = time + ' ' + start + "~" + end + ' A ' + people;
              //time = date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate() + ' ' + start + "~" + end + ' A';//可用時段&櫃台 ex.2019/8/31 13:30~14:00 A 
            }
            else{
              //time = date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate() + ' ' + start + "~" + end + ' B';//可用時段&櫃台 ex.2019/8/31 13:30~14:00 B   
              body_contents.action.label = start + "~" + end + ' B';
              body_contents.action.text = time + ' ' + start + "~" + end + ' B ' + number;
            }
            bubble.body.contents.push(body_contents);
            //timeSlot.push(time);
              break;
            }
			cells++;
          }while((range[++a]!=undefined) && (range[a].getRow()-range[a-1].getRow())==1);
        }
      } 
      if(bubble.body.contents.length>0)
        carousel.contents.push(bubble);
    }//if(0 < date.getDay() & date.getDay() < 6)
  }
  if(carousel.contents.length>0){
    //carousel.contents.push(backtrack());
    return carousel;
  }
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
  var files = DriveApp.getFilesByName('Human Resources Agency').next();//取得雲端硬碟中的試算表
  var spreadsheet = SpreadsheetApp.open(files);//開啟試算表
  var sheet = spreadsheet.getSheetByName("Human Resources Agency");//開啟仲介資訊
  var textFinder = sheet.createTextFinder(userId).findNext();
  if(textFinder == null){
    return false;
  }
  else{
    var statused = sheet.getRange(textFinder.getRow(),5).getValue();
    if(statused == 'none')
      return false;
    else
      return true;
    
  }
}

function getAgency(userId){
  var files = DriveApp.getFilesByName('Human Resources Agency').next();//取得雲端硬碟中的試算表
  var spreadsheet = SpreadsheetApp.open(files);//開啟試算表
  var sheet = spreadsheet.getSheetByName("Human Resources Agency");//開啟仲介資訊
  var textFinder = sheet.createTextFinder(userId);
  var row = textFinder.findNext().getRow();
  return sheet.getRange(row, 1).getValue();
}
function getUser(userId){
  var files = DriveApp.getFilesByName('Human Resources Agency').next();//取得雲端硬碟中的試算表
  var spreadsheet = SpreadsheetApp.open(files);//開啟試算表
  var sheet = spreadsheet.getSheetByName("Human Resources Agency");//開啟仲介資訊
  var textFinder = sheet.createTextFinder(userId);
  var row = textFinder.findNext().getRow();
  return sheet.getRange(row, 2).getValue();
}
function getPhone(userId){
  var files = DriveApp.getFilesByName('Human Resources Agency').next();//取得雲端硬碟中的試算表
  var spreadsheet = SpreadsheetApp.open(files);//開啟試算表
  var sheet = spreadsheet.getSheetByName("Human Resources Agency");//開啟仲介資訊
  var textFinder = sheet.createTextFinder(userId);
  var row = textFinder.findNext().getRow();
  return sheet.getRange(row, 3).getValue();
}
function continueRegisterJSON(){
  var json = {
  "type": "bubble",
  "header": {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "輸入錯誤，是否繼續註冊?"
      }
    ]
  },
  "body": {
    "type": "box",
    "layout": "vertical",
    "spacing": "md",
    "contents": [
      {
        "type": "button",
        "style": "primary",
        "action": {
          "type": "message",
          "label": "繼續註冊流程",
          "text": "繼續註冊"
        }
      },
      {
        "type": "button",
        "style": "primary",
        "action": {
          "type": "message",
          "label": "取消註冊流程",
          "text": "取消註冊"
        }
      }
    ]
  }
}
  return json;
}
function continueReserveJSON(){
  var json = {
  "type": "bubble",
  "header": {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "輸入錯誤，是否繼續預約?"
      }
    ]
  },
  "body": {
    "type": "box",
    "layout": "vertical",
    "spacing": "md",
    "contents": [
      {
        "type": "button",
        "style": "primary",
        "action": {
          "type": "message",
          "label": "繼續預約流程",
          "text": "預約團件"
        }
      },
      {
        "type": "button",
        "style": "primary",
        "action": {
          "type": "message",
          "label": "取消預約流程",
          "text": "取消預約流程"
        }
      }
    ]
  }
}
  return json;
}
function wannaJSON(title){
  var json = {
  "type": "bubble",
  "header": {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": title
      }
    ]
  },
  "body": {
    "type": "box",
    "layout": "vertical",
    "spacing": "md",
    "contents": [
      {
        "type": "button",
        "style": "primary",
        "action": {
          "type": "message",
          "label": "預約團件",
          "text": "預約團件"
        }
      },
      {
        "type": "button",
        "style": "primary",
        "action": {
          "type": "message",
          "label": "取消預約",
          "text": "取消預約"
        }
      }
    ]
  }
}
  return json;
}
function backtrack(){
  var json = {
  "type": "bubble",
  "header": {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "上一步"
      }
    ]
  },
  "body": {
    "type": "box",
    "layout": "vertical",
    "spacing": "md",
    "contents": [
      {
        "type": "button",
        "style": "primary",
        "action": {
          "type": "message",
          "label": "重新輸入預約人數",
          "text": "預約團件"
        }
      },
      {
        "type": "button",
        "style": "primary",
        "action": {
          "type": "message",
          "label": "取消預約流程",
          "text": "取消預約流程"
        }
      }
    ]
  }
   
}
  return json;
}
