# LINE-reservervation-bot
這是一個以LINE BOT與Google試算表為基礎的線上LINE預約系統
在一間單位、小公司、或者個人如果遇到有要開放讓其他人預約時段，像是預約團體辦理案件、或是預約工作時段等等，一般都會想到手動在Excel工作表裡做紀錄。
最近在研究LINE BOT文件時想到可以用LINE BOT搭配Google Apps Script操作Google Sheets來當作後台的資料庫紀錄功能，達到一個線上LINE機器人預約系統，讓一般的顧客(民眾)可以透過加入我們的LINE官方帳號來做預約時間。此系統根據每一個LINE帳號先要使用者註冊會員資訊，註冊完後即可開放預約，另外，關於LINE 機器人在連續對談中要如何記錄每個使用者的狀態是一個重要的問題，這邊的話透過在每個註冊會員資料中紀錄一個狀態碼來作狀態的識別，整個預約系統流程如下圖所示。
<center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/reservation.jpg" style="zoom:50%"/>預約流程</center>


## 整個預約系統大致可分為四個部分

### <span id="1">1. 提出預約</span>
 當使用者提出預約時，後端(Google apps script)會根據使用者唯一的LINE userId來去我們的註冊資訊表(Google sheets)判斷是否有註冊過，如果有則接續3，若無則傳送訊息要求使用者註冊，即[第二部分](#2)
 <center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/img/IMG_2149.PNG" style="zoom:50%"/>預約前須先註冊</center>
 <center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/img/IMG_2151.PNG" style="zoom:50%"/>利用LIFF註冊頁面</center>
 
### <span id="2">2. 使用者註冊</span>
關於讓使用者註冊的程序，我這邊是使用LINE的LIFF(LINE Front-end Framework)。LIFF是LINE在2018年6月新推出的技術，是一個可以在 LINE app 內運作的 web app 平台，可以讓使用者在對話視窗中，不需要另外加入 bot 就直接使用，想要了解更多的話可以參考這邊他們的[官方文件](https://developers.line.biz/en/docs/liff/overview/)。輸入完註冊資訊後如註冊成功會發訊息通知使用者。
<center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/img/IMG_2152.PNG" style="zoom:50%"/></center>

### <span id="3">3. 提出預約</span>
註冊過後的使用者提出預約申請之後，要求使用者輸入要預約的件數，並根據此人數來找尋後台的Google Sheets可預約時段，並且因為我們有狀態的紀錄，因此使用者在此階段如果沒有輸入正確的預約人數，或者只用者想要更改預約人數，或者想取消此次預約，Bot都會根據狀態訊息來做修改。找到可預約時段後，用LINE的Flex Messages來傳送可預約時段，讓使用者可一目了然且操作方便。
<center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/img/IMG_2149.PNG" style="zoom:50%"/>提出預約</center>
<center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/img/IMG_2153.PNG" style="zoom:50%"/>輸入預約人數</center>
<center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/img/IMG_2155.PNG" style="zoom:50%"/>可預約時間</center>
<center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/img/IMG_2156.PNG" style="zoom:50%"/>可修改預約人數或[取消預約](#cancel)</center>
<center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/img/IMG_2157.PNG" style="zoom:50%"/>不正確的預約資訊</center>

### <span id="4">4. 預約結果</span>
 點選Flex Messages上的訊息來選定預約的時間，若成功預約寫入試算表後，回傳確認訊息給使用者。

在<span id="cancel">取消預約</span>方面，使用者提出取消預約，系統會找出此註冊帳號已預約的時段，同樣的以Flex Messages訊息回傳讓使用者點選，點選完後若成功則回傳取消成功的訊息
<center><img src="https://github.com/MingHanChan/line-reserve-bot/blob/master/img/IMG_2160.PNG" style="zoom:50%"/>預約確認&取消預約</center>
