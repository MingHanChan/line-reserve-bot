# LINE-reservervation-bot
這是一個以LINE BOT與Google試算表為基礎的線上LINE預約系統
在一間單位、小公司、或者個人如果遇到有要開放讓其他人預約時段，像是預約團體辦理案件、或是預約工作時段等等，一般都會想到用手動在Excel工作表裡做紀錄。
最近在研究LINE BOT文件時想到可以用LINE BOT搭配Google Apps Script操作Google Sheets來達到一個線上LINE機器人預約系統，讓一般的顧客(民眾)可以透過加入我們的LINE官方帳號來做一個預約時段的動作；此系統根據每一個LINE帳號先要使用者註冊會員資訊，註冊完後即可開放預約，另外，關於LINE 機器人在連續對談中要如何記錄每個使用者的狀態是一個重要的問題，我這邊透過在每個註冊會員資料中紀錄一個狀態碼來作狀態的識別，整個預約系統流程如下圖所示
![](https://github.com/MingHanChan/line-reserve-bot/blob/master/reservation.jpg)
