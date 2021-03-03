/*
 * This is an Task Reminder Bot for Slack API
 * Shota Niimi <http://uznetro.net/>
 */

//global variable
var ss = SpreadsheetApp.getActive();//シート取得
var sheet = SpreadsheetApp.getActiveSheet();

function pf_mailAlert() {
  //セルの値を取得して自動メール送信
  //※GmailApp.sendEmailの送信元仕様について
  //送信元はシートオーナー／スクリプトオーナーの両方になる様（両者が違う場合はそれぞれ送信実行される）
  //Fromはオプション指定できるがそのアカウントでエイリアス設定されたアドレスでのみ送信可能。

  //年月日の定義
  var nowdate = new Date();
  var dateString = Utilities.formatDate(nowdate,"JST","yyyy/MM/dd");
  var lastRow = sheet.getLastRow();
  var rangeC=sheet.getRange("C2:C");
  var rangeD=sheet.getRange("D2:D");

  var mailTo = "niimi_shota@test.net";
  var webHookUrl = "https://hooks.slack.com/services/*******";
  var ssManegemntUrl = "https://docs.google.com/spreadsheets/d/*******";
  for (var i = lastRow; i > 0; i--){
    var checkA = sheet.getRange(i, 1).getValue();//判定する列「プラン」を取得
    var checkD = sheet.getRange(i, 4).getValue();//判定する列「実施終了日」を取得
    var checkC = sheet.getRange(i, 3).getValue();//判定する列「受付終了日」を取得
    var tanto = sheet.getRange(i, 6).getValue();//担当者
    // スプレッドシートの日付項目を文字列として扱う
    rangeD.setNumberFormat('@');
    rangeC.setNumberFormat('@');
    //Moment変換
    var checkDateD = Moment.moment(checkD);
    var checkDateC = Moment.moment(checkC);

    //DEBUG
    //Logger.log(checkDate.format("YYYY/MM/DD"));
    //Logger.log(Moment.moment(dateString).format("YYYY/MM/DD") +  " @@@ " + checkDate.format("YYYY/MM/DD"));
    //Logger.log("加算" + Moment.moment(dateString).add(1,'d').format("YYYY/MM/DD"));

    //▼受付終了日判定
    if(Moment.moment(checkDateC).format("YYYY/MM/DD") === Moment.moment(dateString).format("YYYY/MM/DD")){
      //終了日
      //Logger.log(checkDate.format("YYYY/MM/DD") + "当日です");

      var messegeTitle1 = "【運用報告：受付終了】" + checkA;
      var messegeBody1  = "施策：" + checkA + "　受付終了日になりました。\n担当：" + tanto + "\n受付終了日：" + Moment.moment(checkDateC).format("YYYY/MM/DD");

      GmailApp.sendEmail(mailTo,messegeTitle1,messegeBody1);
      sheet.getRange(i, 3).setBackground("#d3d3d3");
      sheet.getRange(i, 5).setValue("受付終了 >> 配信済");
      //Logger.log("送信確認：受付終了＞＞");

      //slack webhook
      var options = {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : JSON.stringify(
          {
            "text": ":apple: *Scheduled Messages,*",
            "blocks": [],
            "attachments": [
              {
                "color": "#009688",
                "blocks": [
                  {
                    "type": "section",
                    "text": {
                      "type": "mrkdwn",
                      "text": messegeBody1,
                    }
                  },
                  {
                    "type": "divider"
                  },
                  {
                    "type": "section",
                    "text": {
                      "type": "mrkdwn",
                      "text": "<https://docs.google.com/spreadsheets/d/*******|管理表「GoogleSpreadSheet」リンク>",
                    }
                  }
                ]
              }
            ]
          }
        )
      };
      UrlFetchApp.fetch(webHookUrl, options);
    }

    //▼実施終了日判定
    if(Moment.moment(checkDateD).format("YYYY/MM/DD") === Moment.moment(dateString).format("YYYY/MM/DD")){
      //終了日
      //Logger.log(checkDate.format("YYYY/MM/DD") + "当日です");
      var messegeTitle3 = "【運用報告：実施終了】" + checkA;
      var messegeBody3  = "施策：" + checkA + "　が実施終了日になりました。\n担当：" + tanto + "\n実施終了日：" + Moment.moment(checkDateD).format("YYYY/MM/DD");

      GmailApp.sendEmail(mailTo,messegeTitle3,messegeBody3);
      sheet.getRange(i, 3).setBackground("#d3d3d3");
      sheet.getRange(i, 5).setValue("実施終了 >> 配信済");
      //Logger.log("送信確認：実施終了＞＞");

      //slack webhook
      var options = {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : JSON.stringify(
          {
            "text": ":apple: *Scheduled Messages,*",
            "blocks": [],
            "attachments": [
              {
                "color": "#f57c00",
                "blocks": [
                  {
                    "type": "section",
                    "text": {
                      "type": "mrkdwn",
                      "text": messegeBody3,
                    }
                  },
                  {
                    "type": "divider"
                  },
                  {
                    "type": "section",
                    "text": {
                      "type": "mrkdwn",
                      "text": "<https://docs.google.com/spreadsheets/d/*******|PJ管理表「GoogleSpreadSheet」リンク>",
                    }
                  }
                ]
              }
            ]
          }
        )
      };
      UrlFetchApp.fetch(webHookUrl, options);
    }
  }
}
