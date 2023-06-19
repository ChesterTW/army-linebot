


function doPost(e) {
    // 定義 LINE Messenging API Token
    var CHANNEL_ACCESS_TOKEN = 'ILvqdxQQmHF9ty52btDFfGGcsbQtJfaOOBZac4yOLiIftWweCDV4vTTEEJh1eg4kykpSurAa0eSQ8YxoIQ4UrkuUGDwlDk/hRvIwGGsy5qGBmsVU6BCyB/XFBZt+LhcA8eardYzq2plawkM+HxdBSAdB04t89/1O/w1cDnyilFU='; // LINE Bot API Token
    // 以 JSON 格式解析 User 端傳來的 e 資料
    var msg = JSON.parse(e.postData.contents);

    // for debugging
    Logger.log(msg);
    console.log(msg);

    /* 
    * LINE API JSON 解析資訊
    *
    * replyToken : 一次性回覆 token
    * userId : 使用者 user id，查詢 username 用
    * userMessage : 使用者訊息，用於判斷是否為預約關鍵字
    * eventType : 訊息事件類型
    */
    const replyToken = msg.events[0].replyToken;
    const userId = msg.events[0].source.userId;
    const userMessage = msg.events[0].message.text;
    const eventType = msg.events[0].source.type; 

    // 將接收到的文字，以斜線區分成陣列字串。
    const splitedUserMessage = userMessage.split('/');

    // 定義要使用的 Google Sheet，並使用語法 openByUrl 開啟。
    const sheetUrl = 'https://docs.google.com/spreadsheets/d/1l8ztOGnN6lG4YGkyfq-6VaF7m216h-3puq2dU44XGDI/edit?usp=drivesdk';
    const SpreadSheet = SpreadsheetApp.openByUrl(sheetUrl);

    // 定義 Google Sheet 中，各資料表的名稱。（名稱不對的話，會無法使用該資料表。）
    const signList = SpreadSheet.getSheetByName("sign");
    const safeReportList = SpreadSheet.getSheetByName("safeReport");
    const backReportList = SpreadSheet.getSheetByName("backReport");

    // 取得各個資料表中，最後一列的位置。（確保新增數據時，不會覆蓋到舊有資料。）
    var safeReportListCurrentRow = safeReportList.getLastRow(); // 取得工作表最後一欄（ 直欄數 ）
    var signListCurrentRow = signList.getLastRow(); // 取得工作表最後一欄（ 直欄數 ）
    var backReportListCurrentRow = backReportList.getLastRow(); // 取得工作表最後一欄（ 直欄數 ）

    // 確保初始值大於 0，否則 getRange 語法會產生錯誤。（確保）
    safeReportList.getLastRow() > 0 ? safeReportListCurrentRow = safeReportList.getLastRow() : safeReportListCurrentRow = 1;
    signList.getLastRow() > 0 ? signListCurrentRow = signList.getLastRow() : signListCurrentRow = 1;
    backReportList.getLastRow() > 0 ? backReportListCurrentRow = backReportList.getLastRow() : backReportListCurrentRow = 1;

    // 回覆文字，
    // 將根據指令，判斷要回覆使用者什麼樣的訊息，
    // 在後期會加入 JSON，並提供給 Line API，供傳送訊息給使用者。
    var replyMessage = []; 

    // 定義 name，賦予 訓員 的名字。
    var name = get_user_name();

    // 取得帳號名稱，
    // 打 Line API，根據 userId 搜尋該用戶的 Line 名稱。
    function get_user_name() {
        // 判斷為群組成員還是單一使用者
        switch (eventType) {
            case "user":
                var nameUrl = "https://api.line.me/v2/bot/profile/" + userId;
                break;
            case "group":
                var groupId = msg.events[0].source.groupId;
                var nameUrl = "https://api.line.me/v2/bot/group/" + groupId + "/member/" + userId;
                break;
        }

        try {
            //  呼叫 LINE User Info API，以 user ID 取得該帳號的使用者名稱
            var response = UrlFetchApp.fetch(nameUrl, {
                "method": "GET",
                "headers": {
                    "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN,
                    "Content-Type": "application/json"
                },
            });
            var nameData = JSON.parse(response);
            var reportName = nameData.displayName;
        }
        catch {
            reportName = "not avaliable";
        }
        return String(reportName)
    }

    // 傳送訊息，
    // 打 Line API 將傳送訊息給使用者。
    function send_to_line() {
        var url = 'https://api.line.me/v2/bot/message/reply';
        UrlFetchApp.fetch(url, {
            'headers': {
                'Content-Type': 'application/json; charset=UTF-8',
                'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
            },
            'method': 'post',
            'payload': JSON.stringify({
                'replyToken': replyToken,
                'messages': replyMessage,
            }),
        });
    }

    // 取得訓員的電話，
    // 將從 Google Sheet 的 sign 資料表中，根據訓員名字搜索，以取得電話號碼，
    // 並將電話號碼做為回傳值。
    function get_phone_number(){

      // 取得搜索目標的索引值。
      var memberIndex = search_member(signList,signListCurrentRow);

      var phoneNumber = "0";

      // Condition:是否有回報過，若有：進行取代。
      memberIndex != -1 ? phoneNumber=signList.getRange(memberIndex+1, 2).getValue() : null ;

      return phoneNumber;
    }

    // 用戶註冊，
    // 將狀態傳至 Google Sheet 中的 sign 資料表，並將資料表重新排序，
    // 然後會打 Line API 傳訊息給使用者，提示已完成回報。
    function sign(){
      // 要回饋給使用者傳送的訊息
      replyMessage = [
        {
          "type": "text",
          "text": "訓員：" + get_user_name() + "\n" +
                  "電話：" + splitedUserMessage[1] + "\n"
        }
      ]

      signList.getRange(signListCurrentRow + 1, 1).setValue(get_user_name());
      signList.getRange(signListCurrentRow + 1, 2).setValue(splitedUserMessage[1]);

      signListCurrentRow = signList.getLastRow();

      signList.sort(1);

      send_to_line();
    }

    // 安全回報：在家，
    // 將狀態傳至 Google Sheet 中的 safeReport 資料表，並將資料表重新排序，
    // 然後會打 Line API 傳訊息給使用者，提示已完成回報。
    function at_home_report(){

      // 要回饋給使用者傳送的訊息
      replyMessage = [
        {
          "type": "text",
          "text": get_user_name() +"\n" +
                  "狀態：" + "在家\n" +
                  "電話：" + get_phone_number() + "\n"
        }
      ]

      safeReportList.getRange(safeReportListCurrentRow + 1, 1).setValue(get_user_name());
      safeReportList.getRange(safeReportListCurrentRow + 1, 2).setValue("在家");
      safeReportList.getRange(safeReportListCurrentRow + 1, 3).setValue(get_phone_number());
      safeReportListCurrentRow = safeReportList.getLastRow();

      send_to_line();

      safeReportList.sort(1);
    }

    // 安全回報：在外，
    // 將狀態傳至 Google Sheet 中的 safeReport 資料表，並將資料表重新排序，
    // 然後會打 Line API 傳訊息給使用者，提示已完成回報。
    function outside_report(){
      // 要回饋給使用者傳送的訊息
      replyMessage = [
        {
          "type": "text",
          "text": get_user_name() + "\n" +
                  "狀態：" + splitedUserMessage[1] + "，" + splitedUserMessage[2] + "，預計 " + splitedUserMessage[3] +" 返家。\n" +
                  "電話：" + get_phone_number() + "\n"
        }
      ]

      // 傳送資料給數據庫
      safeReportList.getRange(safeReportListCurrentRow + 1, 1).setValue(get_user_name());
      safeReportList.getRange(safeReportListCurrentRow + 1, 2).setValue("在外");
      safeReportList.getRange(safeReportListCurrentRow + 1, 3).setValue(splitedUserMessage[2]);
      safeReportList.getRange(safeReportListCurrentRow + 1, 4).setValue(splitedUserMessage[3]);
      safeReportList.getRange(safeReportListCurrentRow + 1, 5).setValue(get_phone_number());
      safeReportListCurrentRow = safeReportList.getLastRow();

      // 傳送訊息給使用者
      send_to_line();

      safeReportList.sort(1);
    }

    // 安全回報：非住家，
    // 將狀態傳至 Google Sheet 中的 safeReport 資料表，並將資料表重新排序，
    // 然後會打 Line API 傳訊息給使用者，提示已完成回報。
    function not_home_report(){
      // 要回饋給使用者傳送的訊息
      replyMessage = [
        {
          "type": "text",
          "text": get_user_name() + "\n" +
                  "狀態：非住家，" + splitedUserMessage[2] + "。\n" +
                  "電話：" + get_phone_number() + "\n"
        }
      ]

      // 傳送資料給數據庫
      safeReportList.getRange(safeReportListCurrentRow + 1, 1).setValue(get_user_name());
      safeReportList.getRange(safeReportListCurrentRow + 1, 2).setValue("非住家");
      safeReportList.getRange(safeReportListCurrentRow + 1, 3).setValue(splitedUserMessage[2]);
      safeReportList.getRange(safeReportListCurrentRow + 1, 4).setValue(get_phone_number());
      safeReportList.getLastRow() > 0 ? safeReportListCurrentRow = safeReportList.getLastRow() : safeReportListCurrentRow = 1;

      // 傳送訊息給使用者
      send_to_line();

      safeReportList.sort(1);
    }

    // 收假回報，
    // 將狀態傳至 Google Sheet 中的 backReport 資料表，並將資料表重新排序，
    // 然後會打 Line API 傳訊息給使用者，提示已完成回報。
    function back_report(){
      // 要回饋給使用者傳送的訊息
      replyMessage = [
        {
          "type": "text",
          "text": get_user_name() + "\n" +
                  "方式：" + splitedUserMessage[1] + "，" + splitedUserMessage[2] + " 返營。\n" +
                  "狀態：" + splitedUserMessage[3] + "，快篩為" + splitedUserMessage[4] + "。\n"
        }
      ]

      backReportList.getRange(backReportListCurrentRow + 1, 1).setValue(get_user_name());
      backReportList.getRange(backReportListCurrentRow + 1, 2).setValue(splitedUserMessage[1]);
      backReportList.getRange(backReportListCurrentRow + 1, 3).setValue(splitedUserMessage[2]);
      backReportList.getRange(backReportListCurrentRow + 1, 4).setValue(splitedUserMessage[3]);
      backReportList.getRange(backReportListCurrentRow + 1, 5).setValue(splitedUserMessage[4]);

      backReportListCurrentRow = backReportList.getLastRow();

      backReportList.sort(1);

      send_to_line();
    }

    // 退伍回報，
    // 搜尋 Google Sheet 的 sign 資料表，尋找是否有該成員，
    // 若有將進行刪除，並將資料表重新排序。
    function delete_member(){

      // 取得搜索目標的索引值。
      var memberIndex = search_member(signList,signListCurrentRow);

      // Condition:是否有回報過，若有：將該列的資料刪除。
      memberIndex != -1 ? delete_row(signList,memberIndex) : null ;

    }

    // 刪除內容，
    // 從 Google Sheet 資料表中，根據索引值，刪除欄位的內容。
    // 參數：資料表、索引值
    function delete_row(list,index){
      // 將該列後五欄的內容清除
      list.getRange(index+1, 1, 1, 5).clearContent();

      // 升冪排序 A-Z
      list.sort(1);
    }


    // 搜尋成員，
    // 從 Google Sheet 資料表中，根據名字搜尋成員。
    // 參數：資料表、資料表最後一列
    // 回傳：該成員在 Google Sheet 資料表中的索引值
    function search_member(list,currentRow){
      var allMembers = list.getRange(1, 1, currentRow, 1).getValues().flat();
      var memberIndex = allMembers.indexOf(name);
      return memberIndex;
    }
    
    if (typeof replyToken === 'undefined') {
        return;
    };

    if (splitedUserMessage[0] == "回報"){
      if(splitedUserMessage[1] == "在家"){

        // 根據名稱，在所有資料中，搜索目標的索引值。
        var memberIndex = search_member(safeReportList,safeReportListCurrentRow);

        // Condition:是否有回報過，若有：進行取代，若無：進行新增。
        if (memberIndex != -1) {
          // 刪除該列資料
          delete_row(safeReportList,memberIndex);

          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          at_home_report();
        }
        else {
          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          at_home_report();
        }
      }else if (splitedUserMessage[1] == "在外"){

        // 根據名稱，在所有資料中，搜索目標的索引值。
        var memberIndex = search_member(safeReportList,safeReportListCurrentRow);

        // Condition:是否有回報過，若有：進行取代，若無：進行新增。
        if (memberIndex != -1) {
          // 刪除該列資料
          delete_row(safeReportList,memberIndex);

          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          outside_report();
        }
        else {
          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          outside_report();
        }
      }else if (splitedUserMessage[1] == "非住家"){

        // 根據名稱，在所有資料中，搜索目標的索引值。
        var memberIndex = search_member(safeReportList,safeReportListCurrentRow);

        // Condition:是否有回報過，若有：進行取代，若無：進行新增。
        if (memberIndex != -1) {
          // 刪除該列資料
          delete_row(safeReportList,memberIndex);

          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          not_home_report();
        }
        else {

          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          not_home_report();
        }
      }
    }else if (splitedUserMessage[0] == "註冊"){

        // 根據名稱，在所有資料中，搜索目標的索引值。
        var memberIndex = search_member(signList,signListCurrentRow);

        // Condition:是否有回報過，若有：進行取代，若無：進行新增。
        if (memberIndex != -1) {
          // 更新最後一列的位置
          signListCurrentRow = memberIndex;

          // 傳送資料至數據庫
          sign();
        }
        else {
          // 更新最後一列的位置
          signListCurrentRow = signList.getLastRow();

          // 傳送資料至數據庫
          sign();
        }



    }else if (splitedUserMessage[0] == "收假"){

        // 根據名稱，在所有資料中，搜索目標的索引值。
        var memberIndex = search_member(backReportList,backReportListCurrentRow);

        // Condition:是否有回報過，若有：進行取代，若無：進行新增。
        if (memberIndex != -1) {
          // 刪除該列資料
          delete_row(backReportList,memberIndex);

          // 更新最後一列的位置
          backReportListCurrentRow = backReportList.getLastRow();

          // 傳送資料至數據庫
          back_report();
        }
        else {

          // 更新最後一列的位置
          backReportListCurrentRow = backReportList.getLastRow();

          // 傳送資料至數據庫
          back_report();
        }

    }else if (splitedUserMessage[0] == "名單"){

      if(splitedUserMessage[1] == "已安全回報"){

        var reportList = "【 已安全回報人員 】" + "\n" +
                         "————————————";

        var allMembers = safeReportList.getRange(1, 1, safeReportListCurrentRow, 5).getValues().flat();
        
        for (var x = 0; x <= allMembers.length-1;x+=5) {
          
          if(allMembers[x+1] == "在家"){

            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "狀態：" + allMembers[x+1] + "。\n" +
                         "電話：" + allMembers[x+2] + "\n";

          }else if(allMembers[x+1] == "在外"){

            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "狀態：" + allMembers[x+1] + "，" + allMembers[x+2] + "，" + "預計 " + allMembers[x+3] + "到家。\n" +
                         "電話：" + allMembers[x+4] + "\n";

          }else if(allMembers[x+1] == "非住家"){

            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "狀態：" + allMembers[x+1] + "，" + allMembers[x+2] + "。\n" +
                         "電話：" + allMembers[x+3] + "\n";

          }
          
        }

        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "未安全回報"){
        
        var signMembers = signList.getRange(1, 1, signListCurrentRow, 1).getValues().flat();

        var safeReportMembers = safeReportList.getRange(1, 1, safeReportListCurrentRow, 1).getValues().flat();

        var reportList = "【 未安全回報人員 】" + "\n" +
                         "————————————\n";

        var flag = true;

        for(var i in signMembers){
          var sign_member = signMembers[i];
          if(safeReportMembers.indexOf(sign_member) == -1){
            reportList = reportList + sign_member + "\n";
            flag = false;
          }
        }

        flag ? reportList = reportList + "無，全體人員皆已完成回報。\n" : reportList = reportList;
        
        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "已收假回報"){
        // 文字訊息的開頭
        var reportList = "【 已收假回報人員 】" + "\n" +
                         "————————————";

        var allMembers = backReportList.getRange(1, 1, backReportListCurrentRow, 5).getValues().flat();

        for (var x = 0; x <= allMembers.length-1;x+=5) {
          
          if(allMembers[x+1] == "專車"){

            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "方式：專車，" + allMembers[x+2] + " 返營。\n" +
                         "狀態：" + allMembers[x+3] + "，快篩為" + allMembers[x+4] + "。\n";

          }else if(allMembers[x+1] == "步行"){

            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "方式：步行，" + allMembers[x+2] + " 返營。\n" +
                         "狀態：" + allMembers[x+3] + "，快篩為" + allMembers[x+4] + "。\n";

          }
          
        }
        
        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "未收假回報"){
        
        var signMembers = signList.getRange(1, 1, signListCurrentRow, 1).getValues().flat();

        var backReportMembers = backReportList.getRange(1, 1, backReportListCurrentRow, 1).getValues().flat();

        // 文字訊息的開頭
        var reportList = "【 未收假回報人員 】" + "\n" +
                         "————————————\n";

        var flag = true;

        for(var i in signMembers){
          var sign_member = signMembers[i];
          if(backReportMembers.indexOf(sign_member) == -1){
            reportList = reportList + sign_member + "\n";
            flag = false;
          }
        }

        flag ? reportList = reportList + "無，全體人員皆已完成回報。\n" : reportList = reportList;
        
        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "步行"){

        // 文字訊息的開頭
        var reportList = "【 步行返營人員 】" + "\n" +
                         "————————————";
        // 取得 A1 至 最後一列 範圍內的所有資料。
        var allMembers = backReportList.getRange(1, 1, backReportListCurrentRow, 5).getValues().flat();

        // 透過迴圈，將名單資料組合成文字。
        for (var x = 0; x <= allMembers.length-1;x+=5) {
          
          if(allMembers[x+1] == "步行"){
            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "方式：步行，" + allMembers[x+2] + " 返營。\n" +
                         "狀態：" + allMembers[x+3] + "，快篩" + allMembers[x+4] + "。\n";

          }
          
        }
        
        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "專車"){

        // 文字訊息的開頭
        var reportList = "【 專車返營人員 】" + "\n" +
                         "————————————";

        // 取得 A1 至 最後一列 範圍內的所有資料。
        var allMembers = backReportList.getRange(1, 1, backReportListCurrentRow, 5).getValues().flat();

        // 透過迴圈，將名單資料組合成文字。
        for (var x = 0; x <= allMembers.length-1;x+=5) {
          if(allMembers[x+1] == "專車"){
            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "方式：專車，" + allMembers[x+2] + " 返營。\n" +
                         "狀態：" + allMembers[x+3] + "，快篩" + allMembers[x+4] + "。\n";
          }
        }
        
        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "陰性"){
        // 文字訊息的開頭
        var reportList = "【 快篩陰性人員 】" + "\n" +
                         "————————————";

        // 取得 A1 至 最後一列 範圍內的所有資料。
        var allMembers = backReportList.getRange(1, 1, backReportListCurrentRow, 5).getValues().flat();

        // 透過迴圈，將名單資料組合成文字。
        for (var x = 0; x <= allMembers.length-1;x+=5) {
          if(allMembers[x+4] == "陰性"){
            reportList = reportList + "\n" + 
                            allMembers[x] + "\n";
          }
        }
        
        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "陽性"){
        // 文字訊息的開頭
        var reportList = "【 快篩陽性人員 】" + "\n" +
                         "————————————";

        // 取得 A1 至 最後一列 範圍內的所有資料。
        var allMembers = backReportList.getRange(1, 1, backReportListCurrentRow, 5).getValues().flat();

        // 透過迴圈，將名單資料組合成文字。
        for (var x = 0; x <= allMembers.length-1;x+=5) {
          if(allMembers[x+4] == "陽性"){
            reportList = reportList + "\n" + 
                         allMembers[x] + "\n";
          }
        }
        
        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "在外"){
        // 文字訊息的開頭
        var reportList = "【 在外人員 】" + "\n" +
                             "————————————";

        // 取得 A1 至 最後一列 範圍內的所有資料。
        var allMembers = safeReportList.getRange(1, 1, safeReportListCurrentRow, 5).getValues().flat();

        // 透過迴圈，將名單資料組合成文字。
        for (var x = 0; x <= allMembers.length-1;x+=5) {
          if(allMembers[x+1] == "在外"){
            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "狀態：" + allMembers[x+1] + "，" + allMembers[x+2] + "，預計 " + allMembers[x+3] +" 返家。\n" +
                         "電話：" + allMembers[x+4] + "\n";
          }
        }

        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "在家"){
        
        // 文字訊息的開頭
        var reportList = "【 在家人員 】" + "\n" +
                         "————————————";

        // 取得 A1 至 最後一列 範圍內的所有資料。
        var allMembers = safeReportList.getRange(1, 1, safeReportListCurrentRow, 5).getValues().flat();

        // 透過迴圈，將名單資料組合成文字。
        for (var x = 0; x <= allMembers.length-1;x+=5) {
          if(allMembers[x+1] == "在家"){
            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "狀態：" + allMembers[x+1] + "。\n" +
                         "電話：" + allMembers[x+2] + "\n";
          }
        }

        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }else if(splitedUserMessage[1] == "非住家"){
        
        // 文字訊息的開頭
        var reportList = "【 非住家人員 】" + "\n" +
                         "————————————";

        // 取得 A1 至 最後一列 範圍內的所有資料。
        var allMembers = safeReportList.getRange(1, 1, safeReportListCurrentRow, 5).getValues().flat();

        // 透過迴圈，將名單資料組合成文字。
        for (var x = 0; x <= allMembers.length-1;x+=5) {
          if(allMembers[x+1] == "非住家"){
            reportList = reportList + "\n" + 
                         allMembers[x] + "\n" +
                         "狀態：非住家，" + allMembers[x+2] +"。\n" +
                         "電話：" + allMembers[x+3] + "\n";
          }
        }

        // 要傳送的文字
        replyMessage = [
          {
            "type": "text",
            "text": reportList
          }
        ]

        // 傳送訊息
        send_to_line();
      }
    }else if (splitedUserMessage[0] == "退伍"){
      if(splitedUserMessage[1] == "在家"){

        // 根據名稱，在所有資料中，搜索目標的索引值。
        var memberIndex = search_member(safeReportList,safeReportListCurrentRow);

        // Condition:是否有回報過，若有：進行取代，若無：進行新增。
        if (memberIndex != -1) {
          // 刪除該列資料
          delete_row(safeReportList,memberIndex);

          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          at_home_report();

          // 刪除成員
          delete_member();
        }
        else {
          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          at_home_report();

          // 刪除成員
          delete_member();
        }
      }else if (splitedUserMessage[1] == "在外"){

        // 根據名稱，在所有資料中，搜索目標的索引值。
        var memberIndex = search_member(safeReportList,safeReportListCurrentRow);

        // Condition:是否有回報過，若有：進行取代，若無：進行新增。
        if (memberIndex != -1) {
          // 刪除該列資料
          delete_row(safeReportList,memberIndex);

          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          outside_report();

          // 刪除成員
          delete_member();
        }
        else {
          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          at_home_report();

          // 刪除成員
          outside_report();
        }
      }else if (splitedUserMessage[1] == "非住家"){

        // 根據名稱，在所有資料中，搜索目標的索引值。
        var memberIndex = search_member(safeReportList,safeReportListCurrentRow);

        // Condition:是否有回報過，若有：進行取代，若無：進行新增。
        if (memberIndex != -1) {
          // 刪除該列資料
          delete_row(safeReportList,memberIndex);

          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          at_home_report();

          // 刪除成員
          not_home_report();
        }
        else {
          // 更新最後一列的位置
          safeReportListCurrentRow = safeReportList.getLastRow();

          // 傳送資料至 Google Sheet
          at_home_report();

          // 刪除成員
          not_home_report();
        }
      }
    }else if ( splitedUserMessage[0] == "清除"){
      if (splitedUserMessage[1] == "回報") {

        // 清空安全回報、收假回報的所有資料
        safeReportList.clear();
        backReportList.clear();

        // 要傳送的文字
        replyMessage = [{
                  "type": "text",
                  "text": "已清除所有回報資料"
        }]

        // 傳送訊息
        send_to_line();
      }else if (splitedUserMessage[1] == "註冊"){

        // 清空安全回報、收假回報的所有資料
        signList.clear();

        // 要傳送的文字
        replyMessage = [{
                  "type": "text",
                  "text": "已清除註冊名單"
        }]

        // 傳送訊息
        send_to_line();
      }
    }else if (userMessage == "指令集" | userMessage == "指令" | userMessage == "help") {

      // 要傳送給使用者傳送的訊息
      replyMessage = [
        {
          "type": "text",
          "text": "【 指令集 】\n" +
                  "————————————\n" +
                  "※ 使用時無需輸入方框，方框：[ ] 為填入文字的意思。 \n" +
                  "————————————\n" +
                  "使用說明：\n" +
                  "１）　使用說明\n" +
                  "————————————\n" +
                  "註冊電話：\n" +
                  "１）　註冊/[訊員的電話]\n" +
                  "————————————\n" +
                  "安全回報：\n" +
                  "１）　回報/在家\n" +
                  "２）　回報/在外/[做甚麼]/[預計返家時間]\n" +
                  "３）　回報/非住家/[地點]\n" +
                  "————————————\n" +
                  "收假回報：\n" +
                  "１）　收假/[返營方式]/[返營時間]/[目前位置]/[快篩結果]\n" +
                  "————————————\n" +
                  "查看名單：\n" +
                  "１）　名單/已安全回報\n" +
                  "２）　名單/未安全回報\n" + 
                  "３）　名單/已收假回報\n" +
                  "４）　名單/未收假回報\n" +
                  "５）　名單/步行\n" +
                  "６）　名單/專車\n" +
                  "７）　名單/陰性\n" +
                  "８）　名單/陽性\n" +
                  "９）　名單/在外\n" +
                  "１０）名單/在家\n" +
                  "１１）名單/非住家\n" +
                  "————————————\n" +
                  "清除資料：\n" +
                  "１）　清除/回報\n" +
                  "２）　清除/註冊\n" +
                  "————————————\n" +
                  "退伍回報：\n" +
                  "１）　退伍/在家\n" +
                  "２）　退伍/在外/[做甚麼]/[預計返家時間]\n" +
                  "３）　退伍/非住家/[地點]\n"
        }
      ]

      // 傳送訊息給使用者
      send_to_line();
  
    }else if (userMessage == "使用說明" ) {

      // 要傳送的文字
      replyMessage = [
        {
          "type": "text",
          "text": "【 使用說明 】\n" +
                  "在註冊後，即可進行回報。\n" +
                  "１）　註冊指令：註冊/[使用者電話]\n" +
                  "使用時無需輸入方框，方框：[ ] 為填入文字的意思。"
        }
      ]

      // 傳送訊息
      send_to_line();
  
    }

    // 其他非關鍵字的訊息則不回應（ 避免干擾群組聊天室 ）
    else {
        console.log("else here,nothing will happen.")
    }
}