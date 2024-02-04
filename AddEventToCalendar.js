function addEventsToCalendar() {
  // スプレッドシートのデータを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('予定表'); // シート名を指定
  var data = sheet.getDataRange().getValues();

  // カレンダーへのアクセス
  var calendar = CalendarApp.getDefaultCalendar(); // デフォルトのカレンダーを使用

  // データを処理してカレンダーに追加
  for (var i = 1; i < data.length; i++) { // ヘッダー行をスキップ
    var date = data[i][0]; // 日付列 (1行目)
    var timeRange = data[i][1]; // 時間範囲列 (2行目)
    var eventTitle = data[i][2]; // イベント名列 (3行目)
    var location = data[i][3]; // 場所列 (4行目)

    // 時間範囲を解析
    var timeRangeArray = timeRange.split('~');
    var startTimeString = timeRangeArray[0].trim();
    var endTimeString = timeRangeArray[1].trim();
    
    // 開始時間を変換
    var startTimeArray = startTimeString.split(':');
    var startHours = parseInt(startTimeArray[0]);
    var startMinutes = parseInt(startTimeArray[1]);
    
    var startDateObject = new Date(date);
    startDateObject.setHours(startHours, startMinutes, 0);

    // 終了時間を変換
    var endTimeArray = endTimeString.split(':');
    var endHours = parseInt(endTimeArray[0]);
    var endMinutes = parseInt(endTimeArray[1]);
    
    var endDateObject = new Date(date);
    endDateObject.setHours(endHours, endMinutes, 0);


    // 既存のイベントを取得して重複をチェック
    var existingEvents = calendar.getEvents(startDateObject, endDateObject);
    var isDuplicate = false;

    for (var j = 0; j < existingEvents.length; j++) {
      if (existingEvents[j].getTitle() === eventTitle) {
        isDuplicate = true;
        break;
      }
    }

    // 重複がなければイベントを追加
    if (!isDuplicate) {
      var event = calendar.createEvent(eventTitle, startDateObject, endDateObject);
      event.setLocation(location);
    }
  }
}


