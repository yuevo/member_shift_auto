let id = 'GoogleカレンダーIDを入れる';
let calendar = CalendarApp.getCalendarById(id);

let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シフト');
let set_date = sheet.getRange(6, 1).getValue();
let last_date = new Date(set_date.getFullYear(), set_date.getMonth()+1, 0);
let last_day = last_date.getDate()
let last_row = sheet.getRange(7, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
let all_member_shifts = sheet.getSheetValues(7, 1, last_row - 6, last_day);
let all_comments = sheet.getRange("A7:AF35"); 
let all_comments_shifts = all_comments.getNotes()

let today = new Date(set_date.getFullYear(), set_date.getMonth(), 1);
let year = today.getFullYear();
let month = 1 + today.getMonth();
var null_count = [];

function firstCheck() {
  all_member_shifts.forEach( function(member_shifts) {
    var result = member_shifts.filter(null_shift => null_shift == "");
    if (result.length) {
      null_count++;
      nullCheck_(member_shifts)
    }
  })
  if (null_count == "") {
    lastCheck_(all_member_shifts)
  }　else {
    Browser.msgBox(null_count + "件、空のシフトがあったので作業を中断しました。シフト表を修正してから再度実行お願いします。");
  }
}

function nullCheck_(member_shifts) {
  member_shifts.forEach( function(shift, index) {
    if (shift == "") {
      var name = member_shifts[0];
      Browser.msgBox(name + "さんの未入力のシフトがあります。" + month + "/" + (index) + "のシフトに値を入力しましょう。");
    }
  });
}

function lastCheck_(all_member_shifts) {
  var check = Browser.msgBox("【" + month + "月】の見学バディ用Googleカレンダーへ自動入力を行います。", "続行しますか？", Browser.Buttons.OK_CANCEL);
  if (check == 'ok') {
    autoShift_(all_member_shifts);
    Browser.msgBox("完了しました。");
  }
  if (check == 'cancel') {
    Browser.msgBox("処理はキャンセルされました。");
  }
}

function autoShift_(all_member_shifts) {
  all_member_shifts.forEach( function(shifts, i) {
    var name = all_member_shifts[i][0];
    shifts.forEach( function(shift, i) {
      judgeCreate_(shift, i, name);
    })
    autoCommentShift_(name, i)
  })
}

function autoCommentShift_(name, i) {
  if (all_comments_shifts[i] != "undefined") {
    all_comments_shifts[i].forEach( function(comment, i) {
      if (comment != "") {
        judgeCreate_(comment, i, name)
      } 
    });
  }
}

function judgeCreate_(shift, index, name) {
  Utilities.sleep(1000);
  if (shift == "基10-19") {
    kiso10_19_(index, name);
  } else if (shift == "基11-20") {
    kiso11_20_(index, name);
  } else if (shift == "基中11-22") {
    kiso11_22_(index, name);
  } else if (shift == "基14-19") {
    kiso14_19_(index, name);
  } else if (shift == "基14-22") {
    kiso14_22_(index, name);
  } else if (shift == "基10-13") {
    kiso10_13_(index, name);
  } else if (shift == "応10-19") {
    ouyo10_19_(index, name);
  } else if (shift == "応11-20") {
    ouyo11_20_(index, name);
  } else if (shift == "応中11-22") {
    ouyo11_22_(index, name);
  } else if (shift == "応14-22") {
    ouyo14_22_(index, name);
  } else if (shift == "応19-22") {
    ouyo19_22_(index, name);
  } else if (shift == "応10-13") {
    ouyo10_13_(index, name);
  } else if (shift == "応11-13") {
    ouyo11_13_(index, name);
  } else if (shift == "終10-19" || shift == "終中10-19") {
    saishu10_19_(index, name);
  } else if (shift == "終11-20") {
    saishu11_20_(index, name);
  } else if (shift == "終中11-22") {
    saishu11_22_(index, name);
  } else if (shift == "終14-19" || shift == "終中14-19") {
    saishu14_19_(index, name);
  } else if (shift == "終14-22") {
    saishu14_22_(index, name);
  } else if (shift == "終10-13") {
    saishu10_13_(index, name);
  } else if (shift == "終10-16") {
    saishu10_16_(index, name);
  } else if (shift == "終19-22") {
    saishu19_22_(index, name);
  } else if (shift == "C10-13" ) {
    chat10_13_(index, name);
  } else if (shift == "C10-19") {
    chat10_19_(index, name);
  } else if (shift == "C19-22") {
    chat19_22_(index, name);
  } else if (shift == "C14-22") {
    chat14_22_(index, name);
  } else if (shift == "C18-22") {
    chat18_22_(index, name);
  } else if (shift == "R10-13") {
    review10_13_(index, name);
  } else if (shift == "R10-19" || shift == "R中10-19") {
    review10_19_(index, name);
  } else if (shift == "R14-19") {
    review14_19_(index, name);
  } else if (shift == "R18-22") {
    review18_22_(index, name);
  } else if (shift == "R19-22") {
    review19_22_(index, name);
  } else if (shift == "R14-22" || shift == "R中14-22") {
    review14_22_(index, name);
  } else if ( shift == "外A" ||
              shift == "外B" ||
              shift == "外C" ||
              shift == "F10-13" || 
              shift == "F10-19" || 
              shift == "F14-19" || 
              shift == "F14-22" ||
              shift == "公" || 
              shift == "有" || 
              shift == "年" || 
              shift == "C&総" ||
              shift == "総会") {
    return
  } else {
    return
  }
}

function kiso10_19_(i, name) {
  calendar.createEvent(name + '（基礎通話）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function kiso11_20_(i, name) {
  calendar.createEvent(name + '（基礎通話）', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/20:00'));
}

function kiso11_22_(i, name) {
  calendar.createEvent(name + '（基礎通話）', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
  calendar.createEvent(name + '（基礎通話）', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function kiso14_19_(i, name) {
  calendar.createEvent(name + '（基礎通話）', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function kiso14_22_(i, name) {
  calendar.createEvent(name + '（基礎通話）', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function kiso10_13_(i, name) {
  calendar.createEvent(name + '（基礎通話）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function ouyo10_19_(i, name) {
  calendar.createEvent(name + '（応用通話）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function ouyo11_20_(i, name) {
  calendar.createEvent(name + '（応用通話）', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/20:00'));
}

function ouyo11_22_(i, name) {
  calendar.createEvent(name + '（応用通話）', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
  calendar.createEvent(name + '（応用通話）', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function ouyo14_22_(i, name) {
  calendar.createEvent(name + '（応用通話）', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function ouyo19_22_(i, name) {
  calendar.createEvent(name + '（応用通話）', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function ouyo10_13_(i, name) {
  calendar.createEvent(name + '（応用通話）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function ouyo11_13_(i, name) {
  calendar.createEvent(name + '（応用通話）', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function saishu10_19_(i, name) {
  calendar.createEvent(name + '（最終通話）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function saishu11_20_(i, name) {
  calendar.createEvent(name + '（最終通話）', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/20:00'));
}

function saishu11_22_(i, name) {
  calendar.createEvent(name + '（最終通話）', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
  calendar.createEvent(name + '（最終通話）', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function saishu14_22_(i, name) {
  calendar.createEvent(name + '（最終通話）', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function saishu14_19_(i, name) {
  calendar.createEvent(name + '（最終通話）', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function saishu10_13_(i, name) {
  calendar.createEvent(name + '（最終通話）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function saishu10_16_(i, name) {
  calendar.createEvent(name + '（最終通話）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/16:00'));
}

function saishu19_22_(i, name) {
  calendar.createEvent(name + '（最終通話）', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function chat10_13_(i, name) {
  calendar.createEvent(name + '（チャット）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function chat10_19_(i, name) {
  calendar.createEvent(name + '（チャット）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function chat19_22_(i, name) {
  calendar.createEvent(name + '（チャット）', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
} 

function chat14_22_(i, name) {
  calendar.createEvent(name + '（チャット）', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function chat18_22_(i, name) {
  calendar.createEvent(name + '（チャット）', new Date(year + '/' + month + '/' + i + '/18:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function review10_13_(i, name) { 
  calendar.createEvent(name + '（レビュー）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function review10_19_(i, name) {
  calendar.createEvent(name + '（レビュー）', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function review14_19_(i, name) {
  calendar.createEvent(name + '（レビュー）', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function review18_22_(i, name) {
  calendar.createEvent(name + '（レビュー）', new Date(year + '/' + month + '/' + i + '/18:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function review19_22_(i, name) {
  calendar.createEvent(name + '（レビュー）', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));    
}

function review14_22_(i, name) {
  calendar.createEvent(name + '（レビュー）', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}
