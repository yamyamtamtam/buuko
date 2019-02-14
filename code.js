// Channel Access Token
var access_token = "xxxxxx";
// スプレットシート呼び出し
var spreadsheet = SpreadsheetApp.openById("xxxxxx");
var spreadsheet_month = spreadsheet.getSheetByName("month");
var spreadsheet_days = spreadsheet.getSheetByName("days");
var spreadsheet_users = spreadsheet.getSheetByName("users");
var to = [];

/**
  * pushメッセージ送信
 */
function push(text){
  pushUsersGet();
  var url = "https://api.line.me/v2/bot/message/multicast";
  var headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + access_token,
  };
  var postData = {
    "to" : to,
      "messages" : [
        {
        'type':'text',
        'text':text,
      }
    ]
  };
  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };
  return UrlFetchApp.fetch(url, options); 
}

/**
 * 自動返信
 */
function reply(data,text) {
  var url = "https://api.line.me/v2/bot/message/reply";
  var headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + access_token,
  };
  var postData = {
    "replyToken" : data.events[0].replyToken,
    "messages" : [
      {
        'type':'text',
        'text':text,
      }
    ]
  };
  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };
  return UrlFetchApp.fetch(url, options);
}

/**
 * スプレットシート書き込み・読み込み関係
 */
//スプレットシートからユーザーの一覧取得
function pushUsersGet(){
  var sheet_users = spreadsheet_users.getDataRange().getValues();
  var users = sheet_users.filter( //重複を弾く
    function (x, i, self) {
      return self.indexOf(x) === i;
    }
  );
  for(var i = 0; i < users.length; i++){
    if(users[i][0] !== ""){
      to.push(users[i][0]);
    }
  }
}
//シートの3行目までデータ保持、それ以降は2行目に書き込み
//常に最新は3行目に書き込まれて2行目はひとつ前のデータとして保持される
function sheetLimitTwo(sheet_obj,value){
  //月や日が3行目と違ったら、今の3行目のデータを2行目に写し、3行目にデータを書き込み
  if(sheet_obj.getRange(2, 1).getValue() && sheet_obj.getRange(3, 1).getValue() !== value){
    var last_value = sheet_obj.getRange(3, 1).getValue();
    sheet_obj.getRange(2, 1).setValue(last_value);
    sheet_obj.getRange(3, 1).setValue(value);    
  }
  else if(!sheet_obj.getRange(2, 1).getValue() && !sheet_obj.getRange(3, 1).getValue()){
    sheet_obj.getRange(2, 1).setValue(value);
    sheet_obj.getRange(3, 1).setValue(value);
  }
}

/**
  * 月変更時にスプレットシートに書き込み、pushメッセージ送信
  * 今のカレンダー上の月と、その次月のどちらかがセットされる
  * 実行するとmonthシートは古い月(2)と今月(3)、今月がふたつ(monthsReturn実行後)、今月(2)と来月(3)のいずれかの状態になる
 */
function monthsExchange(json){
  var text= '';  
  var sheet_first = "";
  var sheet_second = "";
  var month_next = new Date();
  month_next.setMonth(month_next.getMonth() + 1);
  var month = Utilities.formatDate(new Date(), "JST", "yyyy/M");
  var month_next = Utilities.formatDate(month_next, "JST", "yyyy/M");
  if(spreadsheet_month.getRange(2,1).getValue()){
    var sheet_first = Utilities.formatDate(spreadsheet_month.getRange(2,1).getValue(), "JST", "yyyy/M");
  }
  if(spreadsheet_month.getRange(3,1).getValue()){  
    var sheet_second = Utilities.formatDate(spreadsheet_month.getRange(3,1).getValue(), "JST", "yyyy/M");
  }
  if(sheet_first !== month && sheet_second !== month){ //今月のシート作成
    spreadsheet.insertSheet(month);
    spreadsheet.getSheetByName(month).getRange(1,1).setValue("給料");
    spreadsheet.getSheetByName(month).getRange(1,2).setValue(0);
    sheetLimitTwo(spreadsheet_month,month);
    text = month + 'からの会計を管理するぶひよ。\n\n例「公共(半角スペース)電気(半角スペース)1000」\nのフォーマットで公共料金を入れて欲しいぶひ。\n\n※最初は「公共」にして欲しいぶひ';
    push(text);
  }
  else if(sheet_second !== month_next){ //来月のシート作成
    if(!spreadsheet.getSheetByName(month_next)){
      spreadsheet.insertSheet(month_next);
      spreadsheet.getSheetByName(month_next).getRange(1,1).setValue("給料");
      spreadsheet.getSheetByName(month_next).getRange(1,2).setValue(0);
    }
    sheetLimitTwo(spreadsheet_month,month_next);
    text = month_next + 'からの会計を管理するぶひよ。\n\n例「公共(半角スペース)電気(半角スペース)1000」\nのフォーマットで公共料金を入れて欲しいぶひ。\n\n※最初は「公共」にして欲しいぶひ';
    push(text);
  }
  else{
    text = '月の変更ができないぶひ。次月までしか設定できないぶひよ。前の月に戻る場合は「月 前」で戻れるぶひ';
    reply(json,text);
  }
  //初回のみ日付のシートを作成
  if(spreadsheet_days.getRange(2, 1).getValue() === "" && spreadsheet_days.getRange(3, 1).getValue() === ""){
    var today = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
    today = "【" + month + "】" + today;
    spreadsheet_days.getRange(2, 1).setValue(today);
    spreadsheet_days.getRange(3, 1).setValue(today);
    if(!spreadsheet.getSheetByName(today)){
      spreadsheet.insertSheet(today);
      spreadsheet.getSheetByName(today).getRange(1,1).setValue(today);
      spreadsheet.getSheetByName(today).getRange(2,2).setValue(0);
    }
  }
}

/**
  * 月ひとつ前に戻す時にスプレットシートに書き込み、pushメッセージ送信
 */
function monthsReturn(json){
  var text= '';
  var real_month = new Date();
  var current_month = spreadsheet_month.getRange(3, 1).getValue();
  if(current_month.getTime() > real_month.getTime()){
    var month = Utilities.formatDate(new Date(), "JST", "yyyy/M");
    spreadsheet_month.getRange(3, 1).setValue(month);
    text = '月の設定を' + month + 'に戻したぶひ。また次の月に行きたいときは「月変更」「月」「つき」でやるぶひよ。';
    push(text);
  }
  else{
    text = 'いまの月より前には戻せないぶひ。ぶー子は過去を振り返らない豚ぶひ。';
    reply(json,text);
  }
}

/**
  * 日付変更時にスプレットシートに書き込み、pushメッセージ送信
 */
function daysExchange(){
  var month_obj = spreadsheet_month.getRange(spreadsheet_month.getLastRow(),1).getValue();
  var month = month_obj.getFullYear() + "/" + (month_obj.getMonth() +1);
  var today = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  today = "【" + month + "】" + today;
  sheetLimitTwo(spreadsheet_days,today);
  if(!spreadsheet.getSheetByName(today)){
    spreadsheet.insertSheet(today);
    spreadsheet.getSheetByName(today).getRange(1,1).setValue(today);
    spreadsheet.getSheetByName(today).getRange(2,2).setValue(0);
  }
  var today_message = today.split("】");
  var text = today_message[1] + 'からの会計を管理するぶひよ。\n\n例「ラルズ(半角スペース)2655」\nのフォーマットでかかったお金を入れて欲しいぶひ。';
  push(text);
}

/**
 * 計算関係
 */
//月の給料から公共料金を引いた合計
function monthSum(month){
  var values = spreadsheet.getSheetByName(month).getDataRange().getValues();
  var sum = 0;
  if(values[0][2] !==""){
    var sum = values[0][1];
  }
  for(var i = 1; i < values.length; i++){
    sum = sum - values[i][1];
  }
  return sum;
}
//ある日付から現在まで使った金額合計
function daysSum(day){
  var values = spreadsheet.getSheetByName(day).getDataRange().getValues();
  var sum = 0;
  for(var i = 0; i < values.length; i++){
    if(!isNaN(values[i][1])){
      sum = sum + Number(values[i][1]);
    }
  }
  return sum;
}
//ある日付のうち特定のものの合計金額を合算する
function thingSum(day,thing){
  var thingsum = 0;
  var values = spreadsheet.getSheetByName(day).getDataRange().getValues();
  for(var i = 0; i < values.length; i++){
    if(values[i][0] === thing){
      thingsum = thingsum + Number(values[i][1]);
    }
  }
  return thingsum;
}
//シート一覧から同じ月のシートの合計金額を合算する（今月使った金額の合計を返す）
function daysSumAll(month){
  var dayssumall = 0;
  var sheets = spreadsheet.getSheets();
  var sheets_month = new RegExp("【" + month + "】" + ".*");
  for(var i = 0; i < sheets.length; i++){
    var sheets_name = sheets[i].getName();
    if(sheets_month.test(sheets_name)){
      dayssumall = dayssumall + daysSum(sheets_name);
    }  
  }
  return dayssumall;
}

/**
  * 公共料金を月のスプレットシートに書き込み、pushメッセージ送信
 */
function publicAdd(json,data){
  var month_obj = spreadsheet_month.getRange(3,1).getValue();
  var month = month_obj.getFullYear() + "/" + (month_obj.getMonth() +1);
  var prices = data.split(" ");
  var text = "";
  if(!isNaN(prices[2]) && prices[2] !== ""){ //数値か判定
    spreadsheet.getSheetByName(month).getRange(spreadsheet.getSheetByName(month).getLastRow()+1, 1).setValue(prices[1]);
    spreadsheet.getSheetByName(month).getRange(spreadsheet.getSheetByName(month).getLastRow(), 2).setValue(prices[2]);
    var sum = monthSum(month);
    text = prices[1] + 'に' + prices[2] + '円かかったぶひ。\n\n今月使えるお金は、' + sum + '円ぶひ。\n\n1週間の目安は、' + sum/4 + '円ぶひよ。';
    push(text);
  }
  else{
    text = '数字を入力してほしいぶひ。\n\n例「公共(半角スペース)電気(半角スペース)1000」';
    reply(json,text);
  }

}

/**
  * 給料を月のスプレットシートに書き込み、pushメッセージ送信
 */
function salaryAdd(json,data){
  var month_obj = spreadsheet_month.getRange(3,1).getValue();
  var month = month_obj.getFullYear() + "/" + (month_obj.getMonth() +1);
  var prices = data.split(" ");
  var text = "";
  if(!isNaN(prices[1]) && prices[1] !== ""){ //数値か判定
    spreadsheet.getSheetByName(month).getRange(1, 1).setValue(prices[0]);
    spreadsheet.getSheetByName(month).getRange(1, 2).setValue(prices[1]);
    text = 'お給料は' + prices[1] + '円ぶひ。';
    push(text);
  }
  else{
    text = '数字を入力してほしいぶひ。\n\n例「給料(半角スペース)200000」';
    reply(json,text);
  }

}

/**
  * ある日付から使った金額を日付のスプレットシートに書き込み、pushメッセージ送信
 */
function daysAdd(json,data){
  var month = Utilities.formatDate(spreadsheet_month.getRange(3,1).getValue(), "JST", "yyyy/M");
  var days = spreadsheet_days.getRange(3,1).getValue();
  var days_message = days.split("】");
  var prices = data.split(" ");
  var text = "";
  if(!isNaN(prices[1]) && prices[1] !== ""){ //数値か判定
    spreadsheet.getSheetByName(days).getRange(spreadsheet.getSheetByName(days).getLastRow()+1, 1).setValue(prices[0]);
    spreadsheet.getSheetByName(days).getRange(spreadsheet.getSheetByName(days).getLastRow(), 2).setValue(prices[1]);
    var monthsum = monthSum(month) - daysSumAll(month);
    var dayssum = daysSum(days);
    text = prices[0] + 'に' + prices[1] + '円かかったぶひ。\n\n' + days_message[1] + 'から合計' + dayssum + '円使ったぶひ。\n\n今月は残り' + monthsum + '円使えるぶひぶひ。';
    push(text);
  }
  else{
    text = '数字を入力してほしいぶひ。\n\n例「ラルズ(半角スペース)1000」';
    reply(json,text);
  }
}

/**
  * 今月の会計を一括して、reply送信
 */
function paymentAll(json){
  var month_obj = spreadsheet_month.getRange(3,1).getValue();
  var month = month_obj.getFullYear() + "/" + (month_obj.getMonth() +1);
  var last_days = spreadsheet_days.getRange(3,1).getValue();
  var last_days_message = last_days.split("】");
  var text = '今月の公共料金をのぞいて使える金額は' + monthSum(month) + '円ぶひ。\n\n' + last_days_message[1] + 'から' + daysSum(last_days) + '円使ったぶひ。\n\n今月の残り使える金額は' + ( monthSum(month) -  daysSumAll(month)) + '円ぶひ～';
  reply(json,text);
}

/**
  * 今月の会計のうち、使った内容を指定して合計、reply送信
 */
function thingTotal(json,data){
  var thing = data.split(" ");
  var month_obj = spreadsheet_month.getRange(3,1).getValue();
  var month = month_obj.getFullYear() + "/" + (month_obj.getMonth() +1);
  var sum = 0;
  //シート一覧取得
  var sheets = spreadsheet.getSheets();
  var sheets_month = new RegExp("【" + month + "】" + ".*");
  for(var i = 0; i < sheets.length; i++){
    var sheets_name = sheets[i].getName();
    if(sheets_month.test(sheets_name)){
      sum = sum + thingSum(sheets_name,thing[1]);
    }  
  }
  var text = '今月、' + thing[1] + 'に使った金額は' + sum + '円ぶひ。';
  reply(json,text);
}

/**
  * 使用説明
 */
function Instructions(json){
  var text = '';
  if(!spreadsheet_month.getRange(2,1).getValue()){
    text = 'いちばんはじめは「月変更」「月」「つき」で月を設定してほしいぶひ。これをしないとぶー子はあまり返事できないぶひ～\n\n\n';
  }
  text = text +
  '「月変更」「月」「つき」で月の切り替えをできるぶひ。次月までは行けるぶひ。給料日とかにやるぶひ～\n\n' +
  '月の切り替えを間違えたときは「月 前」「つき まえ」で戻れるぶひ。でも今月より前には戻れないぶひ～\n\n' +
  '「日付変更」「日付」「ひづけ」で基準の日付を設定できるぶひ。いつからいつまでで集計したいときにやるぶひ～\n' +
  '日付変更機能はちょっとわかりづらいぶひ。「日付 説明」で詳しい使い方をおしえるぶひ～\n\n' +
  '「会計」「かいけい」で色々と一気にみれるぶひ。便利ぶひね～\n\n' +
  '例「合計(半角スペース)外食」のフォーマットで今月なにかに使った合計の金額がわかるぶひ。ぶひぶひ～\n\n' +
  '例「給料(半角スペース)200000」のフォーマットで給料を登録するぶひ。いっぱいもらえると嬉しいぶひ～\n\n' +
  '例「公共(半角スペース)電気(半角スペース)1000」のフォーマットで公共料金を登録するぶひ。冬は高くなるぶひ～\n\n' +
  '例「食費(半角スペース)1000」のフォーマットで使ったお金を登録するぶひ。ついついいっぱい食べちゃうぶひ～\n\n' +
  '登録を間違えたときは、「食費(半角スペース)-1000」のようにマイナスを入れればいいぶひよ。\n\n' +
  '食べてすぐ寝ると豚になるぶひ～罪ぶひギルティぶひ～';
  reply(json,text);
}
/**
  * 日付説明
 */
function daysInstructions(json){
  var text = '';
  text = text +
  '「日付変更」「日付」「ひづけ」で基準の日付を設定できるぶひ。\n\n' +
  '例えば、1週間に10000円使うことを目安に家計を管理しているとするぶひ。そうすると、1週間経った日にその1週間でそれだけ使ったか知りたくなるぶひ。\n\n' +
  '具体的には、4/1に「ひづけ」といってくれれば、4/1からの出費を数えるぶひ。そうすると4/1からの出費だけを毎回教えるぶひよ。\n' +
  'そのあと、4/7に次の1週間の出費を数えたくなったら、4/7に「ひづけ」といってくれれば今度は4/1からの出費を忘れて、4/7からの出費を数えるぶひ。\n\n' +
  'これは、月にいくら使ったかの全体の出費とは別に数えているぶひ。なので使わなくても家計管理はできるぶひよ。あくまで便利機能のひとつぶひ。\n\n' +
  '一例のイメージはこんな感じぶひ。\n\n' +
  'やりやすいようにぶー子とお話してくれれば嬉しいぶひ～。';
  reply(json,text);
}
 
/**
 * postされたときの処理
 */
function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  console.log(json.events[0]);
  var type = json.events[0].type;
  if(type === "message"){ //メッセージが送られてきたとき
    var message = json.events[0].message.text;
    var salary = /給料 .*/;
    var public_cost = /公共 .* .*/;
    var days_cost = /.* .*/;
    var days_cost_mistake = /公共 .*/;
    var total = /合計 .*/;
    if(public_cost.test(message)){ //公共料金追加。push
      publicAdd(json,message);
    }
    else if(salary.test(message)){ //給料追加。push
      salaryAdd(json,message);
    }
    else if(total.test(message)){ //今月何にいくら使ったか。reply
      thingTotal(json,message);
    }
    else if(message === "月 前" ||  message === "つき まえ"){ //月戻す。push
      monthsReturn(json);
    }
    else if(message === "日付 説明" ||  message === "ひづけ せつめい"){ //日付の説明。reply
      daysInstructions(json);
    }
    else if(days_cost.test(message) && !days_cost_mistake.test(message)){ //支出追加。push
      daysAdd(json,message);
    }
    else if(message === "会計" || message === "かいけい"){ //合計。reply
      paymentAll(json);
    }
    else if(message === "説明" ||　message === "せつめい"){ //説明書。replay
      Instructions(json);
    }
    else if(message === "日付変更" || message === "日付" || message === "ひづけ"){ //基準日付変更。push
      daysExchange();
    }
    else if(message === "月変更" || message === "月" || message === "つき"){ //月変更。push
      monthsExchange(json);
    }
    else {
      var error = 'なんか間違ってるぶひ。「説明」で使い方がわかるぶひ。\n\n豚もおだてれば木に登るけど木が折れるぶひ～';
      reply(json,error);
    }
  }
  else if(type === "follow"){ //友達追加時にuser_idをスプレットシートに書き込み
    var current_user_id = json.events[0].source.userId;
    var result = "";
    pushUsersGet();
    for(var i = 0; i < to.length; i++){
      if(to[i].indexOf(current_user_id) >= 0){
        result = "exist";
      }
    }
    if(result != "exist"){
      var data = spreadsheet_users.getRange(spreadsheet_users.getLastRow()+1, 1).setValue(current_user_id);
    }
    var thanks = '友達登録ありがとうぶひ！「説明」でぶー子の使い方をおしえるぶひ。';
    reply(json,thanks);
  }
}

