function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('秘密情報')                      //メニュー名
    .addItem('PW取得', 'PwAcquisitionDecision') //PW取得ロボ 起動
    .addToUi();
}


/**
 * PW取得ロボ
 * リクエスト内容が正しいか、権限があるかを判定し、
 * 権限があればPWを受け渡すロボからPWを受け取る
 */
function PwAcquisitionDecision() {
  //ヘッダー名を変えたらここを変更する
  //※権限管理シート、リクエストシート、PW管理SSのヘッダー名は揃える必要がある
  const company_name = "会社名";
  const site_name = "サイト・システム名";
  const usage_name = "用途";
  const account_name = "閲覧権限アカウント";
  const valid_name = "有効";

  //環境設定
  const date = Utilities.formatDate(new Date(),"Asia/Tokyo","yyyy/MM/dd HH:mm:ss");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const user_account = Session.getActiveUser().getEmail(); //アカウント

  const authority_sheet = ss.getSheetByName("A-1：権限管理"); //権限管理シート
  const log_sheet = ss.getSheetByName("A-2：リクエストログ"); //リクエストログシート
  const request_sheet = ss.getSheetByName("B-1：リクエスト"); //リクエストシート


  //リクエスト内容を取得
  const request_array = request_sheet.getRange("A2:B4").getDisplayValues();
  const request_obj = Object.fromEntries(request_array); //使いやすいようにオブジェクトにする

  //権限管理シートの値を検索
  const authority_data = authority_sheet.getRange("A1:F").getDisplayValues(); //権限管理シートの値を取得
  const authority_header = authority_data.shift(); //ヘッダー取得
  
  //各項目の列番号
  const company_index = authority_header.indexOf(company_name);
  const site_index = authority_header.indexOf(site_name);
  const usage_index = authority_header.indexOf(usage_name);
  const account_index = authority_header.indexOf(account_name);
  const valid_index = authority_header.indexOf(valid_name);


  //リクエストログシート更新用の設定
  const log_text = [date,user_account,request_obj[company_name],request_obj[site_name],request_obj[usage_name]];

  //稼働終了時の関数
  const finalize = (status, message) => {
    log_sheet.appendRow([...log_text, status]);
    if (message) ui.alert(message);
    return; // 早期終了用
  };

  //権限管理シートに一致する値が無ければ稼働終了
  const matched_data = authority_data.filter(elem =>
    elem[company_index] == request_obj[company_name] &&
    elem[site_index] == request_obj[site_name] &&
    elem[usage_index] == request_obj[usage_name] 
  );
  if(matched_data.length == 0) return finalize("該当データ無し","リクエスト内容が正しくありません");
  
  //権限が無ければ稼働終了
  if(!matched_data[0][account_index].includes(user_account)) return finalize("権限無し","パスワードの取得権限がありません");

  //無効になっている場合も終了
  if(matched_data[0][valid_index]!="有効") return finalize("PW無効","パスワードが無効になっています。");
  

  //PWを受け渡すロボを稼働する
  const pass_obj = {
    [company_name] : matched_data[0][company_index],
    [site_name] : matched_data[0][site_index],
    [usage_name] : matched_data[0][usage_index]
  }
  const property = PropertiesService.getScriptProperties();
  const url = property.getProperty("PW_DELIVERY_URL");
  const payload = JSON.stringify(pass_obj);
  const options = {
    'method' : 'post',
    'contentType': 'application/json',
    'payload' : payload
  };

  const response = UrlFetchApp.fetch(url,options);
  const res_data = JSON.parse(response.getContentText());
  
  //エラーの場合は稼働終了
  if(res_data.error) return finalize("PW受渡しエラー","パスワードが取得できませんでした。\n管理者に連絡してください。");

  //ダイアログにパスワードを表示
  showPwDialog(res_data);
  return finalize("受渡し成功");
}


function showPwDialog(passObj) {
  const tpl = HtmlService.createTemplateFromFile('PwDialog'); //HTMLファイル名を指定
  tpl.data = passObj;
  const html = tpl.evaluate().setWidth(520).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, '認証情報の表示');
}