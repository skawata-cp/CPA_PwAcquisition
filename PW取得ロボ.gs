/**
 * PW取得ロボ
 * リクエスト内容が正しいか、権限があるかを判定し、
 * 権限があればPWを受け渡すロボからPWを受け取る
 */
function PwAcquisitionDecision(company, site, usage) {
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

  //リクエストログシート更新用の設定
  const log_text = [date, user_account, company, site, usage];

  //稼働終了時の関数
  const finalize = (status, message) => {
    log_sheet.appendRow([...log_text, status]);
    if (message) ui.alert(message);
    return; // 早期終了用
  };

  //========ここから権限確認===========
  //権限管理シートの値を検索
  const authority_data = authority_sheet.getRange("A1:F").getDisplayValues(); //権限管理シートの値を取得
  const authority_header = authority_data.shift(); //ヘッダー取得
  
  //各項目の列番号
  const company_index = authority_header.indexOf(company_name);
  const site_index = authority_header.indexOf(site_name);
  const usage_index = authority_header.indexOf(usage_name);
  const account_index = authority_header.indexOf(account_name);
  const valid_index = authority_header.indexOf(valid_name);

  //権限管理シートに一致する値が無ければ稼働終了
  const matched_data = authority_data.filter(elem =>
    elem[company_index] == company &&
    elem[site_index] == site &&
    elem[usage_index] == usage 
  );
  if(matched_data.length == 0){
    finalize("該当データ無し","リクエスト内容が正しくありません");
    return;
  } 
  
  //権限が無ければ稼働終了
  if(!matched_data[0][account_index].includes(user_account)){
    finalize("権限無し","パスワードの取得権限がありません");
    return;
  } 

  //無効になっている場合も終了
  if(matched_data[0][valid_index]!="有効"){
    finalize("PW無効","パスワードが無効になっています。");
    return;
  } 
  //========権限確認終了===========
  

  //PWを受け渡すロボを稼働する
  const pass_obj = {
    [company_name] : company,
    [site_name] : site,
    [usage_name] : usage
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
  finalize("受渡し成功");
  return res_data; 
}