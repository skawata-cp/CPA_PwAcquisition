function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('秘密情報')
    .addItem('PW取得', 'showPwRequestSidebar')
    .addToUi();
}

function showPwRequestSidebar() {
  // 先にサーバで取得 → HTMLに埋め込む
  const tpl = HtmlService.createTemplateFromFile('PwRequestSidebar');
  tpl.tree = getAuthTree();                 // ← ここでツリーを用意
  const html = tpl.evaluate().setTitle('PWリクエスト選択');
  SpreadsheetApp.getUi().showSidebar(html);
}



//サイドバー作成
function getAuthTree() {
  var cache = CacheService.getScriptCache();
  var key = 'AUTH_TREE_V1';
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  var sh = SpreadsheetApp.getActive().getSheetByName('A-1：権限管理');
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return { companies: [], sitesByCompany: {}, usagesByCompanyAndSite: {} };

  var vals = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var header = vals.shift();

  var iC = header.indexOf('会社名');
  var iS = header.indexOf('サイト・システム名');
  var iU = header.indexOf('用途');

  // Set の代わりにオブジェクトを集合として使う
  var companiesSet = {};
  var sitesByCompanySet = {};            // { company: {site:true,...} }
  var usagesByCompanyAndSiteSet = {};    // { company: { site: {usage:true,...} } }

  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var c = r[iC], s = r[iS], u = r[iU];
    if (!c || !s || !u) continue;

    companiesSet[c] = true;

    if (!sitesByCompanySet[c]) sitesByCompanySet[c] = {};
    sitesByCompanySet[c][s] = true;

    if (!usagesByCompanyAndSiteSet[c]) usagesByCompanyAndSiteSet[c] = {};
    if (!usagesByCompanyAndSiteSet[c][s]) usagesByCompanyAndSiteSet[c][s] = {};
    usagesByCompanyAndSiteSet[c][s][u] = true;
  }

  // 集合 → ソート済み配列へ
  var companies = Object.keys(companiesSet);

  var sitesByCompany = {};
  for (var c in sitesByCompanySet) {
    sitesByCompany[c] = Object.keys(sitesByCompanySet[c]);
  }

  var usagesByCompanyAndSite = {};
  for (var c2 in usagesByCompanyAndSiteSet) {
    usagesByCompanyAndSite[c2] = {};
    var siteMap = usagesByCompanyAndSiteSet[c2];
    for (var s2 in siteMap) {
      usagesByCompanyAndSite[c2][s2] = Object.keys(siteMap[s2]);
    }
  }

  var tree = {
    companies: companies,
    sitesByCompany: sitesByCompany,
    usagesByCompanyAndSite: usagesByCompanyAndSite
  };

  cache.put(key, JSON.stringify(tree), 300); // 5分キャッシュ
  return tree;
}
