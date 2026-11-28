/**
 * 介護記録アプリ ⇔ シャイン管理アプリ 連携（管理アプリのApps Scriptに「kiroku」という名前のファイルで追加する）
 *
 * できること
 *  1. 介護記録アプリの連絡帳PDFを、今の連絡帳フォルダに「YYYYMMDD/連絡帳YYYYMMDD_名前様.pdf」で保存し、
 *     「連絡帳」シートにも1行足す（ご家族のページにすぐ出る。LINE通知はこれまで通り）
 *  2. 公式アカウントをスタッフのグループLINEに入れておくと、「名前」→「写真」で写真をドライブに保存し、
 *     介護記録アプリの連絡帳に取り込めるようにする
 *  3. 合言葉（スクリプト プロパティ ADMIN_KEY）を入れると、管理用の機能を合言葉なしでは使えなくする
 *
 * 使う設定（スクリプト プロパティ）
 *  ADMIN_KEY : 管理アプリ・介護記録アプリと共通の合言葉。空のあいだは今まで通り誰でも使える状態
 */

var KIROKU_PHOTO_ROOT = '介護記録_連絡帳写真';   // マイドライブ直下に自動で作るフォルダ名
var KIROKU_PDF_ROOT   = '介護記録_経過記録PDF';
var KIROKU_PHOTO_MIN  = 10;                    // 名前を送ってから何分間の写真をその人の分にするか

/* ===== 共通 ===== */
function kirokuJson_(o){ o.ok = o.status === 'ok'; return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function adminDeny_(key){
  var k = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
  if (!k || key === k) return null;           // 合言葉が未設定なら今まで通り
  return kirokuJson_({status:'error', message:'合言葉が違います'});
}
function kirokuSheet_(name, header){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function kirokuFolder_(parent, name){ var it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function kirokuRoot_(name){ var it = DriveApp.getRootFolder().getFoldersByName(name); return it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(name); }
function kirokuToday_(){ return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'); }
function kirokuKey_(s){ return String(s||'').replace(/[\s　]/g,'').replace(/(様|さま|さん)$/,'').replace(/^（障）/,'').replace(/髙/g,'高').replace(/﨑/g,'崎'); }

/* ===== GET（管理アプリの doGet から呼ばれる。介護記録用でなければ null） ===== */
function kirokuGet_(e){
  var p = e.parameter;
  if (p.action === 'ping') return kirokuJson_({status:'ok', group: PropertiesService.getScriptProperties().getProperty('KIROKU_GROUP_ID') ? '登録済み' : '未登録'});
  if (p.action === 'photos') {
    var rows = kirokuSheet_('連絡帳写真', ['受信日時','日付','利用者名','ファイルID','送信者']).getDataRange().getValues().slice(1);
    var list = rows.filter(function(r){
      var d = r[1] instanceof Date ? Utilities.formatDate(r[1],'Asia/Tokyo','yyyy-MM-dd') : String(r[1]);
      return d === p.date;
    }).map(function(r){ return {at: String(r[0]), name: r[2], id: r[3]}; });
    return kirokuJson_({status:'ok', photos: list});
  }
  if (p.action === 'photo') {
    var blob = DriveApp.getFileById(p.id).getBlob();
    return kirokuJson_({status:'ok', type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes())});
  }
  return null;
}

/* ===== POST（管理アプリの doPost から呼ばれる。介護記録用でなければ null） ===== */
function kirokuPost_(json){
  if (json.action === 'uploadPdf') {
    try { return kirokuJson_(json.kind === '連絡帳' ? kirokuUploadReport_(json) : kirokuUploadOther_(json)); }
    catch (err) { return kirokuJson_({status:'error', message: String(err)}); }
  }
  if (json.action === 'roster') {
    var sh = kirokuSheet_('介護記録_名簿', ['名前']);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
    var names = json.names || [];
    if (names.length) sh.getRange(2, 1, names.length, 1).setValues(names.map(function(n){ return [n]; }));
    return kirokuJson_({status:'ok', count: names.length});
  }
  return null;
}

// 連絡帳PDF：今の自動アップロードツールと同じ置き場所・名前。シートにもすぐ登録する
function kirokuUploadReport_(b){
  var ymd  = String(b.date || kirokuToday_()).replace(/-/g, '');
  var name = String(b.name || '').replace(/[\s　]/g, '');       // 空白だけ取る（（障）・髙はそのまま）
  if (!name) return {status:'error', message:'名前がありません'};
  var folder = kirokuFolder_(DriveApp.getFolderById(FOLDER_RENNRAKUCHO), ymd);
  var fname = '連絡帳' + ymd + '_' + name + '様.pdf', n = 1;
  while (folder.getFilesByName(fname).hasNext()) fname = '連絡帳' + ymd + '_' + name + '様_' + (n++) + '.pdf';
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b.data), 'application/pdf', fname));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var d = ymd.substring(0,4) + '/' + ymd.substring(4,6) + '/' + ymd.substring(6,8);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('連絡帳').appendRow([name, d, file.getId(), false]); // 5分ごとの同期を待たずに登録
  return {status:'ok', fileId: file.getId(), name: fname};
}
// 経過記録など：介護記録_経過記録PDF / YYYY-MM / ファイル名（同じ名前は上書き）
function kirokuUploadOther_(b){
  var folder = kirokuFolder_(kirokuRoot_(KIROKU_PDF_ROOT), String(b.date || kirokuToday_()).slice(0, 7));
  var name = String(b.filename || '記録.pdf');
  var old = folder.getFilesByName(name); while (old.hasNext()) old.next().setTrashed(true);
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b.data), 'application/pdf', name));
  return {status:'ok', fileId: file.getId(), name: name};
}

/* ===== グループLINE（管理アプリの processEvent の先頭から呼ばれる） ===== */
// ご家族との1対1トーク以外（グループ・複数人トーク）はすべてここで処理し、利用者登録には進ませない
function kirokuGroupEvent_(event){
  var src = event.source || {};
  if (event.type !== 'message') return;
  var props = PropertiesService.getScriptProperties();
  var gid = src.groupId || src.roomId;
  var registered = props.getProperty('KIROKU_GROUP_ID');
  if (!registered) props.setProperty('KIROKU_GROUP_ID', gid);   // 最初に話しかけられたグループを写真用に登録
  else if (registered !== gid) return;                          // それ以外のグループは無視

  var cache = CacheService.getScriptCache();
  var ck = 'kiroku_' + gid + '_' + (src.userId || '');
  var m = event.message;

  if (m.type === 'text') {
    var hit = kirokuFindName_(m.text);
    if (hit) {
      cache.put(ck, hit, KIROKU_PHOTO_MIN * 60);
      kirokuReply_(event.replyToken, hit + ' 様ですね。続けて写真を送ってください（' + KIROKU_PHOTO_MIN + '分間）');
    }
    return;                                                     // ふだんの会話には反応しない
  }
  if (m.type === 'image') {
    var who = cache.get(ck);
    var set = m.imageSet;
    var isLast = !set || set.index === set.total;
    if (!who) { if (isLast) kirokuReply_(event.replyToken, '先に利用者さんの名前だけを送ってから、写真を送ってください'); return; }
    var res = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/message/' + m.id + '/content', {headers: {Authorization: 'Bearer ' + LINE_TOKEN}});
    var d = kirokuToday_();
    var file = kirokuFolder_(kirokuRoot_(KIROKU_PHOTO_ROOT), d).createFile(res.getBlob().setName(d + '_' + who + '_' + m.id + '.jpg'));
    kirokuSheet_('連絡帳写真', ['受信日時','日付','利用者名','ファイルID','送信者'])
      .appendRow([Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'), d, who, file.getId(), src.userId || '']);
    cache.put(ck, who, KIROKU_PHOTO_MIN * 60);
    if (isLast) kirokuReply_(event.replyToken, who + ' 様の連絡帳に写真を入れました' + (set ? '（' + set.total + '枚）' : ''));
  }
}
// 名前だけのメッセージかどうか。介護記録アプリの名簿・利用者リスト・連絡帳シートの名前と照合
function kirokuFindName_(text){
  var t = kirokuKey_(text);
  if (!t || t.length > 12) return '';
  var ss = SpreadsheetApp.getActiveSpreadsheet(), names = {};
  var add = function(n){ n = String(n||'').trim(); if (n && n.length <= 14 && n.indexOf('\n') < 0) names[n] = true; };
  var roster = ss.getSheetByName('介護記録_名簿');
  if (roster) roster.getDataRange().getValues().slice(1).forEach(function(r){ add(r[0]); });
  ss.getSheetByName('利用者リスト').getDataRange().getValues().slice(1).forEach(function(r){ add(r[0]); });
  var rep = ss.getSheetByName('連絡帳'), last = rep.getLastRow();
  if (last > 1) rep.getRange(Math.max(2, last - 800), 1, Math.min(800, last - 1), 1).getValues().forEach(function(r){ add(r[0]); });
  var hits = {};
  Object.keys(names).forEach(function(n){ if (kirokuKey_(n) === t) hits[kirokuKey_(n)] = n.replace(/[\s　]/g, ''); });
  var keys = Object.keys(hits);
  return keys.length === 1 ? hits[keys[0]] : '';
}
function kirokuReply_(token, text){   // 返信（reply）はLINEの無料通数に数えられない
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {Authorization: 'Bearer ' + LINE_TOKEN},
    payload: JSON.stringify({replyToken: token, messages: [{type: 'text', text: text}]})
  });
}
