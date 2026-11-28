/**
 * 介護記録アプリ 連携用プログラム（Google Apps Script）
 * できること
 *  1. LINEグループで「利用者の名前」→「写真」を送ると、写真をGoogleドライブに保存して連絡帳に取り込めるようにする
 *  2. アプリで作ったPDF（経過記録・連絡帳）をGoogleドライブに保存する
 *
 * 最初に「スクリプト プロパティ」に次を入れてください（設定手順書を参照）
 *  LINE_TOKEN : LINE公式アカウントのチャネルアクセストークン（長期）
 *  APP_KEY    : アプリとつなぐための合言葉（設定用ページで1回だけ使う。スタッフは入力不要）
 *  GROUP_ID   : 空でOK。最初に写真を受け取ったグループが自動で登録されます
 */

const ROOT_FOLDER_NAME = 'デイサービスシャイン_介護記録';
// シャイン管理アプリが見ている連絡帳フォルダ（rename_tool.py と同じ）。
// ここに「日付(YYYYMMDD)のフォルダ / 連絡帳YYYYMMDD_名前様.pdf」で入れると、今のDrive同期・LINE通知にそのまま乗る
const RENRAKU_FOLDER_ID = '10PMUjdRbxhR73WUuhw-PRrYbCTBrZpYf';
const PHOTO_MINUTES = 10;   // 名前を送ってから何分間の写真をその人の分にするか

function prop_(k){ return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function setProp_(k,v){ PropertiesService.getScriptProperties().setProperty(k,v); }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, header){
  let sh = ss_().getSheetByName(name);
  if(!sh){ sh = ss_().insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function folder_(path){ // 例: ['連絡帳写真','2026-11-28']
  let f; const it = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  f = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
  path.forEach(p => { const c = f.getFoldersByName(p); f = c.hasNext() ? c.next() : f.createFolder(p); });
  return f;
}
function today_(){ return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'); }
function now_(){ return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); }
function norm_(s){ return String(s||'').replace(/[\s　]/g,'').replace(/(様|さま|さん)$/,'').replace(/髙/g,'高').replace(/﨑/g,'崎').replace(/^（障）/,''); }

/* ===== 受け口（POST） ===== */
function doPost(e){
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) { return json_({ok:false, error:'bad json'}); }
  if (body.events) { handleLine_(body.events); return json_({ok:true}); }   // LINEから
  if (body.key !== prop_('APP_KEY')) return json_({ok:false, error:'合言葉が違います'});
  if (body.action === 'uploadPdf') return json_(uploadPdf_(body));
  if (body.action === 'roster') return json_(saveRoster_(body.names || []));
  return json_({ok:false, error:'unknown action'});
}

/* ===== 受け口（GET）: アプリが写真を取りに来る ===== */
function doGet(e){
  const p = e.parameter || {};
  if (p.key !== prop_('APP_KEY')) return json_({ok:false, error:'合言葉が違います'});
  if (p.action === 'ping') return json_({ok:true, group: prop_('GROUP_ID') ? '登録済み' : '未登録'});
  if (p.action === 'photos') {
    const sh = sheet_('写真', ['受信日時','日付','利用者名','ファイルID','URL','送信者']);
    const rows = sh.getDataRange().getValues().slice(1).filter(r => String(r[1]) === p.date || (r[1] instanceof Date && Utilities.formatDate(r[1],'Asia/Tokyo','yyyy-MM-dd') === p.date));
    return json_({ok:true, photos: rows.map(r => ({at: String(r[0]), name: r[2], id: r[3]}))});
  }
  if (p.action === 'photo') {
    const blob = DriveApp.getFileById(p.id).getBlob();
    return json_({ok:true, type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes())});
  }
  return json_({ok:false, error:'unknown action'});
}

/* ===== LINE ===== */
function handleLine_(events){
  const cache = CacheService.getScriptCache();
  events.forEach(ev => {
    const src = ev.source || {};
    if (src.type !== 'group' || ev.type !== 'message') return;
    const gid = prop_('GROUP_ID');
    if (!gid) setProp_('GROUP_ID', src.groupId);          // 最初のグループを登録
    else if (gid !== src.groupId) return;                   // 登録したグループ以外は無視
    const ck = 'who_' + src.groupId + '_' + (src.userId || '');
    const m = ev.message;

    if (m.type === 'text') {
      const hit = findUser_(m.text);
      if (hit) {
        cache.put(ck, hit, PHOTO_MINUTES * 60);
        reply_(ev.replyToken, hit + ' 様ですね。続けて写真を送ってください（' + PHOTO_MINUTES + '分間）');
      }
      return; // 名前以外の会話は何もしない
    }
    if (m.type === 'image') {
      const who = cache.get(ck);
      const set = m.imageSet;                       // 複数枚まとめて送ったとき
      const isLast = !set || set.index === set.total;
      if (!who) { if (isLast) reply_(ev.replyToken, '先に利用者さんの名前だけを送ってから、写真を送ってください'); return; }
      const res = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/message/' + m.id + '/content', {headers:{Authorization:'Bearer ' + prop_('LINE_TOKEN')}});
      const d = today_();
      const blob = res.getBlob().setName(d + '_' + who + '_' + m.id + '.jpg');
      const file = folder_(['連絡帳写真', d]).createFile(blob);
      sheet_('写真', ['受信日時','日付','利用者名','ファイルID','URL','送信者']).appendRow([now_(), d, who, file.getId(), file.getUrl(), src.userId || '']);
      cache.put(ck, who, PHOTO_MINUTES * 60); // 続けて送れるよう延長
      if (isLast) reply_(ev.replyToken, who + ' 様の連絡帳に写真を入れました' + (set ? '（' + set.total + '枚）' : ''));
    }
  });
}
function findUser_(text){
  const t = norm_(text); if (!t || t.length > 12) return '';
  const names = sheet_('名簿', ['名前']).getDataRange().getValues().slice(1).map(r => String(r[0])).filter(Boolean);
  const exact = names.filter(n => norm_(n) === t);
  return exact.length === 1 ? exact[0] : '';
}
function reply_(token, text){
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    headers:{Authorization:'Bearer ' + prop_('LINE_TOKEN')},
    payload: JSON.stringify({replyToken: token, messages:[{type:'text', text: text}]})
  });
}

/* ===== アプリから ===== */
function saveRoster_(names){
  const sh = sheet_('名簿', ['名前']);
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).clearContent();
  if (names.length) sh.getRange(2, 1, names.length, 1).setValues(names.map(n => [n]));
  return {ok:true, count: names.length};
}
function subfolder_(parent, name){ const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function uploadPdf_(b){
  if (b.kind === '連絡帳') return uploadRenraku_(b);
  return uploadOther_(b);
}
// 連絡帳：管理アプリと同じ置き場所・名前・共有設定にする
function uploadRenraku_(b){
  const ymd = String(b.date || today_()).replace(/-/g, '');            // 20261128
  const name = norm_(b.name);                                        // 空白なし（例：上田キヨ）
  const folder = subfolder_(DriveApp.getFolderById(RENRAKU_FOLDER_ID), ymd);
  let fname = '連絡帳' + ymd + '_' + name + '様.pdf', n = 1;
  while (folder.getFilesByName(fname).hasNext()) fname = '連絡帳' + ymd + '_' + name + '様_' + (n++) + '.pdf'; // 同名は連番（rename_tool と同じ）
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b.data), 'application/pdf', fname));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sheet_('PDF', ['保存日時','種類','日付','ファイル名','URL']).appendRow([now_(), '連絡帳', b.date, fname, file.getUrl()]);
  return {ok:true, url: file.getUrl(), name: fname};
}
function uploadOther_(b){
  // b.kind: '経過記録' | '連絡帳'、b.date: 'yyyy-MM-dd'、b.filename、b.data: base64
  const month = String(b.date || today_()).slice(0, 7);
  const folder = folder_([b.kind === '連絡帳' ? '連絡帳PDF' : '経過記録PDF', month]);
  const name = String(b.filename || '記録.pdf');
  const old = folder.getFilesByName(name); while (old.hasNext()) old.next().setTrashed(true); // 同じ名前は上書き
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b.data), 'application/pdf', name));
  sheet_('PDF', ['保存日時','種類','日付','ファイル名','URL']).appendRow([now_(), b.kind, b.date, name, file.getUrl()]);
  return {ok:true, url: file.getUrl()};
}
