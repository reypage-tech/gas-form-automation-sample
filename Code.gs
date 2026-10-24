/**
 * 問い合わせフォームの自動化（サンプル）
 *
 * Google フォームの回答をスプレッドシートに記録し、送信者に自動返信メールを送る。
 * 毎朝1回、前日からの新着と、まだ対応が終わっていない問い合わせの一覧を担当者にメールする。
 *
 * 店名・人物はすべて架空。初回は setup() をスクリプトエディタから1回だけ実行する。
 * 動作の確認は selfTest()、フォームから1件送信、checkFormSubmit() の順に行う。
 */

/** 設定。店ごとに変わる値はここだけにまとめる。 */
const CONFIG = {
  shopName: 'ピラティススタジオ ノギク（サンプル）',
  ownerEmail: '',          // 空欄なら setup() を実行したアカウントに送る
  sheetName: '問い合わせ',
  summaryHour: 8,          // 毎朝の一覧を送る時刻（スクリプトのタイムゾーン）
  timeZone: 'Asia/Tokyo',
  replyDeadline: '24時間以内',
  skipEmptySummary: true   // 新着も未完了もない日は一覧を送らない
};

/** フォームの質問タイトル。シートの列と自動返信の差し込みに使う。 */
const FIELDS = {
  name: 'お名前',
  email: 'メールアドレス',
  topic: 'ご用件',
  message: 'お問い合わせ内容'
};

const TOPICS = ['体験レッスン', '料金・プラン', 'その他'];
const STATUS = { NEW: '未対応', IN_PROGRESS: '対応中', DONE: '完了' };
const HEADERS = ['受付日時', 'お名前', 'メールアドレス', 'ご用件', 'お問い合わせ内容', '自動返信', '状況', 'メモ'];
const COL = { TIMESTAMP: 0, NAME: 1, EMAIL: 2, TOPIC: 3, MESSAGE: 4, REPLY: 5, STATUS: 6, NOTE: 7 };

/**
 * 初回の準備。フォーム、記録用のシート、2つのトリガーを作る。
 * すでに作ってある場合は新しく作らず、トリガーだけ作り直す。
 * 最初からやり直すときは、古いフォームとシートを削除してから resetSample() を実行する。
 */
function setup() {
  // 権限の確認画面で一部だけ許可された場合は、ここで止めてもう一度確認画面を出す
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);

  const props = PropertiesService.getScriptProperties();
  let ss = openOrNull_(() => SpreadsheetApp.openById(props.getProperty('SHEET_ID')), props.getProperty('SHEET_ID'));
  let form = openOrNull_(() => FormApp.openById(props.getProperty('FORM_ID')), props.getProperty('FORM_ID'));

  if (ss && form) {
    console.log('作成済みのフォームとシートを使います（新しくは作りません）');
  } else {
    ss = createSheet_();
    form = createForm_();
    props.setProperties({ SHEET_ID: ss.getId(), FORM_ID: form.getId() });
    console.log('フォームとシートを新しく作りました');
  }

  removeTriggers();
  ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('sendDailySummary').timeBased()
    .atHour(CONFIG.summaryHour).everyDays(1).inTimezone(CONFIG.timeZone).create();

  console.log('フォーム（編集）: ' + form.getEditUrl());
  console.log('フォーム（回答）: ' + form.getPublishedUrl());
  console.log('シート: ' + ss.getUrl());
  return { ss: ss, form: form };
}

/** 記録用のシートを作る。 */
function createSheet_() {
  const ss = SpreadsheetApp.create(CONFIG.shopName + ' 問い合わせ一覧');
  ss.setSpreadsheetTimeZone(CONFIG.timeZone);
  const sheet = ss.getSheets()[0].setName(CONFIG.sheetName);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // 「状況」列はプルダウンで選ぶ
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(STATUS), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, COL.STATUS + 1, sheet.getMaxRows() - 1, 1).setDataValidation(statusRule);
  return ss;
}

/** 問い合わせフォームを作る。2026年7月以降に作るフォームは公開の指定が要るので、公開した状態で作る。 */
function createForm_() {
  const form = FormApp.create(CONFIG.shopName + ' お問い合わせ', true)
    .setDescription('サンプル用のフォームです。実際の個人情報は入力しないでください。')
    .setCollectEmail(false);
  form.addTextItem().setTitle(FIELDS.name).setRequired(true);
  form.addTextItem().setTitle(FIELDS.email).setRequired(true)
    .setValidation(FormApp.createTextValidation().requireTextIsEmail().build());
  form.addMultipleChoiceItem().setTitle(FIELDS.topic).setChoiceValues(TOPICS).setRequired(true);
  form.addParagraphTextItem().setTitle(FIELDS.message);
  if (form.supportsAdvancedResponderPermissions() && !form.isPublished()) form.setPublished(true);
  return form;
}

/** ID が空か、開けないときは null を返す。 */
function openOrNull_(open, id) {
  if (!id) return null;
  try { return open(); } catch (err) { return null; }
}

/**
 * フォーム送信時のトリガー。自動返信を送り、結果と一緒にシートへ1行追加する。
 * 返信に失敗しても記録は残し、「メモ」に理由を書く。
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e
 */
function onFormSubmit(e) {
  const answers = parseAnswers_(e.response.getItemResponses());
  let replied = 'なし';
  let note = '';

  if (!isValidEmail_(answers.email)) {
    note = 'メールアドレスの形式が正しくない';
  } else {
    try {
      MailApp.sendEmail({
        to: answers.email,
        subject: '【' + CONFIG.shopName + '】お問い合わせを受け付けました',
        body: buildReplyBody_(answers),
        name: CONFIG.shopName
      });
      replied = '送信済み';
    } catch (err) {
      note = '自動返信に失敗: ' + err.message;
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSheet_().appendRow([
      new Date(), answers.name, answers.email, answers.topic, answers.message,
      replied, STATUS.NEW, note
    ]);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 毎朝のトリガー。過去24時間の新着と、状況が「完了」以外の問い合わせを1通にまとめて送る。
 */
function sendDailySummary() {
  const rows = getSheet_().getDataRange().getValues().slice(1);
  const summary = buildSummary_(rows, new Date(), CONFIG.timeZone);
  if (summary.newCount === 0 && summary.openCount === 0 && CONFIG.skipEmptySummary) {
    return;
  }
  MailApp.sendEmail({
    to: CONFIG.ownerEmail || Session.getEffectiveUser().getEmail(),
    subject: '【' + CONFIG.shopName + '】新着' + summary.newCount + '件・未完了' + summary.openCount + '件',
    body: summary.body,
    name: CONFIG.shopName
  });
}

/** このプロジェクトのトリガーをすべて削除する。止めたいときに手で実行する。 */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

/** setup() で覚えたフォームとシートの ID を消し、トリガーも消す。次の setup() で新しく作り直す。 */
function resetSample() {
  removeTriggers();
  PropertiesService.getScriptProperties().deleteAllProperties();
  console.log('ID とトリガーを消しました。古いフォームとシートは Google ドライブから手で削除してください');
}

/* ---- 動作の確認（Google 上で実行する） ---- */

const TEST_MARK = '[自動テスト]';

/**
 * 確認その1。setup() を行い（作成済みなら作り直さない）、項目ごとに OK / NG を実行ログに出す。
 * 自動返信と毎朝の一覧は、実行したアカウント自身のアドレスに1通ずつ届く。
 */
function selfTest() {
  const log = newCheckLog_();
  const made = setup();
  const ss = made.ss;
  const form = made.form;
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  const me = Session.getEffectiveUser().getEmail();

  // 1. 権限
  const auth = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  log.check('権限がすべて許可されている',
    auth.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.NOT_REQUIRED);
  console.log('許可された権限: ' + auth.getAuthorizedScopes().map(u => u.replace('https://www.googleapis.com/auth/', '')).join(', '));

  // 2. シート
  log.check('シートがある', !!sheet);
  const header = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  log.check('シートの見出しが8列そろっている', header.join('|') === HEADERS.join('|'));
  const rule = sheet.getRange(2, COL.STATUS + 1).getDataValidation();
  log.check('「状況」列にプルダウンがある', !!rule);
  log.check('シートのタイムゾーンが ' + CONFIG.timeZone, ss.getSpreadsheetTimeZone() === CONFIG.timeZone);

  // 3. フォーム
  const titles = form.getItems().map(i => i.getTitle());
  log.check('フォームの質問が4つそろっている',
    Object.keys(FIELDS).every(k => titles.indexOf(FIELDS[k]) >= 0) && titles.length === 4);
  log.check('フォームが回答を受け付けている', form.isAcceptingResponses());
  if (form.supportsAdvancedResponderPermissions()) log.check('フォームが公開されている', form.isPublished());

  // 4. トリガー
  const triggers = ScriptApp.getProjectTriggers();
  const submitT = triggers.filter(t => t.getHandlerFunction() === 'onFormSubmit' &&
    t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT && t.getTriggerSourceId() === form.getId());
  const clockT = triggers.filter(t => t.getHandlerFunction() === 'sendDailySummary' &&
    t.getEventType() === ScriptApp.EventType.CLOCK);
  log.check('フォーム送信時のトリガーが1つある', submitT.length === 1);
  log.check('毎朝のトリガーが1つある', clockT.length === 1);

  // 5. 自動返信（スクリプトから送った回答は送信時のトリガーを起こさないので、処理を直接呼ぶ）
  const items = form.getItems();
  const byTitle = {};
  items.forEach(i => { byTitle[i.getTitle()] = i; });
  const response = form.createResponse()
    .withItemResponse(byTitle[FIELDS.name].asTextItem().createResponse('テスト 花子（自動テスト）'))
    .withItemResponse(byTitle[FIELDS.email].asTextItem().createResponse(me))
    .withItemResponse(byTitle[FIELDS.topic].asMultipleChoiceItem().createResponse(TOPICS[0]))
    .withItemResponse(byTitle[FIELDS.message].asParagraphTextItem().createResponse(TEST_MARK + ' 土曜の午前は空いていますか'))
    .submit();
  const before = sheet.getLastRow();
  onFormSubmit({ response: response, source: form });
  const row = sheet.getRange(sheet.getLastRow(), 1, 1, HEADERS.length).getValues()[0];
  log.check('回答がシートに1行増えた', sheet.getLastRow() === before + 1);
  log.check('自動返信を送れた（「自動返信」列が「送信済み」）', row[COL.REPLY] === '送信済み',
    row[COL.REPLY] === '送信済み' ? '' : String(row[COL.NOTE]));
  log.check('新しい行の「状況」が「未対応」', row[COL.STATUS] === STATUS.NEW);
  log.check('受付日時が日付として入っている', row[COL.TIMESTAMP] instanceof Date);

  // 6. 形式の正しくないアドレス（フォームでは弾かれるので、処理だけを確かめる）
  onFormSubmit({ response: fakeResponse_({ name: 'テスト 次郎（自動テスト）', email: 'jiro-at-example',
    topic: TOPICS[1], message: TEST_MARK }) });
  const bad = sheet.getRange(sheet.getLastRow(), 1, 1, HEADERS.length).getValues()[0];
  log.check('形式の正しくないアドレスには返信せず、メモに理由が残る',
    bad[COL.REPLY] === 'なし' && String(bad[COL.NOTE]).indexOf('形式') >= 0);

  // 7. 毎朝の一覧（その場で1回送る）
  let summaryOk = true;
  let summaryErr = '';
  try { sendDailySummary(); } catch (err) { summaryOk = false; summaryErr = err.message; }
  log.check('毎朝の一覧メールを送れた', summaryOk, summaryErr);

  console.log('今日あと送れるメールの数（参考）: ' + MailApp.getRemainingDailyQuota());
  PropertiesService.getScriptProperties().setProperty('TEST_AT', String(Date.now()));
  log.finish('selfTest');
  console.log('次の作業: 下のURLのフォームを開いて1件送信し、そのあと checkFormSubmit を実行してください');
  console.log('フォーム（回答）: ' + form.getPublishedUrl());
}

/**
 * 確認その2。selfTest() のあとにフォームから送った回答が、トリガーでシートに記録され、
 * 自動返信が送られたかを確かめる。
 */
function checkFormSubmit() {
  const log = newCheckLog_();
  const props = PropertiesService.getScriptProperties();
  const since = Number(props.getProperty('TEST_AT') || 0);
  if (!since) {
    console.log('NG  先に selfTest を実行してください');
    return;
  }
  const form = FormApp.openById(props.getProperty('FORM_ID'));
  const responses = form.getResponses().filter(r => r.getTimestamp().getTime() > since &&
    parseAnswers_(r.getItemResponses()).message.indexOf(TEST_MARK) < 0);
  const rows = getSheet_().getDataRange().getValues().slice(1).filter(r =>
    r[COL.TIMESTAMP] instanceof Date && r[COL.TIMESTAMP].getTime() > since &&
    String(r[COL.MESSAGE]).indexOf(TEST_MARK) < 0);

  log.check('フォームに新しい回答が届いている', responses.length > 0,
    responses.length > 0 ? '' : 'フォームの送信がまだのようです');
  log.check('その回答がトリガーでシートに記録された', rows.length > 0,
    rows.length > 0 ? '' : (responses.length > 0 ? '1分ほど待って、もう一度実行してください' : ''));
  if (rows.length > 0) {
    const last = rows[rows.length - 1];
    log.check('自動返信を送れた（「自動返信」列が「送信済み」）', last[COL.REPLY] === '送信済み',
      last[COL.REPLY] === '送信済み' ? '' : String(last[COL.NOTE]));
    console.log('（参考）送信先は実行したアカウントと同じアドレス: ' +
      (String(last[COL.EMAIL]).toLowerCase() === Session.getEffectiveUser().getEmail().toLowerCase() ? 'はい' : 'いいえ'));
  }
  console.log('（参考）新しい回答 ' + responses.length + ' 件、シートの新しい行 ' + rows.length + ' 件');
  log.finish('checkFormSubmit');
}

/** フォームを通さずに処理を試すための、回答の代わり。 */
function fakeResponse_(a) {
  const item = (title, value) => ({ getItem: () => ({ getTitle: () => title }), getResponse: () => value });
  return {
    getItemResponses: () => Object.keys(FIELDS).map(k => item(FIELDS[k], a[k] || ''))
  };
}

/** OK / NG を1行ずつ出し、最後に件数をまとめる。 */
function newCheckLog_() {
  let ok = 0;
  let ng = 0;
  return {
    check: (label, passed, detail) => {
      if (passed) ok++; else ng++;
      console.log((passed ? 'OK  ' : 'NG  ') + label + (!passed && detail ? '（' + detail + '）' : ''));
    },
    finish: name => console.log('== ' + name + ' の結果: OK ' + ok + ' 件 / NG ' + ng + ' 件 =='),
    counts: () => ({ ok: ok, ng: ng })
  };
}

/* ---- ここから下は Google のサービスを使わない関数（test/local_test.js で確認） ---- */

/**
 * フォームの回答を、FIELDS のキーを持つオブジェクトにする。
 * @param {Array<{getItem: Function, getResponse: Function}>} itemResponses
 * @return {{name: string, email: string, topic: string, message: string}}
 */
function parseAnswers_(itemResponses) {
  const byTitle = {};
  itemResponses.forEach(ir => {
    byTitle[String(ir.getItem().getTitle()).trim()] = String(ir.getResponse() || '').trim();
  });
  const out = {};
  Object.keys(FIELDS).forEach(key => { out[key] = byTitle[FIELDS[key]] || ''; });
  return out;
}

/**
 * メールアドレスの形を簡単に確かめる（フォーム側でも検証しているので、ここは念のため）。
 * @param {string} value
 * @return {boolean}
 */
function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ''));
}

/**
 * 自動返信の本文を作る。
 * @param {{name: string, topic: string, message: string}} a
 * @return {string}
 */
function buildReplyBody_(a) {
  return [
    (a.name || 'お客') + '様',
    '',
    CONFIG.shopName + 'へお問い合わせいただき、ありがとうございます。',
    'このメールは自動でお送りしています。',
    '内容を確認し、' + CONFIG.replyDeadline + 'にご連絡します。',
    '',
    '―― お問い合わせの内容 ――',
    'ご用件：' + (a.topic || '（未選択）'),
    a.message || '（本文なし）',
    '',
    '※このメールはサンプル用の架空の店舗から送信しています。'
  ].join('\n');
}

/**
 * 毎朝の一覧の本文と件数を作る。
 * @param {Array<Array<*>>} rows シートの2行目以降
 * @param {Date} now
 * @param {string} timeZone
 * @return {{newCount: number, openCount: number, body: string}}
 */
function buildSummary_(rows, now, timeZone) {
  const since = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const valid = rows.filter(r => r[COL.TIMESTAMP] instanceof Date);
  const recent = valid.filter(r => r[COL.TIMESTAMP] >= since);
  const open = valid.filter(r => r[COL.STATUS] !== STATUS.DONE);

  const lines = ['過去24時間の新着：' + recent.length + '件'];
  const byTopic = {};
  recent.forEach(r => { byTopic[r[COL.TOPIC]] = (byTopic[r[COL.TOPIC]] || 0) + 1; });
  TOPICS.filter(t => byTopic[t]).forEach(t => lines.push('・' + t + '：' + byTopic[t] + '件'));

  lines.push('', '対応が終わっていない問い合わせ：' + open.length + '件');
  open.forEach(r => {
    const when = formatDate_(r[COL.TIMESTAMP], timeZone);
    const flag = r[COL.REPLY] === '送信済み' ? '' : '　※自動返信なし（' + (r[COL.NOTE] || '理由不明') + '）';
    lines.push(when + '　' + (r[COL.STATUS] || STATUS.NEW) + '　' + r[COL.NAME] + '（' + r[COL.TOPIC] + '）' + flag);
  });

  return { newCount: recent.length, openCount: open.length, body: lines.join('\n') };
}

/**
 * 日時を「MM/dd HH:mm」にする。Apps Script の外（テスト）では UTC のまま同じ形にする。
 * @param {Date} date
 * @param {string} timeZone
 * @return {string}
 */
function formatDate_(date, timeZone) {
  if (typeof Utilities !== 'undefined') {
    return Utilities.formatDate(date, timeZone, 'MM/dd HH:mm');
  }
  return date.toISOString().slice(5, 16).replace('-', '/').replace('T', ' ');
}

/** @return {GoogleAppsScript.Spreadsheet.Sheet} */
function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('SHEET_ID がありません。先に setup() を実行してください。');
  return SpreadsheetApp.openById(id).getSheetByName(CONFIG.sheetName);
}

// Node でテストするときだけ使う（Apps Script では module は定義されていない）
if (typeof module !== 'undefined') {
  module.exports = {
    CONFIG, FIELDS, STATUS, COL,
    onFormSubmit, sendDailySummary, setup, selfTest, checkFormSubmit, resetSample,
    parseAnswers_, isValidEmail_, buildReplyBody_, buildSummary_, newCheckLog_
  };
}
