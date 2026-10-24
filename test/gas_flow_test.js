// selfTest() と checkFormSubmit() の流れを、Google のサービスの偽物で Node 上で通す。
// 本物の Google での動きは確かめられない（トリガーが実際に動くか、権限の画面など）。
// 実行: node test/gas_flow_test.js
const assert = require('assert');
const fs = require('fs');
const path = require('path');

const logs = [];
console.log = (...a) => logs.push(a.join(' '));
const out = (...a) => process.stdout.write(a.join(' ') + '\n');

const sent = [];
const props = {};
let triggers = [];
let sheetRows = [];
let validationOn = false;
const responses = [];

const makeSheet = () => ({
  setName() { return this; },
  getMaxRows: () => 1000,
  setFrozenRows() {},
  getRange: (r, c, nr, nc) => ({
    setValues: v => { sheetRows[r - 1] = v[0].slice(); return { setFontWeight() {} }; },
    setDataValidation: () => { validationOn = true; },
    getDataValidation: () => (validationOn ? {} : null),
    getValues: () => [ (sheetRows[r - 1] || []).slice(c - 1, c - 1 + (nc || 1)) ]
  }),
  appendRow: r => sheetRows.push(r),
  getLastRow: () => sheetRows.length,
  getDataRange: () => ({ getValues: () => sheetRows.map(r => r.slice()) })
});
const sheet = makeSheet();
let tz = 'Etc/GMT';
const ss = {
  getId: () => 'SHEET1', getUrl: () => 'https://example.invalid/sheet',
  setSpreadsheetTimeZone: z => { tz = z; }, getSpreadsheetTimeZone: () => tz,
  getSheets: () => [sheet], getSheetByName: () => sheet
};

const items = [];
const mkItem = type => {
  const it = {
    title: '', setTitle(t) { this.title = t; return this; }, getTitle() { return this.title; },
    setRequired() { return this; }, setValidation() { return this; }, setChoiceValues() { return this; },
    createResponse(v) { return { getItem: () => it, getResponse: () => v }; }
  };
  it.asTextItem = it.asMultipleChoiceItem = it.asParagraphTextItem = () => it;
  items.push(it);
  return it;
};
let published = false;
const form = {
  getId: () => 'FORM1', getEditUrl: () => 'https://example.invalid/edit', getPublishedUrl: () => 'https://example.invalid/view',
  setDescription() { return this; }, setCollectEmail() { return this; },
  addTextItem: () => mkItem('text'), addMultipleChoiceItem: () => mkItem('mc'), addParagraphTextItem: () => mkItem('p'),
  supportsAdvancedResponderPermissions: () => true, isPublished: () => published, setPublished(v) { published = v; return this; },
  isAcceptingResponses: () => true, getItems: () => items,
  getResponses: () => responses,
  createResponse: () => {
    const irs = [];
    const fr = {
      withItemResponse(x) { irs.push(x); return fr; },
      submit() { const ts = new Date(); const r = { getItemResponses: () => irs, getTimestamp: () => ts }; responses.push(r); return r; }
    };
    return fr;
  }
};

let created = { ss: 0, form: 0 };
global.SpreadsheetApp = {
  create: () => { created.ss++; return ss; },
  openById: id => { if (id !== 'SHEET1') throw new Error('not found'); return ss; },
  newDataValidation: () => { const b = { requireValueInList: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; }
};
global.FormApp = {
  create: (t, pub) => { created.form++; published = !!pub; return form; },
  openById: id => { if (id !== 'FORM1') throw new Error('not found'); return form; },
  createTextValidation: () => { const b = { requireTextIsEmail: () => b, build: () => ({}) }; return b; }
};
global.MailApp = { sendEmail: m => sent.push(m), getRemainingDailyQuota: () => 100 - sent.length };
global.Session = { getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) };
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: k => (k in props ? props[k] : null),
  setProperty: (k, v) => { props[k] = v; }, setProperties: o => Object.assign(props, o),
  deleteAllProperties: () => Object.keys(props).forEach(k => delete props[k])
}) };
global.LockService = { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) };
const mkTrigger = (fn, type, src) => ({ getHandlerFunction: () => fn, getEventType: () => type, getTriggerSourceId: () => src });
global.ScriptApp = {
  AuthMode: { FULL: 'FULL' }, AuthorizationStatus: { NOT_REQUIRED: 'NOT_REQUIRED' },
  EventType: { ON_FORM_SUBMIT: 'ON_FORM_SUBMIT', CLOCK: 'CLOCK' },
  requireAllScopes: () => {},
  getAuthorizationInfo: () => ({ getAuthorizationStatus: () => 'NOT_REQUIRED',
    getAuthorizedScopes: () => ['https://www.googleapis.com/auth/forms', 'https://www.googleapis.com/auth/script.send_mail'] }),
  getProjectTriggers: () => triggers.slice(),
  deleteTrigger: t => { triggers = triggers.filter(x => x !== t); },
  newTrigger: fn => {
    const b = {
      forForm: f => ({ onFormSubmit: () => ({ create: () => triggers.push(mkTrigger(fn, 'ON_FORM_SUBMIT', f.getId())) }) }),
      timeBased: () => { const t = { atHour: () => t, everyDays: () => t, inTimezone: () => t,
        create: () => triggers.push(mkTrigger(fn, 'CLOCK', null)) }; return t; }
    };
    return b;
  }
};

const mod = { exports: {} };
const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
new Function('module', 'exports', code)(mod, mod.exports);
const g = mod.exports;
const ngLines = () => logs.filter(l => l.startsWith('NG'));

// 1回目: 作成して確認
g.selfTest();
assert.deepStrictEqual(ngLines(), [], ngLines().join('\n'));
assert.ok(logs.some(l => l.startsWith('== selfTest の結果: OK 16 件 / NG 0 件')), logs.join('\n'));
assert.strictEqual(created.ss, 1);
assert.strictEqual(created.form, 1);
assert.strictEqual(published, true);
assert.strictEqual(triggers.length, 2);
assert.strictEqual(sent.length, 2);          // 自動返信 + 毎朝の一覧
assert.strictEqual(sent[0].to, 'owner@example.com');
assert.ok(!logs.join('\n').includes('owner@example.com'), 'ログにメールアドレスを出さない');

// フォーム送信前の checkFormSubmit は NG を出す
logs.length = 0;
g.checkFormSubmit();
assert.ok(logs.some(l => l.startsWith('NG  フォームに新しい回答が届いている')));

// 人がフォームから送った（トリガーが動いた）ことにする
const later = Number(props.TEST_AT) + 1000;
const ir = (t, v) => ({ getItem: () => ({ getTitle: () => t }), getResponse: () => v });
const manual = { getItemResponses: () => [ir('お名前', 'テスト'), ir('メールアドレス', 'owner@example.com'),
  ir('ご用件', 'その他'), ir('お問い合わせ内容', '手で送ったテスト')], getTimestamp: () => new Date(later) };
responses.push(manual);
const RealDate = Date;
global.Date = class extends RealDate { constructor(...a) { a.length ? super(...a) : super(later); } static now() { return later; } };
g.onFormSubmit({ response: manual });
global.Date = RealDate;
logs.length = 0;
g.checkFormSubmit();
assert.deepStrictEqual(ngLines(), [], logs.join('\n'));
assert.ok(logs.some(l => l === '== checkFormSubmit の結果: OK 3 件 / NG 0 件 =='), logs.join('\n'));

// 2回目の selfTest はフォームとシートを作り直さない
logs.length = 0;
g.selfTest();
assert.strictEqual(created.ss, 1);
assert.strictEqual(created.form, 1);
assert.strictEqual(triggers.length, 2);
assert.deepStrictEqual(ngLines(), []);

// resetSample で ID とトリガーが消える
g.resetSample();
assert.strictEqual(triggers.length, 0);
assert.strictEqual(props.SHEET_ID, undefined);

out('selfTest / checkFormSubmit の流れのテストに合格しました');
