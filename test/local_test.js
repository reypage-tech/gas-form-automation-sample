// Google のサービスを簡単な偽物に置き換えて、Code.gs の処理を Node で確かめる。
// 実行: node test/local_test.js
const assert = require('assert');
const fs = require('fs');
const path = require('path');

const sent = [];
const rows = [['受付日時', 'お名前', 'メールアドレス', 'ご用件', 'お問い合わせ内容', '自動返信', '状況', 'メモ']];

global.MailApp = {
  sendEmail: m => {
    if (m.to === 'fail@example.com') throw new Error('Service invoked too many times');
    sent.push(m);
  }
};
global.Session = { getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) };
global.PropertiesService = { getScriptProperties: () => ({ getProperty: () => 'TEST_SHEET' }) };
global.LockService = { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) };
global.SpreadsheetApp = {
  openById: () => ({
    getSheetByName: () => ({
      appendRow: r => rows.push(r),
      getDataRange: () => ({ getValues: () => rows })
    })
  })
};

const mod = { exports: {} };
const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
new Function('module', 'exports', code)(mod, mod.exports);
const g = mod.exports;

const item = (title, value) => ({ getItem: () => ({ getTitle: () => title }), getResponse: () => value });
const submit = (name, email, topic, message) => g.onFormSubmit({
  response: {
    getItemResponses: () => [
      item('お名前', name), item('メールアドレス', email),
      item('ご用件', topic), item('お問い合わせ内容', message)
    ]
  }
});

// メールアドレスの判定
assert.strictEqual(g.isValidEmail_('hanako@example.com'), true);
assert.strictEqual(g.isValidEmail_('hanako@example'), false);
assert.strictEqual(g.isValidEmail_(''), false);

// 正しいアドレス、返信1通、記録は「未対応」
submit('テスト 花子', 'hanako@example.com', '体験レッスン', '土曜の午前は空いていますか');
assert.strictEqual(sent.length, 1);
assert.strictEqual(sent[0].to, 'hanako@example.com');
assert.ok(sent[0].body.startsWith('テスト 花子様'));
assert.ok(sent[0].body.includes('24時間以内'));
assert.strictEqual(rows[1][5], '送信済み');
assert.strictEqual(rows[1][6], '未対応');

// 形式の誤り、返信しないで、メモに理由
submit('テスト 次郎', 'jiro-at-example', '料金・プラン', '');
assert.strictEqual(sent.length, 1);
assert.strictEqual(rows[2][5], 'なし');
assert.ok(rows[2][7].includes('形式'));

// 送信の失敗、記録は残し、メモに理由
submit('テスト 三郎', 'fail@example.com', 'その他', '駐車場はありますか');
assert.strictEqual(rows[3][5], 'なし');
assert.ok(rows[3][7].startsWith('自動返信に失敗'));

// 1件を完了にして、毎朝の一覧を確かめる
rows[1][6] = '完了';
g.sendDailySummary();
const summary = sent[1];
assert.strictEqual(summary.to, 'owner@example.com');
assert.ok(summary.subject.includes('新着3件・未完了2件'));
assert.ok(summary.body.includes('・体験レッスン：1件'));
assert.ok(summary.body.includes('※自動返信なし（メールアドレスの形式が正しくない）'));
assert.ok(!summary.body.includes('テスト 花子'));

// 古い記録だけなら新着0件
const old = [[new Date(Date.now() - 3 * 86400000), '過去', 'a@example.com', 'その他', '', '送信済み', '完了', '']];
const s0 = g.buildSummary_(old, new Date(), 'Asia/Tokyo');
assert.strictEqual(s0.newCount, 0);
assert.strictEqual(s0.openCount, 0);

console.log('すべてのテストに合格しました');
console.log('--- 毎朝の一覧（例） ---');
console.log(summary.subject);
console.log(summary.body);
