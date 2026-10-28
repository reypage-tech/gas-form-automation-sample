// Utilities.formatDate の偽物（テスト用）。指定のタイムゾーンで日時を書式にする。
// 対応する書式: yyyy MM dd HH mm ss（Code.gs が使う 'MM/dd HH:mm' を含む）
global.Utilities = {
  formatDate(date, timeZone, format) {
    const parts = {};
    new Intl.DateTimeFormat('en-US', {
      timeZone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    }).formatToParts(date).forEach(p => { parts[p.type] = p.value; });
    const map = { yyyy: parts.year, MM: parts.month, dd: parts.day, HH: parts.hour, mm: parts.minute, ss: parts.second };
    return format.replace(/yyyy|MM|dd|HH|mm|ss/g, t => map[t]);
  },
};
