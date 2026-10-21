// Приёмник результатов опросника «Агрессия в переговорах»
var SHEET_ID = "1dGT9WEiHR39tKWSyS1l0KjufxoPuWfzzt3GU7zE6sBc";
var SCALES = [
  ["nap","Напористость","c"],["neu","Неуступчивость","c"],["bes","Бескомпромиссность","c"],["pod","Подозрительность","c"],
  ["net","Нетерпимость","d"],["mst","Мстительность","d"],["vsp","Вспыльчивость","d"],["obd","Обидчивость","d"]
];
var RANGE = {c:[15,25], d:[5,15]};

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);
    var sh = answersSheet_();
    var byId = {};
    (data.scores || []).forEach(function (s) { byId[s.id] = Number(s.score); });
    var row = [new Date(), String(data.group || "без метки").slice(0, 80)];
    SCALES.forEach(function (s) { row.push(byId[s[0]]); });
    var ans = (data.answers || []).slice(0, 80).map(Number);
    sh.appendRow(row.concat(ans));
    buildSummary_();
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return ContentService.createTextOutput("Агрессия в переговорах: приёмник работает"); }

function answersSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName("Ответы") || ss.getSheets()[0];
  if (sh.getName() !== "Ответы") sh.setName("Ответы");
  if (sh.getRange(1, 11).getValue() !== "В1") {
    var h = ["Дата", "Группа"].concat(SCALES.map(function (s) { return s[1]; }));
    for (var i = 1; i <= 80; i++) h.push("В" + i);
    if (sh.getMaxColumns() < h.length) sh.insertColumnsAfter(sh.getMaxColumns(), h.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function buildSummary_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var src = answersSheet_().getDataRange().getValues().slice(1);
  var sh = ss.getSheetByName("Сводка") || ss.insertSheet("Сводка", 0);
  sh.clear();
  var groups = {}, order = [];
  src.forEach(function (r) {
    var g = r[1]; if (!groups[g]) { groups[g] = { n: 0, last: r[0], rows: [] }; order.push(g); }
    groups[g].n++; groups[g].rows.push(r); if (r[0] > groups[g].last) groups[g].last = r[0];
  });
  order.sort(function (a, b) { return groups[b].last - groups[a].last; });

  var head = ["Группа", "Участников", "Последний ответ"];
  SCALES.forEach(function (s) { head.push(s[1] + " — среднее", s[1] + " — вне нормы"); });
  var out = [head];
  order.forEach(function (g) {
    var G = groups[g], line = [g, G.n, G.last];
    SCALES.forEach(function (s, k) {
      var vals = G.rows.map(function (r) { return Number(r[2 + k]); });
      var avg = vals.reduce(function (a, b) { return a + b; }, 0) / vals.length;
      var R = RANGE[s[2]];
      var outN = vals.filter(function (v) { return s[2] === "c" ? (v < R[0] || v > R[1]) : v > R[1]; }).length;
      line.push(Math.round(avg * 10) / 10, outN + " из " + G.n);
    });
    out.push(line);
  });
  sh.getRange(1, 1, out.length, head.length).setValues(out);
  sh.getRange(1, 1, 1, head.length).setFontWeight("bold").setWrap(true).setVerticalAlignment("top");
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  sh.getRange(1, 1, 1, 1).setNote("Норма: конструктивная агрессия 15–25, деструктивная 5–15. Красным — среднее группы вне нормы.");
  // подсветка средних вне нормы
  for (var i = 1; i < out.length; i++) {
    SCALES.forEach(function (s, k) {
      var v = out[i][3 + k * 2], R = RANGE[s[2]];
      var bad = s[2] === "c" ? (v < R[0] || v > R[1]) : v > R[1];
      sh.getRange(i + 1, 4 + k * 2).setFontColor(bad ? "#d0301c" : "#1a7f1a").setFontWeight("bold");
    });
  }
  if (out.length > 1) sh.getRange(2, 3, out.length - 1, 1).setNumberFormat("dd.MM.yyyy HH:mm");
}

// Запустить один раз вручную, чтобы выдать доступ и подготовить листы
function setup() { answersSheet_(); buildSummary_(); }
