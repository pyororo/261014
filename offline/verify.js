// 반입 전 검증: 가상데이터 엑셀로 정답지 탐지율을 확인한다.
// 사용: node offline/verify.js data/법인카드_가상데이터_202509-202608.xlsx
const XLSX = require("./src/vendor/xlsx.full.min.js");
const { ENGINE } = require("./src/engine.js");
const file = process.argv[2];
if (!file) { console.error("검증할 가상데이터 엑셀 경로를 넣어 주세요."); process.exit(2); }
const t0 = Date.now();
const wb = XLSX.read(require("fs").readFileSync(file), { cellDates: true });
const sheet = (name) => { const n = wb.SheetNames.find((s) => s.replace(/\s/g, "") === name || s.startsWith(name)); return n ? XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: "", raw: true }) : []; };
const tx = sheet("전체");
const ref = ENGINE.buildRef({ accounts: sheet("계정코드"), industries: sheet("업종코드"), holidays: sheet("공휴일"), rules: sheet("점검기준"), employees: sheet("직원마스터"), vacations: sheet("휴가내역") });
const { rows: all, stats } = ENGINE.prepare(tx, ref);
const rows = ENGINE.applyRules(all, ref, ENGINE.DEFAULT_PARAMS);
const months = [...new Set(rows.map((r) => r.month))].sort();
const month = process.argv[3] || months[months.length - 1];
const period = ENGINE.periodRules(rows, ref, ENGINE.DEFAULT_PARAMS, month);
const g = ENGINE.grade(rows, sheet("정답지"));
const sel = ENGINE.selectMonth(rows, month);
const cnt = { 상: 0, 중: 0, 하: 0 }; sel.forEach((r) => cnt[r.grade]++);
const checks = [
  ["데이터 읽기", `${all.length}/${tx.length}건`, all.length === tx.length],
  ["처리시간", `${((Date.now() - t0) / 1000).toFixed(2)}초`, Date.now() - t0 < 60000],
  ["단건 이상거래", `${g.sHit}/${g.S.length}종`, g.S.length > 0 && g.sHit === g.S.length],
  ["기간 패턴", `${g.tHit}/${g.T.length}종`, g.T.length > 0 && g.tHit === g.T.length],
  ["정상거래 '상' 오분류", `${g.fpHigh}/${g.fpN}건`, g.fpN > 0 && g.fpHigh === 0],
  ["추가 점검항목 추천", `${ENGINE.recommend(rows, ref, month).length}건`, ENGINE.recommend(rows, ref, month).length >= 3],
];
console.log(`분석 대상 ${rows.length}건 (취소 ${stats.cancels}건 상계) · ${month} 선별 ${sel.length}건 (상 ${cnt.상} · 중 ${cnt.중} · 하 ${cnt.하}) · 기간 패턴 ${Object.values(period).reduce((s, x) => s + x.length, 0)}건`);
for (const [k, v, ok] of checks) console.log(`${ok ? "통과" : "실패"}  ${k}: ${v}`);
for (const o of g.list.filter((o) => o.type !== "FP" && !o.detected)) console.log(`  미탐 ${o.id} ${o.title} — ${o.note}`);
process.exit(checks.every((c) => c[2]) ? 0 : 1);
