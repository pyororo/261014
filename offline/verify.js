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
// 추가조건 자체 시험: 엔진 판정과 직접 계산한 건수가 같은지, 잘못된 행은 오류로 걸러지는지
const cp = ENGINE.parseCustom([
  { 조건ID: "U90", 조건명: "백화점 10만원 이상", 컬럼: "가맹점업종", 비교: "같음", 값: "백화점", 가중치: 2 },
  { 조건ID: "U90", 컬럼: "전표금액", 비교: "이상", 값: 100000 },
  { 조건ID: "U91", 조건명: "주말", 컬럼: "요일", 비교: "목록중하나", 값: "토,일", 가중치: 1 },
  { 조건ID: "U92", 조건명: "적요 공란", 컬럼: "전표 품목 텍스트", 비교: "비어있음", 값: "", 가중치: 1 },
  { 조건ID: "U93", 조건명: "오류 확인용", 컬럼: "없는컬럼", 비교: "같음", 값: "x" },
  { 조건ID: "X1", 조건명: "오류 확인용", 컬럼: "전표금액", 비교: "같음", 값: "1" },
], Object.keys(tx[0]));
const cref = ENGINE.buildRef({ accounts: sheet("계정코드"), industries: sheet("업종코드"), holidays: sheet("공휴일"), rules: sheet("점검기준"), employees: sheet("직원마스터"), vacations: sheet("휴가내역") });
cp.rules.forEach((c) => (cref.weights[c.id] = c.weight));
const crow = ENGINE.applyRules(ENGINE.prepare(tx, cref).rows, cref, ENGINE.DEFAULT_PARAMS, cp.rules);
const got = (id) => crow.filter((r) => r.hits.has(id)).length;
const want = { U90: crow.filter((r) => r.industry === "백화점" && r.amount >= 100000).length, U91: crow.filter((r) => r.wd === 0 || r.wd === 6).length, U92: crow.filter((r) => !r.memo).length };
const sameScore = crow.filter((r) => r.hits.has("U90") && !r.adjust).every((r) => r.ruleScore === [...r.hits].reduce((s, h) => s + (cref.weights[h] || 0), 0));
checks.push(["추가조건 판정", Object.keys(want).map((k) => `${k} ${got(k)}/${want[k]}`).join(", ") + `, 오류행 ${cp.errors.length}/2`, Object.keys(want).every((k) => got(k) === want[k]) && cp.errors.length === 2 && cp.rules.length === 3 && sameScore]);

// 규칙 수정 화면 설정 시험: 사용 안 함은 0건, 세부 설정을 넓히면 같거나 늘어남, 기본값은 결과 그대로
const base = (id, rs) => rs.filter((r) => r.hits.has(id)).length;
const runP = (extra) => ENGINE.applyRules(ENGINE.prepare(tx, ref).rows, ref, { ...ENGINE.DEFAULT_PARAMS, ...extra });
const d0 = runP({});
const offN = base("R01", runP({ off: ["R01"] }));
const bareN = base("R06", runP({ bareWords: ["회의"] }));
const splitN = base("R04", runP({ splitDays: 3 }));
const dupN = base("R09", runP({ dupKeys: ["카드", "가맹점", "금액"] }));
const nightN = base("R02", runP({ nightExclude: ["택시"] }));
checks.push(["규칙 수정 설정", `R01 끔 ${offN}건, R06 ${base("R06", d0)}→${bareN}, R04 ${base("R04", d0)}→${splitN}, R09 ${base("R09", d0)}→${dupN}, R02 택시 제외 ${base("R02", d0)}→${nightN}`,
  offN === 0 && bareN >= base("R06", d0) && splitN >= base("R04", d0) && dupN >= base("R09", d0) && nightN <= base("R02", d0) && d0.length === rows.length && d0.every((r, i) => r.score === rows[i].score)]);

// 기준표 합치기 시험: 가상데이터와 기준표 양식(예시 행 포함)을 어떤 순서로 불러와도 결과가 같은지 (v1.6 S16 미탐 재발 방지)
const TPL = new Function(require("fs").readFileSync(require("path").join(__dirname, "src/ref.js"), "utf8") + "; return DEFAULT_REF;")();
const KEYS = { accounts: "계정코드", industries: "업종코드", holidays: "공휴일", rules: "점검기준", employees: "직원마스터", vacations: "휴가내역" };
const dataRef = Object.fromEntries(Object.entries(KEYS).map(([k, n]) => [k, sheet(n)]));
const loadOrder = (srcs) => { const raw = {}; for (const src of srcs) for (const k of Object.keys(KEYS)) if (src[k]) raw[k] = ENGINE.mergeRef(k, raw[k] || [], src[k]).rows; return raw; };
const orderRun = (srcs) => { const r = ENGINE.buildRef(loadOrder(srcs)); const rs = ENGINE.applyRules(ENGINE.prepare(tx, r).rows, r, ENGINE.DEFAULT_PARAMS); ENGINE.periodRules(rs, r, ENGINE.DEFAULT_PARAMS, month); return { g: ENGINE.grade(rs, sheet("정답지")), sel: ENGINE.selectMonth(rs, month).length, vac: r.vacations.length }; };
const oA = orderRun([dataRef, TPL]), oB = orderRun([TPL, dataRef]);
checks.push(["기준표 불러오는 순서", `가상데이터→양식 단건 ${oA.g.sHit}/${oA.g.S.length}·선별 ${oA.sel}건, 양식→가상데이터 ${oB.g.sHit}/${oB.g.S.length}·선별 ${oB.sel}건, 휴가내역 ${oA.vac}행(예시 행 제외)`,
  [oA, oB].every((o) => o.g.sHit === g.sHit && o.g.tHit === g.tHit && o.sel === sel.length && o.vac === dataRef.vacations.length)]);

for (const [k, v, ok] of checks) console.log(`${ok ? "통과" : "실패"}  ${k}: ${v}`);
for (const o of g.list.filter((o) => o.type !== "FP" && !o.detected)) console.log(`  미탐 ${o.id} ${o.title} — ${o.note}`);
process.exit(checks.every((c) => c[2]) ? 0 : 1);
