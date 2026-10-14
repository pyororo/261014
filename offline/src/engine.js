// ───────────── 법인카드 점검 엔진 (규칙 R01~R17, T-R1~T-R4, 위험점수·등급) ─────────────
const ENGINE = (() => {
  const WD = ["일", "월", "화", "수", "목", "금", "토"];
  const RULE_NAMES = {
    R01: "휴일 사용", R02: "심야 사용", R03: "제한업종", R04: "분할결제 의심", R05: "건당 한도 초과",
    R06: "적요 부실", R07: "계정-업종 불일치", R09: "중복결제", R11: "해외 사용", R12: "라운드 금액",
    R13: "의미 불일치(AI)", R14: "원거리 근접 사용", R15: "정산 지연", R16: "공용카드 사용자 미기재", R17: "휴가 중 사용",
  };
  const RULE_IDS = Object.keys(RULE_NAMES);
  const DEFAULT_WEIGHTS = { R01: 2, R02: 2, R03: 5, R04: 4, R05: 3, R06: 2, R07: 3, R09: 3, R11: 1, R12: 1, R13: 2, R14: 3, R15: 2, R16: 2, R17: 3 };
  const T_NAMES = { "T-R1": "동일가맹점 반복", "T-R2": "사용액 급증", "T-R3": "연말·분기말 집중", "T-R4": "정액 반복" };
  const DEFAULT_PARAMS = {
    nightStart: 23, nightEnd: 6, splitGap: 60, memoMin: 5, settleDays: 60, round: 100000, sharedMin: 300000, distGap: 60,
    t1Count: 8, t2Ratio: 2, t2Floor: 500000, t3Share: 0.5, t3Mult: 1.3, t4Months: 3,
  };
  const TITLES = ["대표", "사장", "회장", "전무", "상무", "이사", "본부장", "지점장", "실장", "부장", "차장", "과장", "팀장", "대리", "원장", "국장", "센터장", "단장", "고객"];
  const COUNTERPART_RE = new RegExp("ㅇㅇㅇ|\\d+\\s*[인명]|" + TITLES.join("|"));
  const PREAPPROVAL_RE = /사전승인|GW-\d{6}-\d+/;
  const TITLE_KEYWORDS = { 하이패스: ["주유소", "주차장"], 통행료: ["주유소", "주차장"], 주유: ["주유소"], 주차: ["주차장"], 택시: ["택시"], 항공: ["항공사"], 숙박: ["숙박"], 도서: ["서적"] };
  const TRAVEL_WORDS = /출장|이동|택시|귀가|방문|외근|교통/;
  const DIST_EXCLUDE = new Set(["온라인쇼핑", "PG일반(인증)", "해외가맹점", "택시", "항공사", "철도", "주유소", "주차장", "고속버스"]);

  // ── 추가조건(사용자 정의 단건 조건): 기준표 '추가조건' 시트, 같은 조건ID 행은 AND ──
  const OPS = ["같음", "다름", "포함", "미포함", "시작", "이상", "이하", "초과", "미만", "목록중하나", "비어있음", "비어있지않음"];
  const DERIVED = {
    요일: (r) => r.wdName, 승인시: (r) => r.hour, 휴일: (r) => (r.holiday ? "Y" : "N"), 승인월: (r) => r.month,
    공휴일명: (r) => r.holName, 시도: (r) => sido(r.addr), 제한구분: (r) => r.restrict,
  };
  let CUSTOM = new Map();
  const ruleName = (id) => RULE_NAMES[id] || CUSTOM.get(id)?.name || id;
  function parseCustom(rows, columns) {
    const known = new Set([...columns, ...Object.keys(DERIVED)]);
    const byId = new Map(), errors = [];
    rows.forEach((x, i) => {
      const id = str(x["조건ID"]), col = str(x["컬럼"]), op = str(x["비교"]).replace(/\s/g, "");
      if (!id && !col) return;
      const where = `추가조건 ${i + 2}행${id ? `(${id})` : ""}`;
      if (!/^U\d+$/i.test(id)) return errors.push(`${where}: 조건ID는 U01처럼 U로 시작하는 번호로 적어 주세요.`);
      if (!known.has(col)) return errors.push(`${where}: '${col}' 컬럼이 SAP 엑셀에 없습니다.`);
      if (!OPS.includes(op)) return errors.push(`${where}: 비교 '${op}'는 쓸 수 없습니다. (${OPS.join(", ")})`);
      const val = x["값"] instanceof Date ? ymd(toDate(x["값"])) : str(x["값"]);
      if (["이상", "이하", "초과", "미만"].includes(op) && (val === "" || !isFinite(+val.replace(/[,\s원]/g, "")))) return errors.push(`${where}: '${op}' 비교에는 숫자 값을 적어 주세요.`);
      const key = id.toUpperCase();
      if (!byId.has(key)) byId.set(key, { id: key, name: "", weight: null, conds: [] });
      const rule = byId.get(key);
      if (str(x["조건명"]) && !rule.name) rule.name = str(x["조건명"]);
      if (str(x["가중치"]) !== "" && rule.weight === null) rule.weight = num(x["가중치"]);
      rule.conds.push({ col, op, val });
    });
    const rules = [...byId.values()].map((r) => ({ ...r, name: r.name || r.id, weight: r.weight ?? 1 }));
    return { rules, errors };
  }
  // 추가조건 비교용 숫자 변환: 숫자가 아니면 NaN (num()은 0을 돌려주므로 쓰지 않음)
  const strictNum = (v) => { const t = String(v ?? "").replace(/[,\s원]/g, ""); return t !== "" && isFinite(+t) ? +t : NaN; };
  function testCond(r, c) {
    const v = c.col in DERIVED ? DERIVED[c.col](r) : r.raw[c.col];
    const sv = v instanceof Date ? ymd(toDate(v)) : str(v);
    switch (c.op) {
      case "같음": return sv === c.val || (!isNaN(strictNum(sv)) && strictNum(sv) === strictNum(c.val));
      case "다름": return !testCond(r, { ...c, op: "같음" });
      case "포함": return sv.includes(c.val);
      case "미포함": return !sv.includes(c.val);
      case "시작": return sv.startsWith(c.val);
      case "이상": return strictNum(sv) >= strictNum(c.val);
      case "이하": return strictNum(sv) <= strictNum(c.val);
      case "초과": return strictNum(sv) > strictNum(c.val);
      case "미만": return strictNum(sv) < strictNum(c.val);
      case "목록중하나": return c.val.split(",").map((x) => x.trim()).includes(sv);
      case "비어있음": return sv === "";
      case "비어있지않음": return sv !== "";
    }
    return false;
  }

  const pad = (n, w = 2) => String(n).padStart(w, "0");
  const ymd = (d) => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const parseYmd8 = (s) => { s = String(s).replace(/\D/g, ""); return new Date(Date.UTC(+s.slice(0, 4), +s.slice(4, 6) - 1, +s.slice(6, 8))); };
  function toDate(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    if (typeof v === "number") {
      if (v > 19000000) return parseYmd8(v);
      return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000); // 엑셀 일련번호
    }
    const s = String(v).trim();
    if (/^\d{8}$/.test(s)) return parseYmd8(s);
    const m = s.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
    return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
  }
  const str = (v) => (v == null ? "" : String(v).trim());
  const num = (v) => { const n = typeof v === "number" ? v : Number(String(v ?? "").replace(/[,\s원]/g, "")); return isFinite(n) ? n : 0; };
  const sido = (addr) => str(addr).split(/\s+/)[0] || "";
  const area2 = (addr) => str(addr).split(/\s+/).slice(0, 2).join(" ");
  const monthKey = (d) => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}`;
  const won = (x) => Math.round(x).toLocaleString("ko-KR") + "원";
  const nf = (x) => Math.round(x).toLocaleString("ko-KR");

  // ── 기준표 정리 ──
  function buildRef(raw) {
    const accounts = new Map();
    for (const a of raw.accounts) {
      const code = str(a["G/L 계정"]);
      if (!code) continue;
      accounts.set(code, {
        name: str(a["G/L 계정명"]),
        allowed: new Set(str(a["허용 가맹점업종"] ?? a["허용업종"]).split(",").map((x) => x.trim()).filter(Boolean)),
        limit: num(a["건당 한도(원)"] ?? a["건당 한도"] ?? a["건당한도"]) || null,
      });
    }
    const industries = new Map();
    for (const i of raw.industries) industries.set(str(i["업종코드"]), { name: str(i["가맹점업종"]), restrict: str(i["제한구분"]) || "일반" });
    const industryByName = new Map([...industries.values()].map((v) => [v.name, v]));
    const holidays = new Map();
    for (const h of raw.holidays) { const d = toDate(h["일자"]); if (d) holidays.set(ymd(d), str(h["명칭"])); }
    const weights = { ...DEFAULT_WEIGHTS };
    for (const r of raw.rules || []) { const id = str(r["규칙ID"]); const w = num(r["위험가중치"]); if (id in weights && w) weights[id] = w; }
    const empByName = new Map(), empById = new Map();
    for (const e of raw.employees || []) {
      const rec = { id: str(e["사번"]), name: str(e["이름"]), dept: str(e["부서"]), work: str(e["근무지"]) };
      empById.set(rec.id, rec); if (rec.name) empByName.set(rec.name, rec);
    }
    const vacations = [];
    for (const v of raw.vacations || []) vacations.push({ id: str(v["사번"]), name: str(v["이름"]), kind: str(v["휴가구분"]), from: toDate(v["시작일"]), to: toDate(v["종료일"]) });
    return { accounts, industries, industryByName, holidays, weights, empByName, empById, vacations };
  }

  // ── 전처리: 텍스트 컬럼, 승인일시, 휴일, 취소 상계 ──
  function prepare(rawRows, ref) {
    const seen = new Set();
    const rows = [];
    let dup = 0;
    for (const r of rawRows) {
      const no = str(r["승인번호"]);
      if (!no) continue;
      if (seen.has(no)) { dup++; continue; }
      seen.add(no);
      const d = toDate(r["승인일"]);
      if (!d) continue;
      const t = str(r["승인시간"]).replace(/\D/g, "").padStart(6, "0");
      const hh = +t.slice(0, 2), mi = +t.slice(2, 4), ss = +t.slice(4, 6);
      const dt = d.getTime() + ((hh * 60 + mi) * 60 + ss) * 1000;
      const key = ymd(d);
      const wd = d.getUTCDay();
      const holName = ref.holidays.get(key) || "";
      const ind = ref.industries.get(str(r["업종코드"])) || ref.industryByName.get(str(r["가맹점업종"]));
      const acc = ref.accounts.get(str(r["G/L 계정"]));
      const ref참조 = str(r["참조"]);
      const sabun = ref참조.includes("/") ? ref참조.split("/")[0].trim() : "";
      const emp = ref.empById.get(sabun) || ref.empByName.get(str(r["직원명"])) || null;
      rows.push({
        no, card: str(r["카드번호"]), kind: str(r["구분"]), emp: str(r["직원명"]), empId: emp ? emp.id : sabun, work: emp ? emp.work : "",
        dept: str(r["예산부서명"]), deptCode: str(r["예산부서"]), date: d, dkey: key, month: monthKey(d), time: `${pad(hh)}:${pad(mi)}`, hour: hh, dt,
        wd, wdName: WD[wd], holiday: wd === 0 || wd === 6 || !!holName, holName,
        gl: str(r["G/L 계정"]), glName: str(r["G/L 계정명"]) || (acc ? acc.name : ""), industry: str(r["가맹점업종"]) || (ind ? ind.name : ""), indCode: str(r["업종코드"]),
        merchant: str(r["가맹점명"]), addr: str(r["가맹점주소"]), amount: num(r["전표금액"]), cancel: str(r["취소구분"]) === "취소",
        title: str(r["결재문서제목"]), memo: str(r["전표 품목 텍스트"]), overseas: str(r["국내해외코드명"]) === "해외", currency: str(r["통화"]),
        raw: r, acctDate: toDate(r["회계일자"]), restrict: ind ? ind.restrict : "일반", limit: acc ? acc.limit : null, allowed: acc ? acc.allowed : null,
      });
    }
    // 취소 상계: 같은 카드·가맹점·절대금액, 승인일시가 가장 가까운 미매칭 매입
    const purchases = rows.filter((r) => !r.cancel);
    const used = new Set();
    let matched = 0, unmatched = 0;
    for (const c of rows.filter((r) => r.cancel)) {
      let best = null;
      for (const p of purchases) {
        if (used.has(p) || p.card !== c.card || p.merchant !== c.merchant || p.amount !== -c.amount) continue;
        if (!best || Math.abs(p.dt - c.dt) < Math.abs(best.dt - c.dt)) best = p;
      }
      if (best) { used.add(best); best.offset = c.no; c.offset = best.no; matched++; } else unmatched++;
      c.excluded = true;
    }
    for (const p of used) p.excluded = true;
    return { rows, stats: { dup, cancels: matched + unmatched, matched, unmatched } };
  }

  // ── 단건·다건 규칙 ──
  function applyRules(all, ref, P, custom = []) {
    CUSTOM = new Map(custom.map((c) => [c.id, c]));
    const rows = all.filter((r) => !r.excluded);
    for (const r of rows) r.hits = new Set();
    const hit = (r, id) => r.hits.add(id);
    const byKey = (keyFn) => { const m = new Map(); for (const r of rows) { const k = keyFn(r); if (!m.has(k)) m.set(k, []); m.get(k).push(r); } return m; };

    for (const r of rows) {
      if (r.holiday) hit(r, "R01");
      if (r.hour >= P.nightStart || r.hour < P.nightEnd) hit(r, "R02");
      if (r.restrict === "금지" || r.restrict === "사전승인") hit(r, "R03");
      if (r.limit && r.amount > r.limit) hit(r, "R05");
      const memo = r.memo;
      if (memo.length < P.memoMin || (r.glName.includes("접대비") && !COUNTERPART_RE.test(memo))) hit(r, "R06");
      if (r.allowed && r.allowed.size && !r.allowed.has(r.industry)) hit(r, "R07");
      if (r.overseas) hit(r, "R11");
      if (r.amount >= P.round && r.amount % P.round === 0) hit(r, "R12");
      if (r.acctDate && (r.acctDate - r.date) / 86400000 > P.settleDays) hit(r, "R15");
      if (r.kind.includes("공용") && !r.emp && r.amount >= P.sharedMin) hit(r, "R16");
      const r13 = r13Check(r);
      if (r13) { hit(r, "R13"); r.r13Reason = r13; }
      if (!r.kind.includes("공용") && (r.empId || r.emp)) { // 휴가 중 본인 카드 사용
        const v = ref.vacations.find((v) => (v.id ? v.id === r.empId : v.name === r.emp) && v.from && v.to && r.date >= v.from && r.date <= v.to);
        if (v) { hit(r, "R17"); r.vacation = v; }
      }
    }
    // R04 분할결제
    for (const g of byKey((r) => `${r.card}|${r.merchant}|${r.dkey}`).values()) {
      if (g.length < 2) continue;
      g.sort((a, b) => a.dt - b.dt);
      let cluster = [g[0]];
      const flush = () => {
        const lim = cluster[0].limit;
        if (cluster.length >= 2 && lim && cluster.reduce((s, x) => s + x.amount, 0) > lim) cluster.forEach((x) => { hit(x, "R04"); x.splitSum = cluster.reduce((s, y) => s + y.amount, 0); x.splitN = cluster.length; });
      };
      for (let i = 1; i < g.length; i++) {
        if (g[i].dt - g[i - 1].dt > P.splitGap * 60000) { flush(); cluster = []; }
        cluster.push(g[i]);
      }
      flush();
    }
    // R09 중복결제(상계 후 남은 건끼리)
    for (const g of byKey((r) => `${r.card}|${r.merchant}|${r.amount}|${r.dkey}`).values()) if (g.length >= 2) g.forEach((x) => hit(x, "R09"));
    // R14 원거리 근접 사용
    for (const g of byKey((r) => r.card).values()) {
      g.sort((a, b) => a.dt - b.dt);
      for (let i = 1; i < g.length; i++) {
        const a = g[i - 1], b = g[i];
        if (DIST_EXCLUDE.has(a.industry) || DIST_EXCLUDE.has(b.industry)) continue;
        const sa = sido(a.addr), sb = sido(b.addr);
        if (sa && sb && sa !== sb && b.dt - a.dt <= P.distGap * 60000) {
          const gap = Math.round((b.dt - a.dt) / 60000);
          hit(a, "R14"); hit(b, "R14");
          a.distNote = b.distNote = `${sa} ${a.time} → ${sb} ${b.time} (${gap}분 간격)`;
        }
      }
    }
    // 위험점수·등급 + AI 오탐 조정
    for (const c of custom) for (const r of rows) if (c.conds.every((x) => testCond(r, x))) hit(r, c.id);
    const W = ref.weights;
    for (const r of rows) {
      r.ruleScore = [...r.hits].reduce((s, id) => s + (W[id] || 0), 0);
      const soft = ["R01", "R02"].filter((id) => r.hits.has(id));
      r.adjust = 0;
      if (r.hits.size && soft.length === r.hits.size && PREAPPROVAL_RE.test(r.memo)) {
        r.adjust = -soft.length;
        r.adjustReason = "적요에 사전승인 문서번호가 있어 소명 가능 → 저위험으로 조정";
      }
      r.score = r.ruleScore ? Math.max(r.ruleScore + r.adjust, 1) : 0;
      r.banned = r.hits.has("R03") && r.restrict === "금지";
      r.grade = r.banned || r.score >= 6 ? "상" : r.score >= 3 ? "중" : r.score >= 1 ? "하" : "";
      r.reason = reasonSentence(r);
      r.checks = checkItems(r);
    }
    return rows;

    function r13Check(r) {
      for (const [kw, ok] of Object.entries(TITLE_KEYWORDS)) {
        if (r.title.includes(kw) && !ok.includes(r.industry)) return `결재문서제목('${r.title}')과 실제 업종(${r.industry})·적요('${r.memo}')가 서로 다름`;
      }
      if (r.glName.includes("여비교통비") && (r.industry === "일반음식점" || r.industry === "일식")) {
        const same = r.work && area2(r.addr) === r.work;
        if (r.memo.includes("회식") || (same && !TRAVEL_WORDS.test(r.memo))) {
          return `${same ? `근무지(${r.work}) 안` : "출장 표시 없이"} 식사('${r.memo}')를 여비교통비로 처리 — 복리후생비 등 계정 오분류 의심`;
        }
      }
      return "";
    }
  }

  function phrases(r) {
    const p = [];
    const h = r.hits;
    if (h.has("R01")) p.push(`휴일(${r.holName || r.wdName + "요일"}) 사용`);
    if (h.has("R02")) p.push(`심야 ${r.time} 승인`);
    if (h.has("R03")) p.push(`${r.industry}은(는) ${r.restrict} 업종`);
    if (h.has("R04")) p.push(`같은 날 ${r.splitN}회로 나눠 결제해 합계 ${won(r.splitSum)}가 한도 ${won(r.limit)} 초과`);
    if (h.has("R05")) p.push(`건당 한도 ${won(r.limit)} 초과`);
    if (h.has("R06")) p.push(!r.memo ? "적요 공란" : r.memo.length < 5 ? `적요 부실('${r.memo}')` : `접대 상대방 미기재('${r.memo}')`);
    if (h.has("R07")) p.push(`${r.glName} 계정에 허용되지 않은 업종(${r.industry})`);
    if (h.has("R09")) p.push("같은 날 같은 금액이 취소 없이 2번 이상 승인");
    if (h.has("R11")) p.push(`해외 결제(${r.currency || "외화"})`);
    if (h.has("R12")) p.push(`정액 ${won(r.amount)} 결제`);
    if (h.has("R13")) p.push(r.r13Reason);
    if (h.has("R14")) p.push(`원거리 근접 사용: ${r.distNote}`);
    if (h.has("R15")) p.push(`승인 후 ${Math.round((r.acctDate - r.date) / 86400000)}일 지나 정산`);
    if (h.has("R16")) p.push("공용카드인데 실제 사용자 미기재");
    for (const id of h) if (CUSTOM.has(id)) p.push(`추가조건 '${CUSTOM.get(id).name}' 해당`);
    if (h.has("R17")) p.push(`${r.vacation.kind || "휴가"} 기간(${ymd(r.vacation.from)}~${ymd(r.vacation.to)}) 중 사용`);
    return p;
  }
  function reasonParts(r) {
    return phrases(r).concat(r.adjustReason ? [r.adjustReason] : []);
  }
  function reasonShort(r) {
    const p = phrases(r);
    return p.join(", ") + (r.adjustReason ? ` (${r.adjustReason})` : "");
  }
  function reasonSentence(r) {
    const p = phrases(r);
    if (!p.length) return "";
    let s = `${r.emp || "사용자 미기재(공용카드)"}의 ${r.merchant}(${r.industry}) ${won(r.amount)} 결제: ${p.join(", ")}.`;
    if (r.adjustReason) s += ` 다만 ${r.adjustReason}.`;
    return s;
  }
  const CHECK = {
    R01: "휴일 사용 사유와 사전승인 문서번호 확인", R02: "심야 사용 사유와 실제 사용자 확인", R04: "나눠 결제한 사유, 실제 이용 인원·영수증 확인",
    R05: "한도 초과 사유와 사전 품의 여부 확인", R07: "실제 사용 목적에 맞는 계정으로 재분류 검토", R09: "카드사 매출전표로 이중 결제 확인, 취소 누락이면 환수",
    R11: "해외 출장 품의서와 사용 내역 대조", R12: "정액 결제 사유(선결제·상품권성 여부) 확인", R13: "결재문서·영수증으로 실제 사용 내역 대조",
    R14: "카드 실물 소지자 확인, 카드 대여·복제 여부 점검", R15: "정산 지연 사유와 증빙 보관 여부 확인", R16: "공용카드 대장으로 실제 사용자 확인 후 전표 보완",
    R17: "휴가 중 사용 경위 확인(대리 사용·사적 사용 여부)",
  };
  function checkItems(r) {
    const out = [];
    for (const id of ["R03", "R17", "R14", "R04", "R09", "R01", "R02", "R05", "R06", "R07", "R13", "R16", "R11", "R12", "R15"]) {
      if (!r.hits.has(id)) continue;
      if (id === "R03") out.push(r.restrict === "금지" ? "금지업종 사용 경위 소명 및 환수 검토" : "사전 품의(결재) 문서 확인");
      else if (id === "R06") out.push(r.glName.includes("접대비") ? "접대 상대방(소속·직함·인원)과 목적 확인" : "적요 보완 요청(사용 목적·참석자)");
      else out.push(CHECK[id]);
    }
    for (const id of r.hits) if (CUSTOM.has(id)) out.push(`추가조건 '${CUSTOM.get(id).name}' 해당 사유 확인`);
    return [...new Set(out)];
  }

  // ── 기간(시계열) 규칙 ──
  function addMonths(m, k) { let [y, mo] = m.split("-").map(Number); mo += k; while (mo < 1) { mo += 12; y--; } while (mo > 12) { mo -= 12; y++; } return `${y}-${pad(mo)}`; }
  function businessDays(month, ref) {
    const [y, m] = month.split("-").map(Number);
    const out = [];
    for (let d = new Date(Date.UTC(y, m - 1, 1)); d.getUTCMonth() === m - 1; d = new Date(d.getTime() + 86400000)) {
      const wd = d.getUTCDay();
      if (wd !== 0 && wd !== 6 && !ref.holidays.has(ymd(d))) out.push(ymd(d));
    }
    return out;
  }
  function periodRules(rows, ref, P, endMonth) {
    const result = { "T-R1": [], "T-R2": [], "T-R3": [], "T-R4": [] };
    const months = [...new Set(rows.map((r) => r.month))].sort();
    for (const r of rows) r.tHits = new Set();
    const win3 = [addMonths(endMonth, -2), addMonths(endMonth, -1), endMonth];
    const person = (r) => r.emp || r.card;

    // T-R1 동일가맹점 반복 (직원×가맹점, 최근 3개월 ≥ N회)
    const g1 = new Map();
    for (const r of rows) if (win3.includes(r.month)) { const k = `${person(r)}|${r.merchant}`; if (!g1.has(k)) g1.set(k, []); g1.get(k).push(r); }
    for (const [k, g] of g1) {
      if (g.length < P.t1Count) continue;
      const wdCount = {}; g.forEach((r) => (wdCount[r.wdName] = (wdCount[r.wdName] || 0) + 1));
      const [topWd, topN] = Object.entries(wdCount).sort((a, b) => b[1] - a[1])[0];
      g.forEach((r) => r.tHits.add("T-R1"));
      result["T-R1"].push({ who: person(g[0]), dept: g[0].dept, merchant: g[0].merchant, n: g.length, amount: sum(g), focus: topN / g.length >= 0.6 ? `${topWd}요일 ${topN}회 집중` : "", rows: g });
    }
    // T-R2 사용액 급증 (월 사용액 ÷ 직전 6개월 평균)
    const byEmp = groupBy(rows.filter((r) => r.emp), (r) => r.emp);
    for (const [emp, g] of byEmp) {
      const mAmt = {}; g.forEach((r) => (mAmt[r.month] = (mAmt[r.month] || 0) + r.amount));
      let best = null;
      for (const m of win3) {
        const prev = [1, 2, 3, 4, 5, 6].map((k) => addMonths(m, -k)).filter((x) => months.includes(x));
        if (prev.length < 3) continue;
        const avg = prev.reduce((s, x) => s + (mAmt[x] || 0), 0) / prev.length;
        const cur = mAmt[m] || 0;
        if (avg > 0 && cur >= P.t2Floor && cur / avg > P.t2Ratio && (!best || cur / avg > best.ratio)) best = { month: m, cur, avg, ratio: cur / avg };
      }
      if (best) {
        const hitRows = g.filter((r) => win3.includes(r.month) && r.month >= best.month);
        hitRows.forEach((r) => r.tHits.add("T-R2"));
        result["T-R2"].push({ who: emp, dept: g[0].dept, ...best, rows: hitRows, series: months.map((m) => mAmt[m] || 0) });
      }
    }
    // T-R3 분기말 집중 (부서, 월말 5영업일 비중 > 50% AND 월 사용액 ≥ 부서 월평균 × 1.3)
    const byDept = groupBy(rows, (r) => r.dept);
    for (const [dept, g] of byDept) {
      const mAmt = {}; g.forEach((r) => (mAmt[r.month] = (mAmt[r.month] || 0) + r.amount));
      const avg = months.reduce((s, m) => s + (mAmt[m] || 0), 0) / months.length;
      for (const m of months) {
        if (![3, 6, 9, 12].includes(+m.slice(5))) continue;
        const last5 = new Set(businessDays(m, ref).slice(-5));
        const inM = g.filter((r) => r.month === m);
        const tail = inM.filter((r) => last5.has(r.dkey));
        const share = sum(tail) / (mAmt[m] || 1);
        if (share > P.t3Share && mAmt[m] >= avg * P.t3Mult) {
          tail.forEach((r) => r.tHits.add("T-R3"));
          result["T-R3"].push({ who: dept, month: m, share, total: mAmt[m], avg, rows: tail });
        }
      }
    }
    // T-R4 정액 반복 (동일 직원·가맹점·금액, N개월 이상 연속)
    const g4 = groupBy(rows, (r) => `${person(r)}|${r.merchant}|${r.amount}`);
    for (const [, g] of g4) {
      const ms = [...new Set(g.map((r) => r.month))].sort();
      let run = 1, bestRun = 1;
      for (let i = 1; i < ms.length; i++) { run = ms[i] === addMonths(ms[i - 1], 1) ? run + 1 : 1; bestRun = Math.max(bestRun, run); }
      if (bestRun >= P.t4Months) {
        g.forEach((r) => r.tHits.add("T-R4"));
        result["T-R4"].push({ who: person(g[0]), dept: g[0].dept, merchant: g[0].merchant, amount: g[0].amount, months: bestRun, rows: g });
      }
    }
    return result;
  }
  const sum = (g) => g.reduce((s, r) => s + r.amount, 0);
  function groupBy(arr, fn) { const m = new Map(); for (const x of arr) { const k = fn(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); } return m; }

  function selectMonth(rows, month) {
    const order = { 상: 0, 중: 1, 하: 2 };
    return rows
      .filter((r) => r.score > 0 && (r.month === month || (r.hits.has("R15") && r.acctDate && monthKey(r.acctDate) === month)))
      .sort((a, b) => order[a.grade] - order[b.grade] || b.score - a.score || a.dt - b.dt);
  }

  // ── 정답지 대조 ──
  function grade(rows, answers) {
    const byNo = new Map(rows.map((r) => [r.no, r]));
    const groups = groupBy(answers.filter((a) => str(a["시나리오ID"])), (a) => str(a["시나리오ID"]));
    const out = [];
    for (const [id, g] of groups) {
      const type = id.startsWith("T") ? "T" : id.startsWith("FP") ? "FP" : "S";
      const items = g.map((a) => ({ no: str(a["승인번호"]), expect: str(a["기대 탐지규칙"] ?? a["기대탐지규칙"]), row: byNo.get(str(a["승인번호"])) }));
      const expectIds = [...new Set(items.flatMap((i) => i.expect.match(/T-R\d|R\d{2}/g) || []))];
      let detected, note;
      if (type === "T") {
        const tIds = expectIds.filter((e) => e.startsWith("T-"));
        const tHit = items.filter((i) => i.row && tIds.some((e) => i.row.tHits.has(e)));
        detected = tHit.length > 0;
        note = detected ? `${nf(tHit.length)}/${nf(items.length)}건에 ${tIds.join("·")} 표시` : "기간 규칙 미탐";
      } else if (type === "FP") {
        detected = null;
      } else {
        const ok = items.filter((i) => i.row && i.row.score > 0);
        detected = ok.length === items.length;
        const matched = expectIds.filter((e) => items.some((i) => i.row && i.row.hits.has(e)));
        note = detected ? (matched.length ? `기대규칙 ${matched.join("·")} 적중` : "선별됨(다른 규칙)") : `${nf(items.length - ok.length)}건 미선별`;
      }
      out.push({ id, type, title: str(g[0]["시나리오"]), expect: expectIds.join(", "), items, detected, note });
    }
    const S = out.filter((o) => o.type === "S"), T = out.filter((o) => o.type === "T");
    const fpRows = out.filter((o) => o.type === "FP").flatMap((o) => o.items).filter((i) => i.row);
    return {
      list: out, S, T, sHit: S.filter((o) => o.detected).length, tHit: T.filter((o) => o.detected).length,
      fpN: fpRows.length, fpHigh: fpRows.filter((i) => i.row.grade === "상").length, fpGrades: countBy(fpRows.map((i) => i.row.grade || "미선별")),
    };
  }
  function countBy(arr) { const o = {}; arr.forEach((x) => (o[x] = (o[x] || 0) + 1)); return o; }

  // ── AI 추가 점검항목 추천 (기존 규칙에 없는 패턴, 집계 근거 포함) ──
  function recommend(rows, ref, endMonth) {
    const win3 = [addMonths(endMonth, -2), addMonths(endMonth, -1), endMonth];
    const recent = rows.filter((r) => win3.includes(r.month));
    const recs = [];
    // 1) 한도 바로 밑 결제 집중
    const near = recent.filter((r) => r.limit && r.amount >= r.limit * 0.9 && r.amount <= r.limit);
    const nearBy = [...groupBy(near, (r) => r.emp || r.card)].sort((a, b) => b[1].length - a[1].length);
    recs.push({
      title: "건당 한도 90~100% 구간 결제 집중", basis: `최근 3개월 ${nf(near.length)}건이 계정 한도의 90~100% 구간` + (nearBy.length ? `, 최다 ${nearBy[0][0]} ${nf(nearBy[0][1].length)}건` : ""),
      why: "한도를 의식한 금액 조정(분할결제의 변형) 가능성. R05는 초과분만 잡음", rule: "직원별 3개월 내 한도 90% 이상 결제 3건 이상", count: near.length,
    });
    // 2) 같은 날 같은 가맹점 다른 카드 결제 (카드 분산)
    const multi = [...groupBy(recent, (r) => `${r.merchant}|${r.dkey}`).values()].filter((g) => new Set(g.map((r) => r.card)).size >= 2);
    recs.push({
      title: "같은 날 같은 가맹점, 서로 다른 카드 결제", basis: `최근 3개월 ${nf(multi.length)}개 가맹점-일자에서 2장 이상 카드 사용`,
      why: "한 자리 비용을 여러 사람 카드로 나눠 R04(동일 카드)를 피하는 패턴", rule: "가맹점·승인일 동일, 카드 2장 이상, 합계 > 건당한도", count: multi.length,
    });
    // 3) 적요 문구 복사 반복
    const memoDup = [...groupBy(recent.filter((r) => r.memo.length >= 5), (r) => `${r.emp}|${r.memo}`).values()].filter((g) => g.length >= 4);
    recs.push({
      title: "같은 적요 문구 반복 사용", basis: `최근 3개월 동일 직원이 같은 적요를 4회 이상 쓴 경우 ${nf(memoDup.length)}건` + (memoDup.length ? ` (예: '${memoDup[0][0].memo}' ${nf(memoDup[0].length)}회)` : ""),
      why: "실제 상대방·목적이 바뀌어도 적요를 복사해 형식만 맞추는 경우. R06은 형식만 확인", rule: "직원×적요 3개월 4회 이상 → 상대방 실명 증빙 확인", count: memoDup.length,
    });
    // 4) 금요일 저녁 접대비 집중
    const fri = recent.filter((r) => r.glName.includes("접대비") && r.wd === 5 && r.hour >= 18);
    const ent = recent.filter((r) => r.glName.includes("접대비"));
    recs.push({
      title: "금요일 저녁 접대비 비중", basis: `최근 3개월 접대비 ${nf(ent.length)}건 중 금요일 18시 이후 ${nf(fri.length)}건(${ent.length ? Math.round((fri.length / ent.length) * 100) : 0}%)`,
      why: "주말 직전 사적 모임을 접대비로 처리하는 경우가 많은 시간대", rule: "직원별 접대비 중 금요일 저녁 비중 50% 이상", count: fri.length,
    });
    // 5) 월말 마지막 영업일 접대비 정액
    const lastDays = new Set(win3.map((m) => businessDays(m, ref).slice(-1)[0]));
    const endEnt = recent.filter((r) => lastDays.has(r.dkey) && r.glName.includes("접대비"));
    recs.push({
      title: "월 마지막 영업일 접대비 결제", basis: `최근 3개월 마지막 영업일 접대비 ${nf(endEnt.length)}건, ${won(sum(endEnt))}`,
      why: "잔여 예산 소진·선결제 목적의 월말 결제 의심. T-R3은 부서·분기말만 봄", rule: "직원별 월 마지막 영업일 접대비 2개월 이상", count: endEnt.length,
    });
    return recs.sort((a, b) => b.count - a.count);
  }

  function monthlySummary(rows) {
    const months = [...new Set(rows.map((r) => r.month))].sort();
    return months.map((m, i) => {
      const g = rows.filter((r) => r.month === m);
      const amt = sum(g);
      const prev3 = months.slice(Math.max(0, i - 2), i + 1).map((x) => sum(rows.filter((r) => r.month === x)));
      return { month: m, n: g.length, amount: amt, avg: g.length ? amt / g.length : 0, ma3: prev3.reduce((s, x) => s + x, 0) / prev3.length, flagged: g.filter((r) => r.score > 0).length };
    });
  }

  return { OPS, DERIVED: Object.keys(DERIVED), ruleName, reasonShort, reasonParts, parseCustom, RULE_NAMES, RULE_IDS, T_NAMES, DEFAULT_PARAMS, DEFAULT_WEIGHTS, buildRef, prepare, applyRules, periodRules, selectMonth, grade, recommend, monthlySummary, addMonths, won, ymd, toDate, businessDays };
})();


if (typeof module !== "undefined") module.exports = { ENGINE };
