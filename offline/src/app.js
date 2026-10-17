// ───────────── 화면 ─────────────
(() => {
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const won = ENGINE.won;
  // 직원명은 카드 주인, 참조는 전표 작성자 → '권윤슬(20016/권윤슬)', 직원명 공란이면 '부서공용(20012/류단비)'
  const who = (r) => `${r.emp || "부서공용"}${r.writer ? `(${r.writer})` : ""}`;
  const fmt = (n) => Math.round(n).toLocaleString("ko-KR");
  // 천 단위 쉼표: 소수는 그대로 두고 정수부만 끊음 (0.5 → 0.5, 300000 → 300,000)
  const fmt2 = (n) => Number(n).toLocaleString("ko-KR", { maximumFractionDigits: 4 });
  const unComma = (v) => Number(String(v).replace(/[,\s원]/g, ""));
  const commaInput = (el) => { const raw = el.value.replace(/[^\d.]/g, ""); if (raw === "" || raw.endsWith(".")) { el.value = raw; return; } el.value = fmt2(Number(raw)); };
  const condVal = (x) => (["이상", "이하", "초과", "미만"].includes(x.op) && isFinite(unComma(x.val)) ? fmt2(unComma(x.val)) : x.val);
  // 엑셀 저장: 금액 성격 컬럼은 숫자 그대로 두고 표시 형식만 #,##0 (승인번호·계정코드 같은 번호 컬럼은 제외)
  const MONEY_COL = /금액|한도|사용액|평균|합계|공급가액|부가세|봉사료|지표값/;
  function commaSheet(ws) {
    if (!ws["!ref"]) return ws;
    const rg = XLSX.utils.decode_range(ws["!ref"]);
    for (let c = rg.s.c; c <= rg.e.c; c++) {
      const h = ws[XLSX.utils.encode_cell({ r: rg.s.r, c })];
      if (!h || !MONEY_COL.test(String(h.v))) continue;
      for (let r = rg.s.r + 1; r <= rg.e.r; r++) { const cell = ws[XLSX.utils.encode_cell({ r, c })]; if (cell && cell.t === "n") cell.z = "#,##0"; }
    }
    return ws;
  }
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  };

  const REF_SHEETS = { accounts: "계정코드", industries: "업종코드", holidays: "공휴일", rules: "점검기준", employees: "직원마스터", vacations: "휴가내역", custom: "추가조건" };
  const REF_LABELS = REF_SHEETS;
  const S = {
    source: "none", files: [], raw: new Map(), sheetRows: 0, readCount: 0, refRaw: {}, refFiles: {}, answers: null, loadMs: 0,
    params: { ...JSON.parse(JSON.stringify(ENGINE.DEFAULT_PARAMS)), ...(store.get("cc_rules", {}).params || {}) }, weights: null,
    weightOverride: store.get("cc_rules", {}).weights || null, customRows: store.get("cc_rules", {}).customRows || null,
    month: null, tab: ["list", "answer", "period", "recs", "browse"].includes(store.get("cc_tab", "list")) ? store.get("cc_tab", "list") : "list", grade: "전체", review: "전체", sel: null,
    reviews: store.get("cc_reviews", {}), ms: 0,
  };
  let R = null; // 계산 결과

  // ── 계산 파이프라인 ──
  function run() {
    if (!S.raw.size || !S.refRaw.accounts) { renderEmpty(); return; }
    document.body.classList.remove("is-empty");
    const t0 = performance.now();
    const { ref, prep, rows, months, period, list, custom, columns } = compute(S.params, S.weightOverride, S.customRows || S.refRaw.custom || []);
    S.weights = { ...ref.weights };
    const graded = S.answers && S.answers.length ? ENGINE.grade(rows, S.answers) : null;
    // 기간 규칙에 걸린 대상월 건도 목록에 표시 (단건 점수 없음 → T 표시만)
    const recs = ENGINE.recommend(rows, ref, S.month);
    const summary = ENGINE.monthlySummary(rows);
    R = { ref, prep, rows, months, period, graded, list, recs, summary, custom, columns };
    render();
    S.ms = S.loadMs + (performance.now() - t0);
    renderScore();
  }

  // 같은 데이터에 기준값·가중치·추가조건을 넣어 계산 (규칙 수정 화면의 미리보기에도 씀)
  function compute(params, weightOverride, customRows) {
    const ref = ENGINE.buildRef(S.refRaw);
    const columns = [...new Set([...S.raw.values()].slice(0, 50).flatMap((r) => Object.keys(r)))];
    const custom = ENGINE.parseCustom(customRows, columns);
    for (const c of custom.rules) ref.weights[c.id] = c.weight;
    if (weightOverride) for (const [k, v] of Object.entries(weightOverride)) if (k in ref.weights || custom.rules.some((c) => c.id === k)) ref.weights[k] = v;
    const prep = ENGINE.prepare([...S.raw.values()], ref);
    const rows = ENGINE.applyRules(prep.rows, ref, params, custom.rules);
    const months = [...new Set(rows.map((r) => r.month))].sort();
    if (!S.month || !months.includes(S.month)) S.month = months[months.length - 1];
    const period = ENGINE.periodRules(rows, ref, params, S.month);
    const list = ENGINE.selectMonth(rows, S.month);
    return { ref, prep, rows, months, period, list, custom, columns };
  }

  // ── 렌더 ──
  function render() {
    $("#month").innerHTML = R.months.map((m) => `<option value="${m}" ${m === S.month ? "selected" : ""}>${m.replace("-", "년 ")}월</option>`).join("");
    const refTxt = Object.entries(REF_LABELS).map(([k, v]) => {
      const n = S.refRaw[k]?.length || 0, src = S.refFiles[k] || [];
      const by = src.length > 1 ? ` (${src.map((x) => `${fmt(x.n)}행 ${esc(x.file.replace(/\.[^.]+$/, ""))}`).join(" + ")})` : "";
      return `${v} ${n ? `${fmt(n)}행${by}` : "–"}`;
    }).join(" · ");
    const srcTxt = `<span class="tag">업로드</span> ${S.files.map(esc).join(", ")} · 승인번호 기준 중복 제거 후 ${fmt(S.raw.size)}건 · 기준표: ${refTxt}`;
    $("#srcLine").innerHTML = srcTxt + ` · 분석 대상 ${fmt(R.rows.length)}건 (취소 ${fmt(R.prep.stats.cancels)}건 상계 제외 ${fmt(R.prep.stats.matched * 2)}건${R.prep.stats.unmatched ? `, 미매칭 ${fmt(R.prep.stats.unmatched)}건` : ""})`;
    renderTabs(); renderList(); renderAnswers(); renderPeriod(); renderRecs(); renderBrowseFilters(); renderBrowse();
    if (!$("#ruleDrawer").hidden) renderRuleDrawer();
  }

  function renderScore() {
    const g = R.graded;
    const cnt = { 상: 0, 중: 0, 하: 0 }; R.list.forEach((r) => cnt[r.grade]++);
    const crit = [
      { lbl: "데이터 읽기", val: `${fmt(S.readCount)}<small> / ${fmt(S.sheetRows)}행</small>`, tgt: "원본 행수와 읽은 건수 일치", ok: S.readCount === S.sheetRows && S.sheetRows > 0 },
      { lbl: "업로드 → 목록 표시", val: `${(S.ms / 1000).toFixed(2)}<small> 초</small>`, tgt: "1분 이내", ok: S.ms < 60000 },
      { lbl: `${S.month.replace("-", "년 ")}월 선별`, val: `${fmt(R.list.length)}<small> 건</small>`, tgt: `상 ${fmt(cnt.상)} · 중 ${fmt(cnt.중)} · 하 ${fmt(cnt.하)}`, ok: undefined },
    ];
    if (g) crit.push(
      { lbl: "정답지 단건 탐지", val: `${g.sHit}<small> / ${g.S.length}종</small>`, tgt: "전부 탐지", ok: g.S.length > 0 && g.sHit === g.S.length },
      { lbl: "정답지 기간 패턴", val: `${g.tHit}<small> / ${g.T.length}종</small>`, tgt: "전부 탐지", ok: g.T.length > 0 && g.tHit === g.T.length },
      { lbl: "정상거래 '상' 오분류", val: `${fmt(g.fpHigh)}<small> / ${fmt(g.fpN)}건</small>`, tgt: "0건", ok: g.fpN > 0 && g.fpHigh === 0 });
    else crit.push(
      { lbl: "분석 대상", val: `${fmt(R.rows.length)}<small> 건</small>`, tgt: `취소 ${fmt(R.prep.stats.cancels)}건 상계${R.prep.stats.unmatched ? `, 미매칭 ${fmt(R.prep.stats.unmatched)}건` : ""}`, ok: undefined },
      { lbl: "기간 패턴", val: `${fmt(Object.values(R.period).reduce((s, x) => s + x.length, 0))}<small> 건</small>`, tgt: "T-R1~T-R4 누적 분석", ok: undefined },
      { lbl: "추가 점검항목 추천", val: `${fmt(R.recs.length)}<small> 건</small>`, tgt: "최근 3개월 집계 기반", ok: undefined });
    const badge = (ok) => ok === undefined ? "" : `<span class="verdict ${ok ? "ok" : "ng"}">${ok ? "정상" : "확인 필요"}</span>`;
    $("#score").innerHTML = crit.map((c) => `<div class="crit"><div class="lbl">${c.lbl}${badge(c.ok)}</div><div class="val">${c.val}</div><div class="tgt">${c.tgt}</div></div>`).join("");
    $("#overall").innerHTML = (S.readCount !== S.sheetRows ? `<span class="ng-txt">승인번호가 빈 행 ${fmt(S.sheetRows - S.readCount)}개는 읽지 않았습니다. 원본 파일의 합계행·빈 행인지 확인하세요.</span> ` : "") +
      (g ? "정답지 시트가 있어 탐지율을 채점했습니다. " : "") + "처리시간은 엑셀 읽기부터 목록 표시까지 이 PC에서 잰 값입니다.";
    $("#scoreNote").textContent = `파일 ${S.files.length}개`;
  }

  function renderEmpty() {
    document.body.classList.add("is-empty");
    const need = [];
    if (!S.refRaw.accounts) need.push("기준표 엑셀(계정코드·업종코드·공휴일 시트)");
    if (!S.raw.size) need.push("SAP 추출 엑셀(승인번호·전표금액 컬럼이 있는 시트)");
    $("#srcLine").innerHTML = S.files.length ? `<span class="tag">업로드</span> ${S.files.map(esc).join(", ")} · <span class="ng-txt">아직 필요: ${need.join(", ")}</span>` : "";
    $("#score").innerHTML = ""; $("#overall").textContent = ""; $("#scoreNote").textContent = "";
  }

  function renderTabs() {
    document.querySelectorAll(".tab").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.tab === S.tab)));
    document.querySelectorAll("[data-pane]").forEach((p) => (p.hidden = p.dataset.pane !== S.tab));
    $("#cntList").textContent = fmt(R.list.length);
    $("#cntAns").textContent = R.graded ? `${R.graded.sHit + R.graded.tHit}/${R.graded.S.length + R.graded.T.length}` : "";
    $("#cntPer").textContent = fmt(Object.values(R.period).reduce((s, a) => s + a.length, 0));
    $("#cntRec").textContent = fmt(R.recs.length);
    $("#ansTab").hidden = !R.graded;
    if (!R.graded && S.tab === "answer") { S.tab = "list"; renderTabs(); }
  }

  const short = (s) => (s.length > 11 ? s.slice(0, 10) + "…" : s);
  const chip = (id, cls) => { const name = ENGINE.ruleName(id) === id ? ENGINE.T_NAMES[id] || id : ENGINE.ruleName(id);
    return `<button type="button" class="chip${cls}" data-rule="${esc(id)}" title="${esc(id)} ${esc(name)} · 눌러서 규칙 설명 보기">${esc(id)} <span class="cn">${esc(short(name))}</span></button>`; };
  // 규칙 성격별 분류: 칩 색과 칸 강조에 씀
  const RULE_CAT = { R03: "ban", R04: "money", R05: "money", R09: "money", R12: "money", R01: "when", R02: "when", R11: "when", R14: "when", R17: "when", R06: "doc", R07: "doc", R13: "doc", R15: "doc", R16: "doc" };
  const catOf = (id) => RULE_CAT[id] || (id.startsWith("T-") ? "period" : id.startsWith("U") ? "user" : "doc");
  const ruleChips = (r) => [...r.hits].sort().map((h) => chip(h, ` cat-${catOf(h)}${h === "R03" && r.restrict === "금지" ? " ban-hard" : ""}${h.startsWith("U") ? " u" : ""}`)).join("") + [...(r.tHits || [])].map((h) => chip(h, " t cat-period")).join("");
  // 근거가 있는 칸에 붙일 강조 class
  // 걸린 규칙별로 담당자가 확인할 칸(형광펜 위치)
  const RULE_SPOTS = { R01: ["day"], R02: ["time"], R03: ["ind"], R04: ["amt"], R05: ["amt"], R06: ["memo"], R07: ["gl"], R09: ["amt"], R11: ["mer"], R12: ["amt"],
    R13: ["memo", "gl", "title"], R14: ["time", "mer"], R15: ["settle"], R16: ["who"], R17: ["who", "day"] };
  const COL_SPOT = { 가맹점업종: "ind", 업종코드: "ind", 제한구분: "ind", 전표금액: "amt", "전표 품목 텍스트": "memo", "G/L 계정명": "gl", "G/L 계정": "gl", 가맹점명: "mer", 가맹점주소: "mer", 시도: "mer",
    직원명: "who", 요일: "day", 휴일: "day", 공휴일명: "day", 승인일: "day", 승인시: "time", 승인시간: "time", 결재문서제목: "title" };
  function spotMarker(r) {
    const spots = new Map();
    const add = (spot, id) => { if (!spots.has(spot)) spots.set(spot, []); spots.get(spot).push(`${id} ${ENGINE.ruleName(id)}`); };
    const spotsOf = (id) => RULE_SPOTS[id] || (R.custom.rules.find((x) => x.id === id)?.conds.map((x) => COL_SPOT[x.col]).filter(Boolean) ?? []);
    for (const id of r.hits) spotsOf(id).forEach((s) => add(s, id));
    // 근거 종류별 표시: 시각·요일·금액·사람·가맹점 = 빨강 굵게, 업종 = 금지 빨간 바탕/그 외 주황, 계정 = 보라 물결 밑줄, 적요·제목 = 노란 형광펜
    const kind = (spot) => (spot === "ind" ? (r.restrict === "금지" ? "ev-ban" : "ev-warn") : spot === "gl" ? "ev-wave" : spot === "memo" || spot === "title" ? "ev-memo" : "ev-red");
    const m = (spot, html) => (spots.has(spot) ? `<mark class="ev ${kind(spot)}" title="확인할 부분: ${esc([...new Set(spots.get(spot))].join(", "))}">${html}</mark>` : html);
    m.has = (spot) => spots.has(spot);
    return m;
  }
  const reviewPill = (r) => { const v = S.reviews[r.no]?.result || "미검토"; return `<span class="rv rv-${v}">${v}</span>`; };

  // ── 점검 규칙표 (오른쪽 패널) ──
  function ruleDesc(id) {
    const P = S.params, w = (x) => fmt(x);
    const D = {
      R01: `승인일이 토·일 또는 공휴일 시트에 있는 날${P.holidayExclude.length ? ` (${P.holidayExclude.join("·")} 제외)` : ""}`,
      R02: `승인 시각이 ${P.nightStart}시 이후 또는 ${P.nightEnd}시 이전${P.nightExclude.length ? ` (${P.nightExclude.join("·")} 제외)` : ""}`,
      R03: `가맹점 업종의 제한구분이 ${P.restrictTargets.join(" 또는 ") || "(대상 없음)"} (금지 업종은 점수와 관계없이 '상')`,
      R04: (P.splitDays || 0) > 0 ? `같은 카드·가맹점에 ${P.splitDays}일 이내로 ${fmt(P.splitMin)}건 이상, 합계가 계정 건당 한도 초과` : `같은 카드·가맹점·날짜에 ${P.splitGap}분 간격 이내로 ${fmt(P.splitMin)}건 이상, 합계가 계정 건당 한도 초과`,
      R05: "전표금액이 계정코드 시트의 건당 한도 초과",
      R06: `적요가 비었거나 ${P.memoMin}자 미만, 또는 접대비인데 상대방(ㅇㅇㅇ·직함·n인) 표기 없음${P.bareWords.length ? `, 또는 적요가 '${P.bareWords.join("', '")}'뿐` : ""}`,
      R07: "가맹점 업종이 계정코드 시트의 허용 가맹점업종에 없음",
      R09: `같은 ${P.dupKeys.join("·") || "(기준 없음)"}가 2건 이상 (취소로 상계된 건 제외)`,
      R11: "국내해외코드명이 해외",
      R12: `금액이 ${w(P.round)}원 단위로 떨어지고 ${w(P.round)}원 이상`,
      R13: "결재문서제목의 비용 항목(하이패스·항공·숙박 등)과 업종이 다르거나, 근무지 안 식사·회식을 여비교통비로 처리",
      R14: `같은 카드의 연속 2건이 시·도가 다르고 ${P.distGap}분 이내 (${P.distExclude.length}개 업종 제외: 온라인·교통 등)`,
      R15: `회계일자 − 승인일이 ${P.settleDays}일 초과`,
      R16: `공용카드이고 직원명이 비었으며 금액이 ${w(P.sharedMin)}원 이상`,
      R17: "휴가내역 시트의 휴가 기간 중 본인 카드 사용",
      "T-R1": `같은 직원·가맹점이 최근 3개월 ${fmt(P.t1Count)}회 이상`,
      "T-R2": `직원 월 사용액이 직전 6개월 평균의 ${P.t2Ratio}배 초과 (월 ${w(P.t2Floor)}원 이상)`,
      "T-R3": `분기말 월에 부서의 월말 5영업일 사용액이 ${Math.round(P.t3Share * 100)}% 초과, 월 사용액이 부서 평균의 ${P.t3Mult}배 이상`,
      "T-R4": `같은 직원·가맹점·금액이 ${P.t4Months}개월 이상 연속`,
    };
    if (D[id]) return D[id];
    const c = R.custom.rules.find((x) => x.id === id);
    return c ? c.conds.map((x) => `${x.col} ${x.op}${x.val !== "" ? " " + condVal(x) : ""}`).join(" 그리고 ") : "";
  }
  function renderRuleDrawer() {
    if (!R) return;
    const cnt = {};
    R.list.forEach((r) => { r.hits.forEach((h) => (cnt[h] = (cnt[h] || 0) + 1)); r.tHits.forEach((h) => (cnt[h] = (cnt[h] || 0) + 1)); });
    const off = new Set(S.params.off || []);
    // 가중치 1~5점은 점 5개, 그 밖(0점·6점 이상)은 숫자로 표시
    const dots = (w) => w >= 1 && w <= 5 && Number.isInteger(w) ? `<span class="dots" title="가중치 ${w}점" aria-label="가중치 ${w}점">${[1, 2, 3, 4, 5].map((k) => `<i class="${k <= w ? "on" : ""}"></i>`).join("")}</span>` : `<span class="wnum${w > 5 ? " big" : ""}" title="가중치 ${w}점">${fmt(w)}점</span>`;
    const more = (id, n) => `<div class="rr-more"><div class="rr-crit"><span>판단 기준</span>${esc(ruleDesc(id))}</div>${n ? `<button type="button" class="rf" data-filter="${esc(id)}">${S.ruleFilter === id ? "목록 필터 해제" : `이 규칙 ${fmt(n)}건만 목록에서 보기`}</button>` : ""}</div>`;
    const row = (id, name, weight, showW) => {
      const n = cnt[id] || 0, open = S.ruleFocus === id;
      return `<li class="rr${open ? " open" : ""}${n ? "" : " zero"}${off.has(id) ? " off" : ""}" id="rule-${esc(id)}">
        <button type="button" class="rr-top" data-toggle="${esc(id)}" aria-expanded="${open}">
          <b class="rr-id">${esc(id)}</b>
          <span class="rr-nm"><span class="rr-name">${esc(name)}${off.has(id) ? ' <em class="rr-off">사용 안 함</em>' : ""}</span><span class="rr-help">${esc(ENGINE.RULE_HELP[id] || "")}</span></span>
          ${showW ? (weight !== null ? dots(weight) : '<span class="rr-tag">표시만</span>') : ""}
          <span class="rr-n${n ? " hit" : ""}">${n ? fmt(n) + "건" : "–"}</span>
        </button>${open ? more(id, n) : ""}</li>`;
    };
    const scale = `<div class="scale" aria-label="등급 기준"><span class="sc sc-하"><b>하</b>1~2점</span><span class="sc sc-중"><b>중</b>3~5점</span><span class="sc sc-상"><b>상</b>6점 이상</span></div>
      <ul class="scale-note"><li><b>위험점수</b> = 걸린 규칙 가중치의 합</li><li><b>금지 업종</b>은 점수와 관계없이 바로 <b>상</b></li><li><b>사전승인 감점</b>: 휴일·심야만 걸리고 적요에 사전승인 문서번호가 있으면 규칙당 1점 감점</li></ul>`;

    const ids = ENGINE.RULE_IDS, hitN = ids.filter((id) => cnt[id]).length;
    const only = !!S.ruleOnlyHit, keep = (id) => !only || cnt[id] || S.ruleFocus === id;
    const wOf = (id) => S.weights[id] ?? 0;
    const groups = [...new Set(ids.map(wOf))].sort((a, b) => b - a).map((w) => [w, ids.filter((id) => wOf(id) === w && keep(id))]).filter(([, g]) => g.length);
    const sum = `<div class="rsum"><div><b>${fmt(hitN)}</b><span>/ ${fmt(ids.length)}개 규칙이 ${esc(S.month)}에 걸림</span></div>
      <div class="seg2" role="group" aria-label="보기"><button type="button" data-only="0" aria-pressed="${!only}">전체 규칙</button><button type="button" data-only="1" aria-pressed="${only}">걸린 규칙만</button></div></div>`;
    const grp = (w, g) => `<section class="wgrp w${w}"><h3>${dots(w)}<span>${w === 0 ? "가중치 0점 · 점수에 안 들어감" : `가중치 ${fmt(w)}점`}</span><small>${g.length}개</small></h3><ul>${g.map((id) => row(id, ENGINE.RULE_NAMES[id], w, false)).join("")}</ul></section>`;
    const cus = R.custom.rules.filter((c) => keep(c.id)), per = Object.keys(ENGINE.T_NAMES).filter(keep);
    $("#ruleBody").innerHTML = sum + scale + groups.map(([w, g]) => grp(w, g)).join("") +
      (cus.length ? `<section class="wgrp"><h3><span>추가조건</span><small>${S.customRows ? "규칙 수정 화면에서 고친 조건" : "기준표 시트"}</small></h3><ul>${cus.map((c) => row(c.id, c.name, S.weights[c.id] ?? c.weight, true)).join("")}</ul></section>` : "") +
      (per.length ? `<section class="wgrp"><h3><span>기간 규칙</span><small>점수 없이 표시만</small></h3><ul>${per.map((id) => row(id, ENGINE.T_NAMES[id], null, false)).join("")}</ul></section>` : "");
    if (S.ruleFocus) document.getElementById("rule-" + S.ruleFocus)?.scrollIntoView({ block: "nearest" });
  }
  function openRules(id) {
    S.ruleFocus = id || null;
    $("#ruleDrawer").hidden = false; document.body.classList.add("drawer-open");
    $("#rulesBtn").setAttribute("aria-expanded", "true");
    renderRuleDrawer();
  }
  function closeRules() { $("#ruleDrawer").hidden = true; document.body.classList.remove("drawer-open"); $("#rulesBtn").setAttribute("aria-expanded", "false"); S.ruleFocus = null; }
  document.addEventListener("click", (e) => {
    const c = e.target.closest(".chip[data-rule]");
    if (c) { e.stopPropagation(); openRules(c.dataset.rule); return; }
  }, true);
  $("#rulesBtn").addEventListener("click", () => ($("#ruleDrawer").hidden ? openRules() : closeRules()));
  $("#ruleClose").addEventListener("click", closeRules);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !$("#ruleDrawer").hidden) closeRules(); });
  $("#ruleBody").addEventListener("click", (e) => {
    const t = e.target.closest("[data-toggle]"); if (t) { S.ruleFocus = S.ruleFocus === t.dataset.toggle ? null : t.dataset.toggle; renderRuleDrawer(); return; }
    const o = e.target.closest("[data-only]"); if (o) { S.ruleOnlyHit = o.dataset.only === "1"; renderRuleDrawer(); return; }
    const b = e.target.closest("[data-filter]"); if (!b) return;
    S.ruleFilter = S.ruleFilter === b.dataset.filter ? null : b.dataset.filter; S.grade = "전체"; S.tab = "list";
    renderTabs(); renderList(); renderRuleDrawer();
  });

  function renderList() {
    const counts = { 전체: R.list.length, 상: 0, 중: 0, 하: 0 };
    R.list.forEach((r) => counts[r.grade]++);
    $("#gradeFilter").innerHTML = `<span class="flabel">등급</span>` + Object.entries(counts).map(([k, v]) => `<button type="button" data-g="${k}" aria-pressed="${S.grade === k}">${k} <span class="mono">${fmt(v)}</span></button>`).join("");
    const reviewOf = (r) => S.reviews[r.no]?.result || "미검토";
    const rcount = { 전체: R.list.length, 미검토: 0, 적정: 0, 소명요청: 0, 부적정: 0 };
    R.list.forEach((r) => rcount[reviewOf(r)]++);
    if (!(S.review in rcount)) S.review = "전체";
    $("#reviewFilter").innerHTML = `<span class="flabel">검토</span>` + Object.entries(rcount).map(([k, v]) => `<button type="button" data-r="${k}" aria-pressed="${S.review === k}">${k} <span class="mono">${v}</span></button>`).join("");
    if (S.ruleFilter && !R.list.some((r) => r.hits.has(S.ruleFilter) || r.tHits.has(S.ruleFilter))) S.ruleFilter = null;
    if (S.ruleFilter) $("#gradeFilter").insertAdjacentHTML("beforeend", `<button type="button" class="rulef" data-clear-rule="1" aria-pressed="true">${esc(S.ruleFilter)} ${esc(ENGINE.ruleName(S.ruleFilter) === S.ruleFilter ? ENGINE.T_NAMES[S.ruleFilter] || "" : ENGINE.ruleName(S.ruleFilter))} 만 보는 중 ✕</button>`);
    const rows = R.list.filter((r) => (S.grade === "전체" || r.grade === S.grade) && (S.review === "전체" || reviewOf(r) === S.review) && (!S.ruleFilter || r.hits.has(S.ruleFilter) || r.tHits.has(S.ruleFilter)));
    // 선택한 건이 지금 목록(등급·검토결과·규칙 필터)에 없으면 상세를 닫음
    if (S.sel && !rows.some((r) => r.no === S.sel)) S.sel = null;
    const done = R.list.filter((r) => S.reviews[r.no]?.result).length;
    $("#reviewProgress").textContent = `검토 입력 ${fmt(done)}/${fmt(R.list.length)}건`;
    renderMailBtn();
    // 상세 영역을 표 밖으로 되돌린 뒤 표를 다시 그리고, 선택한 행 바로 아래로 옮김
    const det = $("#detail"); if (det.closest("#listTbl")) $("#listTbl").closest(".tbl-wrap").after(det);
    $("#listTbl").innerHTML = `<thead><tr><th>등급</th><th>점수</th><th>걸린 규칙</th><th>승인일시</th><th>직원 · 부서</th><th>가맹점 · 적요</th><th style="text-align:right">금액</th><th>선별 사유</th><th>검토</th></tr></thead><tbody>` +
      (rows.length ? rows.map((r) => { const m = spotMarker(r); return `<tr class="click sev-${r.grade} rvrow-${S.reviews[r.no]?.result || "미검토"} ${S.sel === r.no ? "on" : ""}" data-no="${r.no}" tabindex="0">
        <td><span class="g g-${r.grade}">${r.grade}</span></td><td class="n">${r.score}${r.adjust ? `<div class="muted" title="사전승인 기재로 조정">${r.ruleScore}→${r.score}</div>` : ""}</td>
        <td><div class="chips">${ruleChips(r)}</div></td>
        <td class="when"><span class="mono">${r.dkey.slice(5)}</span>(${m("day", r.holName || r.wdName)}) <span class="mono">${m("time", r.time)}</span>${r.hits.has("R17") ? `<div class="muted">${m("day", "휴가 중")}</div>` : ""}${r.hits.has("R15") ? `<div class="muted">${m("settle", `정산 ${Math.round((r.acctDate - r.date) / 86400000)}일`)}</div>` : ""}</td>
        <td style="white-space:nowrap">${m("who", r.emp ? esc(who(r)) : `<span class="ng-txt">부서공용</span>${r.writer ? esc(`(${r.writer})`) : ""}`)}<div class="muted">${esc(r.dept)}</div></td>
        <td class="mer">${m("mer", esc(r.merchant))}${r.overseas ? ` <span class="muted">(${esc(r.currency || "해외")})</span>` : ""}<div class="muted">${m("ind", esc(r.industry))} · ${m("gl", esc(r.glName))}</div>
          <div class="memo-line"><span class="lbl">적요</span><span>${m("memo", r.memo ? esc(r.memo) : '<span class="ng-txt">적요 공란</span>')}</span></div>${m.has("title") ? `<div class="memo-line"><span class="lbl">제목</span><span>${m("title", esc(r.title))}</span></div>` : ""}</td>
        <td class="n">${m("amt", fmt(r.amount))}${r.limit && (r.hits.has("R05") || r.hits.has("R04")) ? `<div class="muted">한도 ${fmt(r.limit)}</div>` : ""}</td>
        <td class="reason"><ul class="why">${ENGINE.reasonParts(r).map((x) => `<li>${esc(x)}</li>`).join("")}</ul></td><td style="white-space:nowrap">${reviewPill(r)}</td></tr>`; }).join("")
        : `<tr><td colspan="9" class="empty">선택한 조건(등급·검토결과·규칙)에 해당하는 선별 건이 없습니다.</td></tr>`) + "</tbody>";
    const selTr = S.sel && document.querySelector(`#listTbl tr[data-no="${S.sel}"]`);
    if (selTr) { const tr = document.createElement("tr"); tr.className = "detail-row"; tr.innerHTML = '<td colspan="9"></td>'; selTr.after(tr); tr.firstChild.append(det); }
    renderDetail();
  }

  function renderDetail() {
    const r = R.list.find((x) => x.no === S.sel);
    if (!r) { $("#detail").innerHTML = `<p class="empty">행을 누르면 그 행 바로 아래에 건별 상세와 검토결과 입력란이 펼쳐집니다.</p>`; return; }
    const rv = S.reviews[r.no] || {};
    $("#detail").innerHTML = `<div class="detail">
      <div class="panel"><h2><span class="g g-${r.grade}">${r.grade}</span> 승인번호 <span class="mono">${r.no}</span> <small>위험점수 ${r.score}${r.adjust ? ` (규칙 ${r.ruleScore}점, AI 조정 ${r.adjust})` : ""}</small></h2>
        <dl class="kv">
          <dt>승인일시</dt><dd class="mono">${r.dkey} ${r.time} (${r.holName || r.wdName + "요일"})</dd>
          <dt>카드 · 구분</dt><dd class="mono">${esc(r.card)} · ${esc(r.kind)}</dd>
          <dt>직원 · 부서</dt><dd>${esc(who(r))} · ${esc(r.dept)}${r.work ? ` (근무지 ${esc(r.work)})` : ""}</dd>
          <dt>가맹점</dt><dd>${esc(r.merchant)} · ${esc(r.industry)} (${esc(r.restrict)})<br><span class="muted">${esc(r.addr)}</span></dd>
          <dt>계정</dt><dd>${esc(r.gl)} ${esc(r.glName)}${r.limit ? ` · 건당한도 ${won(r.limit)}` : ""}</dd>
          <dt>금액</dt><dd class="mono">${won(r.amount)}${r.overseas ? ` (${esc(r.currency)})` : ""}</dd>
          <dt>결재문서제목</dt><dd>${esc(r.title)}</dd>
          <dt>적요</dt><dd>${r.memo ? esc(r.memo) : '<span class="ng-txt">공란</span>'}</dd>
          <dt>회계일자</dt><dd class="mono">${r.acctDate ? ENGINE.ymd(r.acctDate) : "-"}</dd>
          <dt>걸린 규칙</dt><dd>${[...r.hits].sort().map((h) => `${h} ${ENGINE.ruleName(h)}`).join(" · ")}${[...r.tHits].map((h) => ` · ${h} ${ENGINE.T_NAMES[h]}`).join("")}</dd>
        </dl>
      </div>
      <div style="display:grid;gap:12px;align-content:start">
        <div class="ai-box">
          <h3>선별 사유 <span class="muted" style="font-weight:400">· 규칙 기반 자동 문장</span></h3>
          <p>${esc(r.reason)}</p>
          <h3>추가 확인 필요사항</h3>
          <ul>${r.checks.map((c) => `<li>${esc(c)}</li>`).join("")}</ul>
        </div>
        <div class="panel review">
          <h2>검토결과 <small>다음 달 AI 추천의 근거 데이터</small></h2>
          <div class="row seg" id="seg">${["적정", "소명요청", "부적정"].map((v) => `<button type="button" data-v="${v}" aria-pressed="${rv.result === v}">${v}</button>`).join("")}</div>
          <label for="opinion">검토의견</label>
          <textarea id="opinion" placeholder="예: 사전승인 문서 GW-202608-021 확인, 적정">${esc(rv.opinion || "")}</textarea>
        </div>
      </div></div>`;
  }

  function renderAnswers() {
    const g = R.graded;
    if (!g) { $("#ansSummary").textContent = "업로드한 파일에 정답지 시트가 없어 채점할 수 없습니다."; $("#ansTbl").innerHTML = ""; return; }
    $("#ansSummary").innerHTML = `단건 <b>${g.sHit}/${g.S.length}</b> · 기간 패턴 <b>${g.tHit}/${g.T.length}</b> · 정상거래(FP) ${fmt(g.fpN)}건 중 '상' <b>${fmt(g.fpHigh)}건</b>. 정답지는 채점에만 쓰고 선별 로직에는 넣지 않습니다.`;
    const fpGroup = g.list.filter((o) => o.type === "FP");
    const body = g.list.filter((o) => o.type !== "FP").map((o) => `<tr>
      <td class="mono">${o.id}</td><td>${esc(o.title)}</td><td class="mono">${esc(o.expect)}</td>
      <td>${o.items.map((i) => i.row ? `<div class="chips" style="margin-bottom:2px">${o.type === "S" ? `<span class="g g-${i.row.grade}" style="width:20px;height:18px;font-size:11px">${i.row.grade || "–"}</span>` : ""}${o.type === "S" ? ruleChips({ hits: i.row.hits, tHits: new Set() }) : [...i.row.tHits].map((h) => `<span class="chip t">${h}</span>`).join("")}</div>` : `<div class="ng-txt">승인번호 ${i.no} 없음</div>`).slice(0, o.type === "T" ? 1 : 9).join("")}${o.type === "T" ? `<div class="muted">${fmt(o.items.length)}건 묶음</div>` : ""}</td>
      <td>${o.detected ? `<span class="ok-txt">탐지</span>` : `<span class="ng-txt">미탐</span>`}<div class="muted">${esc(o.note)}</div></td></tr>`).join("");
    const fpRow = fpGroup.length ? `<tr><td class="mono">FP01–${fpGroup[fpGroup.length - 1].id.slice(2)}</td><td>${esc(fpGroup[0].title)}</td><td class="mono">${esc(fpGroup[0].expect)}</td>
      <td>${Object.entries(g.fpGrades).map(([k, v]) => `${k} ${fmt(v)}건`).join(" · ")}</td><td>${g.fpHigh === 0 ? `<span class="ok-txt">상 0건</span>` : `<span class="ng-txt">상 ${g.fpHigh}건</span>`}<div class="muted">사전승인 문서번호로 저위험 조정</div></td></tr>` : "";
    $("#ansTbl").innerHTML = `<thead><tr><th>ID</th><th>시나리오</th><th>기대 규칙</th><th>앱 결과(등급·걸린 규칙)</th><th>판정</th></tr></thead><tbody>${body}${fpRow}</tbody>`;
  }

  function renderPeriod() {
    const sm = R.summary;
    const W = 720, H = 230, pl = 52, pr = 10, pt = 12, pb = 28;
    const max = Math.max(...sm.map((m) => m.amount), 1);
    const step = niceStep(max / 4);
    const top = Math.ceil(max / step) * step;
    const bw = (W - pl - pr) / sm.length;
    const y = (v) => pt + (H - pt - pb) * (1 - v / top);
    let svg = "";
    for (let v = 0; v <= top; v += step) svg += `<line x1="${pl}" x2="${W - pr}" y1="${y(v)}" y2="${y(v)}" stroke="var(--rule-2)"/><text x="${pl - 6}" y="${y(v) + 4}" text-anchor="end">${fmt(v / 1e6)}백만</text>`;
    sm.forEach((m, i) => {
      const x = pl + i * bw + bw * 0.18, w = bw * 0.64;
      svg += `<rect x="${x}" y="${y(m.amount)}" width="${w}" height="${y(0) - y(m.amount)}" rx="2" fill="${m.month === S.month ? "var(--accent)" : "var(--rule)"}"><title>${m.month}: ${won(m.amount)} (${fmt(m.n)}건)</title></rect>`;
      svg += `<text x="${pl + i * bw + bw / 2}" y="${H - 10}" text-anchor="middle">${m.month.slice(2).replace("-", ".")}</text>`;
    });
    const pts = sm.map((m, i) => `${pl + i * bw + bw / 2},${y(m.ma3)}`).join(" ");
    svg += `<polyline points="${pts}" fill="none" stroke="var(--mid)" stroke-width="2"/>`;
    const last = sm[sm.length - 1];
    if (last) svg += `<circle cx="${pl + (sm.length - 1) * bw + bw / 2}" cy="${y(last.ma3)}" r="3.5" fill="var(--mid)"/>`;
    $("#chart").innerHTML = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="월별 사용액 막대 차트">${svg}</svg>`;

    $("#monthTbl").innerHTML = `<thead><tr><th>월</th><th style="text-align:right">건수</th><th style="text-align:right">금액</th><th style="text-align:right">건당평균</th><th style="text-align:right">전월 대비</th><th style="text-align:right">선별</th></tr></thead><tbody>` +
      sm.map((m, i) => { const p = sm[i - 1]; const d = p ? ((m.amount - p.amount) / p.amount) * 100 : null;
        return `<tr${m.month === S.month ? ' class="on"' : ""}><td class="mono">${m.month}</td><td class="n">${fmt(m.n)}</td><td class="n">${fmt(m.amount)}</td><td class="n">${fmt(m.avg)}</td><td class="n">${d === null ? "–" : (d > 0 ? "+" : "") + d.toFixed(1) + "%"}</td><td class="n">${fmt(m.flagged)}</td></tr>`; }).join("") + "</tbody>";

    const P = S.params;
    const defs = {
      "T-R1": `직원×가맹점 최근 3개월(${ENGINE.addMonths(S.month, -2)}~${S.month}) ${fmt(P.t1Count)}회 이상`,
      "T-R2": `직원 월 사용액 ÷ 직전 6개월 평균 > ${P.t2Ratio} (월 ${fmt(P.t2Floor)}원 이상, 최근 3개월)`,
      "T-R3": `분기말 월 부서 월말 5영업일 비중 > ${P.t3Share * 100}% 그리고 월 사용액 ≥ 부서 월평균 × ${P.t3Mult}`,
      "T-R4": `동일 직원·가맹점·금액 ${P.t4Months}개월 이상 연속`,
    };
    const item = {
      "T-R1": (x) => `<tr><td>${esc(x.who)}<div class="muted">${esc(x.dept)}</div></td><td>${esc(x.merchant)}</td><td class="n">${fmt(x.n)}회</td><td class="n">${fmt(x.amount)}</td><td>${esc(x.focus) || "–"}</td></tr>`,
      "T-R2": (x) => `<tr><td>${esc(x.who)}<div class="muted">${esc(x.dept)}</div></td><td class="mono">${x.month}</td><td class="n">${fmt(x.cur)}</td><td class="n">${fmt(x.avg)}</td><td class="n"><b>×${x.ratio.toFixed(1)}</b></td><td>${spark(x.series)}</td></tr>`,
      "T-R3": (x) => `<tr><td>${esc(x.who)}</td><td class="mono">${x.month}</td><td class="n">${(x.share * 100).toFixed(0)}%</td><td class="n">${fmt(x.total)}</td><td class="n">×${(x.total / x.avg).toFixed(1)}</td></tr>`,
      "T-R4": (x) => `<tr><td>${esc(x.who)}<div class="muted">${esc(x.dept)}</div></td><td>${esc(x.merchant)}</td><td class="n">${fmt(x.amount)}</td><td class="n">${x.months}개월</td></tr>`,
    };
    const heads = { "T-R1": ["직원", "가맹점", "3개월 건수", "금액", "요일 집중"], "T-R2": ["직원", "기준월", "월 사용액", "직전 6개월 평균", "배율", "12개월 추이"], "T-R3": ["부서", "월", "월말 5영업일 비중", "월 사용액", "부서 평균 대비"], "T-R4": ["직원", "가맹점", "금액", "연속"] };
    $("#patterns").innerHTML = Object.keys(ENGINE.T_NAMES).map((k) => `<div class="panel pattern"><div class="head"><b>${k}</b><h2 style="margin:0">${ENGINE.T_NAMES[k]}</h2><span class="tag">${fmt(R.period[k].length)}건</span></div>
      <div class="def">${defs[k]}</div>
      ${R.period[k].length ? `<div class="tbl-wrap"><table><thead><tr>${heads[k].map((h, i) => `<th${i >= 2 && k !== "T-R1" ? ' style="text-align:right"' : ""}>${h}</th>`).join("")}</tr></thead><tbody>${R.period[k].map(item[k]).join("")}</tbody></table></div>` : `<div class="empty">해당 패턴 없음</div>`}</div>`).join("");
  }
  function spark(series) {
    const w = 96, h = 22, max = Math.max(...series, 1);
    const pts = series.map((v, i) => `${(i / (series.length - 1)) * (w - 4) + 2},${h - 2 - (v / max) * (h - 4)}`).join(" ");
    const lp = pts.split(" ").pop().split(",");
    return `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" aria-hidden="true"><polyline points="${pts}" fill="none" stroke="var(--accent)" stroke-width="1.5"/><circle cx="${lp[0]}" cy="${lp[1]}" r="2.5" fill="var(--high)"/></svg>`;
  }
  function niceStep(x) { const p = Math.pow(10, Math.floor(Math.log10(x))); const f = x / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p; }

  function renderRecs() {
    const base = R.recs.map((r) => `<div class="rec"><h3>${esc(r.title)}</h3><div class="basis">${esc(r.basis)}</div><div class="why">${esc(r.why)}</div><div class="how">규칙안: ${esc(r.rule)}</div></div>`).join("");
    $("#recs").innerHTML = base;
  }

  // ── 점검 규칙 수정 화면 ──
  // 고친 값은 [적용]을 눌러야 반영. 이 PC 브라우저에 기억하고, 기준표 엑셀('규칙설정' 시트)로 내보내 다른 PC에서도 씀
  const RULE_KEYS = {
    R01: ["holidayExclude"], R02: ["nightStart", "nightEnd", "nightExclude"], R03: ["restrictTargets"], R04: ["splitDays", "splitGap", "splitMin"],
    R06: ["memoMin", "entWords", "counterpart", "bareWords"], R09: ["dupKeys"], R12: ["round"], R13: ["titleMap", "travelAcct", "mealInds", "travelWords"],
    R14: ["distGap", "distExclude"], R15: ["settleDays"], R16: ["sharedMin", "sharedWords"],
    "T-R1": ["t1Count"], "T-R2": ["t2Ratio", "t2Floor"], "T-R3": ["t3Share", "t3Mult"], "T-R4": ["t4Months"],
  };
  const DETAIL_RULES = ["R01", "R02", "R03", "R04", "R06", "R09", "R13", "R14", "R16"];
  const T_IDS = Object.keys(ENGINE.T_NAMES);
  const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  const clone = (x) => JSON.parse(JSON.stringify(x));
  let D = null, BASE = null, SIM = null, simTimer = null;

  const customFromRule = (c) => ({ id: c.id, name: c.name, weight: S.weights[c.id] ?? c.weight ?? 1, conds: c.conds.map((x) => ({ col: x.col, op: x.op, val: x.val })) });
  const customToRows = (list) => list.flatMap((c) => c.conds.map((x, i) => ({ 조건ID: c.id, 조건명: i ? "" : c.name, 컬럼: x.col, 비교: x.op, 값: x.val, 가중치: i ? "" : c.weight })));
  function ruleChanged(id, a, b) {
    if ((a.params.off || []).includes(id) !== (b.params.off || []).includes(id)) return true;
    if (!T_IDS.includes(id) && (a.weights[id] ?? 0) !== (b.weights[id] ?? 0)) return true;
    return (RULE_KEYS[id] || []).some((k) => !same(a.params[k], b.params[k]));
  }
  function customChanged(c, a, b) {
    const o = b.custom.find((x) => x.id === c.id);
    return !o || !same(c, o) || (a.params.off || []).includes(c.id) !== (b.params.off || []).includes(c.id) || (a.weights[c.id] ?? 0) !== (b.weights[c.id] ?? 0);
  }
  function changeCount(a, b) {
    const ids = [...ENGINE.RULE_IDS, ...T_IDS].filter((id) => ruleChanged(id, a, b)).length;
    const cus = a.custom.filter((c) => customChanged(c, a, b)).length + b.custom.filter((c) => !a.custom.some((x) => x.id === c.id)).length;
    return ids + cus;
  }
  const monthCounts = (list) => { const n = {}; list.forEach((r) => { r.hits.forEach((h) => (n[h] = (n[h] || 0) + 1)); (r.tHits || new Set()).forEach((h) => (n[h] = (n[h] || 0) + 1)); }); return n; };
  const gradeCounts = (list) => { const g = { 상: 0, 중: 0, 하: 0 }; list.forEach((r) => g[r.grade] !== undefined && g[r.grade]++); return g; };

  function openEditor(tab) {
    if (!R) return;
    BASE = { params: clone(S.params), weights: { ...S.weights }, custom: R.custom.rules.map(customFromRule) };
    D = { ...clone(BASE), tab: tab || "R", open: new Set() };
    SIM = null;
    $("#ruleEditor").hidden = false; document.body.style.overflow = "hidden";
    renderEditor(); simulate();
  }
  function closeEditor() { $("#ruleEditor").hidden = true; document.body.style.overflow = ""; D = null; }
  function simulate() {
    clearTimeout(simTimer);
    simTimer = setTimeout(() => {
      if (!D) return;
      try { SIM = compute(D.params, D.weights, customToRows(D.custom)); } catch (e) { console.error(e); SIM = null; }
      renderEditorCounts();
    }, 250);
  }

  // 입력 칸 만들기
  const pIn = (k, opt = {}) => {
    const v = D.params[k], was = BASE.params[k], c = !same(v, was);
    const shown = opt.pct ? fmt2(Math.round(v * 1000) / 10) : fmt2(v);
    return `<input class="${opt.w ? "w " : ""}${c ? "chg" : ""}" data-p="${k}"${opt.pct ? " data-pct" : ""} data-money inputmode="decimal" value="${shown}" aria-label="${esc(opt.label || k)}">${c ? `<span class="was">${opt.pct ? fmt2(Math.round(was * 1000) / 10) : fmt2(was)}</span>` : ""}`;
  };
  const chipsIn = (k, hint) => {
    const cur = D.params[k] || [], was = BASE.params[k] || [];
    return `<div class="chipset">${cur.map((x, i) => `<span class="ce${was.includes(x) ? "" : " add"}">${esc(x)}<button type="button" data-ck="${k}" data-ci="${i}" title="빼기" aria-label="${esc(x)} 빼기">×</button></span>`).join("")}${was.filter((x) => !cur.includes(x)).map((x) => `<span class="ce rm">${esc(x)}</span>`).join("")}<input class="chip-in" data-cadd="${k}" placeholder="+ 추가 (Enter)" list="${hint || ""}"></div>`;
  };
  const checks = (k, opts) => opts.map((o) => `<label><input type="checkbox" data-cb="${k}" value="${esc(o)}" ${(D.params[k] || []).includes(o) ? "checked" : ""}> ${esc(o)}</label>`).join("");
  const field = (label, body, help) => `<div class="df"><div class="dl">${label}</div><div class="dv">${body}${help ? `<div class="dh2">${help}</div>` : ""}</div></div>`;

  function critHtml(id) {
    const P = D.params;
    const dayTxt = (P.splitDays || 0) > 0 ? `${P.splitDays}일 이내` : "같은 날";
    return {
      R01: `승인일이 토·일 또는 공휴일(공휴일 시트)${P.holidayExclude.length ? `, ${esc(P.holidayExclude.join("·"))} 제외` : ""}`,
      R02: `승인시각 ${pIn("nightStart", { label: "심야 시작" })} 시 이상 또는 ${pIn("nightEnd", { label: "심야 종료" })} 시 미만${P.nightExclude.length ? `, ${esc(P.nightExclude.join("·"))} 제외` : ""}`,
      R03: `업종코드 시트 제한구분이 ${esc(P.restrictTargets.join("·") || "(없음)")} <span class="tag ban">금지는 점수 무관 '상'</span>`,
      R04: `같은 카드·가맹점, ${dayTxt} ${fmt(P.splitMin)}건 이상, 합계가 건당 한도 초과${(P.splitDays || 0) > 0 ? "" : `, 간격 ${pIn("splitGap", { label: "분할결제 간격" })} 분 이내`}`,
      R05: "전표금액이 계정별 건당 한도 초과 (계정코드 시트)",
      R06: `적요 ${pIn("memoMin", { label: "적요 최소 글자수" })} 자 미만, 접대비인데 상대방 표기 없음${P.bareWords.length ? `, 또는 '${esc(P.bareWords[0])}' 등 단어만 기재` : ""}`,
      R07: "가맹점업종이 계정의 허용 업종에 없음 (계정코드 시트)",
      R09: `같은 ${esc(P.dupKeys.join("·") || "(기준 없음)")} 2건 이상, 취소로 상계 안 됨`,
      R11: "국내해외코드명 = 해외",
      R12: `금액이 ${pIn("round", { w: 1, label: "라운드 단위" })} 원 단위로 떨어지고 그 이상`,
      R13: "결재문서제목·적요·계정·업종의 의미가 서로 맞지 않음",
      R14: `같은 카드 연속 2건이 시·도가 다르고 ${pIn("distGap", { label: "원거리 간격" })} 분 이내 (제외 업종 ${fmt(P.distExclude.length)}개)`,
      R15: `회계일자 − 승인일이 ${pIn("settleDays", { label: "정산 기한" })} 일 초과`,
      R16: `공용카드, 직원명 공란, 금액 ${pIn("sharedMin", { w: 1, label: "공용카드 기준금액" })} 원 이상`,
      R17: "휴가내역 기간 중 본인 카드 사용 (휴가내역 시트)",
      "T-R1": `같은 직원·가맹점이 최근 3개월 ${pIn("t1Count", { label: "반복 횟수" })} 회 이상`,
      "T-R2": `직원 월 사용액이 직전 6개월 평균의 ${pIn("t2Ratio", { label: "급증 배율" })} 배 초과, 월 ${pIn("t2Floor", { w: 1, label: "최소 월 사용액" })} 원 이상`,
      "T-R3": `분기말 월 부서의 월말 5영업일 비중 ${pIn("t3Share", { pct: 1, label: "월말 비중(%)" })} % 초과, 부서 평균의 ${pIn("t3Mult", { label: "부서 평균 배수" })} 배 이상`,
      "T-R4": `같은 직원·가맹점·금액이 ${pIn("t4Months", { label: "연속 개월" })} 개월 이상 연속`,
    }[id];
  }
  function detailHtml(id) {
    const P = D.params;
    switch (id) {
      case "R01": return field("제외 업종", chipsIn("holidayExclude", "dlInd"), "여기 넣은 업종은 휴일에 써도 R01로 보지 않습니다.");
      case "R02": return field("제외 업종", chipsIn("nightExclude", "dlInd"), "예: 야근 후 귀가 택시를 심야 사용에서 빼려면 '택시'를 넣습니다.");
      case "R03": return field("대상 제한구분", checks("restrictTargets", ["금지", "사전승인", "주의"]), "업종코드 시트의 제한구분 중 걸리게 할 값입니다. '금지'는 점수와 관계없이 등급 '상'입니다.");
      case "R04": return field("묶는 범위", `<select data-p="splitDays" class="${same(P.splitDays, BASE.params.splitDays) ? "" : "chg"}">${[0, 1, 2, 3, 5, 7].map((d) => `<option value="${d}" ${+P.splitDays === d ? "selected" : ""}>${d ? `${d}일 이내` : "같은 날"}</option>`).join("")}</select> 같은 카드·같은 가맹점`, "N일 이내로 바꾸면 날짜를 넘겨 나눠 결제한 경우도 묶어 봅니다.") +
        ((P.splitDays || 0) > 0 ? "" : field("결제 간격", `${pIn("splitGap", { label: "결제 간격" })} 분 이내 연속 결제를 한 묶음으로 봄`)) +
        field("최소 건수", `${pIn("splitMin", { label: "최소 건수" })} 건 이상`) +
        field("합계 기준", "계정별 건당 한도 초과", "건당 한도는 기준표 엑셀 '계정코드' 시트 값을 씁니다.");
      case "R06": return field("최소 글자수", `${pIn("memoMin", { label: "최소 글자수" })} 자 미만이면 부실`) +
        field("접대비 계정 판단", chipsIn("entWords"), "계정명에 이 단어가 있으면 상대방 표기를 확인합니다.") +
        field("상대방 표기로 인정", chipsIn("counterpart"), "'n인·n명'은 '3인', '4명'처럼 숫자+인/명, 'ㅇㅇㅇ'은 이름 가림 표기입니다.") +
        field("이 단어만 적으면 부실", chipsIn("bareWords"), "적요가 이 단어로만 되어 있으면(띄어쓰기 무시) 글자수와 관계없이 부실로 봅니다.");
      case "R09": return field("같다고 볼 항목", checks("dupKeys", ["카드", "가맹점", "금액", "날짜"]), "체크한 항목이 모두 같은 건이 2건 이상이면 중복결제로 봅니다.");
      case "R13": return field("제목 단어 ↔ 맞는 업종", `<table class="kw"><thead><tr><th>결재문서제목에 있으면</th><th>실제 업종이 이 중 하나여야 함 (쉼표 구분)</th><th></th></tr></thead><tbody>${P.titleMap.map(([k, v], i) => `<tr><td><input data-tm="${i}" data-tc="0" value="${esc(k)}" aria-label="제목 단어"></td><td><input data-tm="${i}" data-tc="1" value="${esc(v)}" aria-label="맞는 업종"></td><td><button type="button" class="x" data-tmdel="${i}" aria-label="행 삭제">×</button></td></tr>`).join("")}</tbody></table><button class="addline" type="button" data-tmadd>+ 행 추가</button>`) +
        field("근무지 식사를 여비교통비로", `계정명에 ${chipsIn("travelAcct")} 업종이 ${chipsIn("mealInds", "dlInd")}`, "적요에 '회식'이 있거나, 근무지 안인데 아래 출장 단어가 없으면 표시합니다.") +
        field("출장으로 인정할 단어", chipsIn("travelWords"));
      case "R14": return field("간격", `${pIn("distGap", { label: "간격" })} 분 이내, 시·도(주소 앞 2글자)가 다르면`) +
        field("제외 업종", chipsIn("distExclude", "dlInd"), "위치와 결제 장소가 다를 수 있는 업종(온라인·교통)은 비교하지 않습니다.");
      case "R16": return field("공용카드 판단", `구분 값에 ${chipsIn("sharedWords")} 이(가) 있으면 공용카드`) + field("기준금액", `${pIn("sharedMin", { w: 1, label: "기준금액" })} 원 이상`);
    }
    return "";
  }
  function ruleRow(id, isT) {
    const off = (D.params.off || []).includes(id), chg = ruleChanged(id, D, BASE), open = D.open.has(id);
    const w = D.weights[id] ?? 0, ww = BASE.weights[id] ?? 0;
    const name = isT ? ENGINE.T_NAMES[id] : ENGINE.RULE_NAMES[id];
    const more = DETAIL_RULES.includes(id) ? `<div><button class="more" type="button" data-open="${id}" aria-expanded="${open}">${open ? "세부 설정 접기 ▴" : "세부 설정 ▾"}${RULE_KEYS[id].some((k) => !same(D.params[k], BASE.params[k])) ? ` <span class="tag">변경</span>` : ""}</button></div>` : "";
    return `<tr class="${chg ? "chg" : ""} ${off ? "off" : ""} ${open ? "opened" : ""}">
      <td><button type="button" class="tg ${off ? "" : "on"}" role="switch" aria-checked="${!off}" aria-label="${id} 사용" data-off="${id}"></button></td>
      <td class="mono"><b>${id}</b></td>
      <td>${esc(name)}${off ? '<span class="tag">사용 안 함</span>' : ""}<div class="rh2">${esc(ENGINE.RULE_HELP[id] || "")}</div>${more}</td>
      <td class="cr">${critHtml(id)}</td>
      ${isT ? "" : `<td><span class="wt ${w !== ww ? "chg" : ""}"><button type="button" data-wd="${id}" data-delta="-1" aria-label="가중치 낮춤">−</button><span>${w}</span><button type="button" data-wd="${id}" data-delta="1" aria-label="가중치 높임">+</button></span>${w !== ww ? `<span class="was">${ww}</span>` : ""}</td>`}
      <td class="cnt2" data-cnt="${id}"></td></tr>` + (open ? `<tr class="drow"><td></td><td colspan="${isT ? 4 : 5}"><div class="dbox">${detailHtml(id)}</div></td></tr>` : "");
  }
  function renderEditor() {
    if (!D) return;
    const body = $("#edBody"), keep = body.scrollTop;
    const offU = D.custom.length;
    $("#edSub").textContent = `바꾼 값은 아래 [적용]을 눌러야 반영됩니다. 오른쪽 숫자는 ${S.month} 선별 건수(현재 → 바꾼 뒤 예상)입니다.`;
    $("#edTabs").innerHTML = [["R", "단건·다건 규칙", `R ${ENGINE.RULE_IDS.length}`], ["U", "추가조건", `U ${offU}`], ["T", "기간 규칙", `T ${T_IDS.length}`]]
      .map(([k, t, n]) => `<button type="button" role="tab" data-edtab="${k}" aria-selected="${D.tab === k}">${t}<span class="n">${n}</span></button>`).join("");
    const inds = [...new Set([...R.ref.industries.values()].map((x) => x.name).concat(R.prep.rows.map((r) => r.industry)).filter(Boolean))].sort();
    const dl = `<datalist id="dlInd">${inds.map((x) => `<option value="${esc(x)}">`).join("")}</datalist>`;
    if (D.tab === "R") {
      body.innerHTML = dl + `<table><thead><tr><th style="width:46px">사용</th><th style="width:48px">ID</th><th style="width:260px">점검항목</th><th>판단 기준 <span style="font-weight:400">(칸 안의 숫자를 바꿀 수 있음)</span></th><th style="width:118px">가중치</th><th style="width:96px;text-align:right">${+S.month.slice(5)}월 선별</th></tr></thead><tbody>${ENGINE.RULE_IDS.map((id) => ruleRow(id)).join("")}</tbody></table>
        <p class="hint">등급: 금지업종 → 상 · 6점 이상 상 · 3~5점 중 · 1~2점 하 · 휴일·심야만 걸리고 적요에 사전승인 문서번호가 있으면 규칙당 1점 감점. 판단 방식 자체를 바꾸려면 외부망에서 새 버전을 만들어야 합니다.</p>`;
    } else if (D.tab === "T") {
      body.innerHTML = `<table><thead><tr><th style="width:46px">사용</th><th style="width:48px">ID</th><th style="width:260px">점검항목</th><th>판단 기준</th><th style="width:96px;text-align:right">${+S.month.slice(5)}월 선별</th></tr></thead><tbody>${T_IDS.map((id) => ruleRow(id, true)).join("")}</tbody></table>
        <p class="hint">기간 규칙은 위험점수에 더하지 않고 목록에 표시만 합니다. 오른쪽 숫자는 ${esc(S.month)} 선별 목록 중 이 패턴에 해당하는 건수입니다.</p>`;
    } else body.innerHTML = customHtml();
    body.scrollTop = keep;
    renderEditorCounts();
  }
  function customHtml() {
    const errs = R.custom.errors;
    const cols = [...new Set([...R.columns, ...ENGINE.DERIVED])];
    const sel = (list, v, attrs, c) => `<select ${attrs} class="${c ? "chg" : ""}">${(list.includes(v) || !v ? list : [v, ...list]).map((x) => `<option ${x === v ? "selected" : ""}>${esc(x)}</option>`).join("")}</select>`;
    const card = (c, ui) => {
      const o = BASE.custom.find((x) => x.id === c.id), off = (D.params.off || []).includes(c.id), chg = customChanged(c, D, BASE);
      const lines = c.conds.map((x, li) => { const ox = o && o.conds[li];
        return `<div class="uc"><span>${li ? "그리고" : "조건"}</span>${sel(cols, x.col, `data-u="${ui}" data-ul="${li}" data-uf="col" aria-label="컬럼"`, !ox || ox.col !== x.col)}${sel(ENGINE.OPS, x.op, `data-u="${ui}" data-ul="${li}" data-uf="op" aria-label="비교"`, !ox || ox.op !== x.op)}<input class="${!ox || ox.val !== x.val ? "chg" : ""}" data-u="${ui}" data-ul="${li}" data-uf="val" value="${esc(x.op && ["이상", "이하", "초과", "미만"].includes(x.op) ? condVal(x) : x.val)}" aria-label="값"${["이상", "이하", "초과", "미만"].includes(x.op) ? " data-money" : ""}><button type="button" class="x" data-uldel="${ui}" data-ul="${li}" title="이 줄 삭제" aria-label="이 줄 삭제">×</button></div>`; }).join("");
      return `<div class="uc-card ${chg ? "chg" : ""} ${off ? "off" : ""}"><div class="uc-head"><button type="button" class="tg ${off ? "" : "on"}" role="switch" aria-checked="${!off}" aria-label="${c.id} 사용" data-off="${c.id}"></button><b class="mono">${c.id}</b><input class="nm" data-u="${ui}" data-uf="name" value="${esc(c.name)}" aria-label="조건명" style="border:1px solid var(--rule);border-radius:4px;padding:4px 6px;background:var(--sheet);color:var(--ink)">
        <span class="meta">가중치 <span class="wt ${(D.weights[c.id] ?? c.weight) !== (o ? BASE.weights[c.id] ?? o.weight : -1) ? "chg" : ""}"><button type="button" data-wd="${c.id}" data-delta="-1" aria-label="가중치 낮춤">−</button><span>${D.weights[c.id] ?? c.weight}</span><button type="button" data-wd="${c.id}" data-delta="1" aria-label="가중치 높임">+</button></span>
        <span class="cnt2" data-cnt="${c.id}"></span><button type="button" class="x" data-udel="${ui}" style="border:0;background:none;color:var(--ink-3)">삭제</button></span></div>
        ${lines}<button class="addline" type="button" data-uladd="${ui}">+ 조건 줄 추가 (그리고)</button></div>`;
    };
    const next = "U" + String(Math.max(0, ...D.custom.map((c) => +c.id.slice(1) || 0)) + 1).padStart(2, "0");
    return (errs.length ? `<div class="ed-err"><b>기준표 엑셀의 추가조건 시트에서 ${fmt(errs.length)}행을 적용하지 못했습니다.</b><br>${errs.map(esc).join("<br>")}<br>아래에서 새로 만들거나 엑셀을 고쳐 다시 불러오세요.</div>` : "") +
      (S.customRows ? `<p class="hint">화면에서 고친 추가조건을 쓰고 있습니다(기준표 엑셀의 추가조건 시트 대신). [기본값으로 되돌리기]를 누르면 기준표 시트 내용으로 돌아갑니다.</p>` : "") +
      `<p class="hint" style="margin-top:12px">같은 조건ID 안의 줄은 모두 만족해야 걸립니다. 한 건씩 판단하는 조건만 만들 수 있습니다. 숫자 비교(이상·이하·초과·미만)는 숫자로 적습니다.</p>` +
      D.custom.map(card).join("") + `<button class="addrule" type="button" data-uadd="${next}">+ 새 추가조건 만들기 (${next})</button>`;
  }
  function renderEditorCounts() {
    if (!D) return;
    const cur = monthCounts(R.list), hitAll = {};
    R.rows.forEach((r) => r.hits.forEach((h) => (hitAll[h] = (hitAll[h] || 0) + 1)));
    const sim = SIM ? monthCounts(SIM.list) : null, simAll = {};
    if (SIM) SIM.rows.forEach((r) => r.hits.forEach((h) => (simAll[h] = (simAll[h] || 0) + 1)));
    document.querySelectorAll("#edBody [data-cnt]").forEach((el) => {
      const id = el.dataset.cnt, n = cur[id] || 0, a = sim ? sim[id] || 0 : null;
      const tail = a === null ? `<span class="muted-note"> …</span>` : a !== n ? ` <span class="${a > n ? "up" : "dn"}">→ ${fmt(a)}</span>` : "";
      el.innerHTML = id.startsWith("U") ? `전체 적중 ${fmt(hitAll[id] || 0)}${sim && (simAll[id] || 0) !== (hitAll[id] || 0) ? ` → ${fmt(simAll[id] || 0)}` : ""} · ${+S.month.slice(5)}월 선별 ${fmt(n)}${tail}` : `${fmt(n)}${tail}`;
    });
    const g0 = gradeCounts(R.list), n0 = R.list.length;
    const n1 = SIM ? SIM.list.length : null, g1 = SIM ? gradeCounts(SIM.list) : null;
    $("#edSum").innerHTML = `<span>바꾼 항목 <b>${fmt(changeCount(D, BASE))}개</b> · ${esc(S.month)} 선별 <b>${fmt(n0)}건${n1 !== null && n1 !== n0 ? ` → ${fmt(n1)}건` : ""}</b> (상 ${fmt(g0.상)}${g1 && g1.상 !== g0.상 ? ` → ${fmt(g1.상)}` : ""})</span><small>적용하면 바로 다시 계산하고 이 PC 브라우저에 기억합니다. 다른 PC에서도 쓰려면 기준표 엑셀로 저장하세요.</small>`;
  }

  // 편집 이벤트
  const ed = $("#ruleEditor");
  function editChanged(rerender = true) { if (rerender) renderEditor(); else renderEditorCounts(); simulate(); }
  ed.addEventListener("click", (e) => {
    const t = e.target.closest("button"); if (!t || !D) return;
    const P = D.params;
    if (t.dataset.edtab) { D.tab = t.dataset.edtab; $("#edBody").scrollTop = 0; renderEditor(); return; }
    if (t.dataset.off) { const id = t.dataset.off; P.off = P.off.includes(id) ? P.off.filter((x) => x !== id) : [...P.off, id]; return editChanged(); }
    if (t.dataset.open) { D.open.has(t.dataset.open) ? D.open.delete(t.dataset.open) : D.open.add(t.dataset.open); return renderEditor(); }
    if (t.dataset.wd) { const id = t.dataset.wd; const c = D.custom.find((x) => x.id === id); const v = Math.min(10, Math.max(0, (D.weights[id] ?? c?.weight ?? 0) + +t.dataset.delta)); D.weights[id] = v; if (c) c.weight = v; return editChanged(); }
    if (t.dataset.ck) { P[t.dataset.ck] = P[t.dataset.ck].filter((_, i) => i !== +t.dataset.ci); return editChanged(); }
    if (t.dataset.tmdel !== undefined) { P.titleMap.splice(+t.dataset.tmdel, 1); return editChanged(); }
    if (t.hasAttribute("data-tmadd")) { P.titleMap.push(["", ""]); return renderEditor(); }
    if (t.dataset.uldel !== undefined) { const c = D.custom[+t.dataset.uldel]; c.conds.splice(+t.dataset.ul, 1); if (!c.conds.length) c.conds.push({ col: "가맹점업종", op: "같음", val: "" }); return editChanged(); }
    if (t.dataset.uladd !== undefined) { D.custom[+t.dataset.uladd].conds.push({ col: "전표금액", op: "이상", val: "" }); return renderEditor(); }
    if (t.dataset.udel !== undefined) { const c = D.custom.splice(+t.dataset.udel, 1)[0]; delete D.weights[c.id]; P.off = P.off.filter((x) => x !== c.id); return editChanged(); }
    if (t.dataset.uadd) { D.custom.push({ id: t.dataset.uadd, name: "새 조건", weight: 1, conds: [{ col: "가맹점업종", op: "같음", val: "" }] }); D.weights[t.dataset.uadd] = 1; return renderEditor(); }
  });
  ed.addEventListener("input", (e) => {
    const el = e.target; if (!D) return;
    if (el.dataset.money !== undefined && el.dataset.uf !== "val") commaInput(el);
    // 글자 입력 중에는 화면을 다시 그리지 않고(커서·클릭 유지) 그 칸의 변경 표시와 건수만 갱신
    if (el.dataset.p && el.tagName === "INPUT") { const v = unComma(el.value); if (el.value !== "" && isFinite(v)) { const k = el.dataset.p; D.params[k] = el.dataset.pct !== undefined ? v / 100 : v; el.classList.toggle("chg", !same(D.params[k], BASE.params[k])); editChanged(false); } return; }
    if (el.dataset.tm !== undefined) { D.params.titleMap[+el.dataset.tm][+el.dataset.tc] = el.value; return editChanged(false); }
    if (el.dataset.u !== undefined && el.tagName === "INPUT") {
      const c = D.custom[+el.dataset.u];
      if (el.dataset.uf === "name") c.name = el.value; else c.conds[+el.dataset.ul].val = el.value.trim();
      return editChanged(false);
    }
  });
  ed.addEventListener("change", (e) => {
    const el = e.target; if (!D) return;
    if (el.dataset.p && el.tagName === "SELECT") { D.params[el.dataset.p] = +el.value; return editChanged(); }
    if (el.dataset.cb) { const k = el.dataset.cb; D.params[k] = [...ed.querySelectorAll(`[data-cb="${k}"]`)].filter((x) => x.checked).map((x) => x.value); return editChanged(); }
    if (el.dataset.u !== undefined && el.tagName === "SELECT") { D.custom[+el.dataset.u].conds[+el.dataset.ul][el.dataset.uf] = el.value; return editChanged(); }
  });
  ed.addEventListener("keydown", (e) => {
    const el = e.target;
    if (e.key === "Escape") { closeEditor(); return; }
    if (e.key === "Enter" && el.dataset.cadd) {
      e.preventDefault(); const k = el.dataset.cadd, v = el.value.trim();
      if (v && !D.params[k].includes(v)) { D.params[k] = [...D.params[k], v]; editChanged(); setTimeout(() => ed.querySelector(`[data-cadd="${k}"]`)?.focus(), 0); }
      else el.value = "";
    }
  });
  ed.addEventListener("click", (e) => { if (e.target === ed) closeEditor(); });
  $("#edClose").addEventListener("click", closeEditor);
  $("#edCancel").addEventListener("click", closeEditor);
  $("#edDefault").addEventListener("click", () => {
    const sheetW = ENGINE.buildRef(S.refRaw).weights;
    const sheetCustom = ENGINE.parseCustom(S.refRaw.custom || [], R.columns).rules;
    D.params = clone(ENGINE.DEFAULT_PARAMS);
    D.weights = { ...sheetW }; sheetCustom.forEach((c) => (D.weights[c.id] = c.weight));
    D.custom = sheetCustom.map((c) => ({ id: c.id, name: c.name, weight: c.weight, conds: c.conds.map((x) => ({ col: x.col, op: x.op, val: x.val })) }));
    D.resetCustom = true;
    editChanged();
  });
  $("#edApply").addEventListener("click", () => {
    if (!D) return;
    S.params = clone(D.params);
    S.weightOverride = { ...D.weights };
    const rows = customToRows(D.custom);
    S.customRows = D.resetCustom && same(rows, customToRows(ENGINE.parseCustom(S.refRaw.custom || [], R.columns).rules.map((c) => ({ ...c, weight: D.weights[c.id] ?? c.weight })))) ? null : rows;
    saveRuleConfig();
    closeEditor(); run();
    flash("#editBtn", "적용했습니다");
  });
  $("#edExport").addEventListener("click", () => {
    if (!D) return;
    const wb = XLSX.utils.book_new();
    for (const [k, name] of Object.entries(REF_SHEETS)) {
      let data = S.refRaw[k] && S.refRaw[k].length ? S.refRaw[k] : DEFAULT_REF[k];
      if (k === "rules") { const ids = new Set(data.map((r) => String(r["규칙ID"]))); data = data.map((r) => (r["규칙ID"] in D.weights ? { ...r, 위험가중치: D.weights[r["규칙ID"]] } : r)); ENGINE.RULE_IDS.filter((id) => !ids.has(id)).forEach((id) => data.push({ 규칙ID: id, 점검항목: ENGINE.RULE_NAMES[id], 위험가중치: D.weights[id] })); }
      if (k === "custom") data = customToRows(D.custom);
      XLSX.utils.book_append_sheet(wb, commaSheet(XLSX.utils.json_to_sheet(data.length ? data : [{}])), name);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(configRows(D.params, D.weights)), "규칙설정");
    saveFile(wb, "card-audit_reference.xlsx");
    flash("#edExport", "저장했습니다");
  });
  // '규칙설정' 시트: 사람이 읽는 설명 + 다시 읽을 값(JSON). 이 시트가 있는 기준표를 불러오면 같은 설정으로 복원
  function configRows(params, weights) {
    return [["항목", "값", "설명"],
      ["params", JSON.stringify(params), "기준값·세부 설정·사용 안 함 규칙 (점검 규칙 수정 화면에서 만든 값, 직접 고치지 마세요)"],
      ["weights", JSON.stringify(weights), "규칙별 가중치"],
      ["저장일시", new Date().toLocaleString("ko-KR"), ""]];
  }
  function readConfigSheet(rows) {
    const get = (k) => { const r = rows.find((x) => x["항목"] === k); try { return r ? JSON.parse(r["값"]) : null; } catch { return null; } };
    const p = get("params"), w = get("weights");
    if (p) S.params = { ...clone(ENGINE.DEFAULT_PARAMS), ...p };
    if (w) S.weightOverride = w;
    S.customRows = null; // 추가조건은 같은 파일의 추가조건 시트를 씀
    saveRuleConfig();
  }
  function saveRuleConfig() { store.set("cc_rules", { params: S.params, weights: S.weightOverride || null, customRows: S.customRows || null }); }
  // 결과 엑셀 '적용 기준' 시트: 그달 어떤 기준으로 점검했는지 증적
  function appliedRows() {
    const def = ENGINE.DEFAULT_PARAMS, sheetW = ENGINE.buildRef(S.refRaw).weights, P = S.params;
    const diff = (id) => (RULE_KEYS[id] || []).filter((k) => !same(P[k], def[k])).map((k) => `${k}: ${Array.isArray(P[k]) ? JSON.stringify(P[k]) : P[k]} (기본 ${Array.isArray(def[k]) ? JSON.stringify(def[k]) : def[k]})`).join("; ");
    const out = ENGINE.RULE_IDS.map((id) => ({ 구분: "단건·다건", 규칙ID: id, 점검항목: ENGINE.RULE_NAMES[id], 사용: P.off.includes(id) ? "사용 안 함" : "사용", 가중치: S.weights[id] ?? "", "기준표 가중치": sheetW[id] ?? "", 판단기준: ruleDesc(id), "기본값과 다른 설정": [diff(id), (S.weights[id] ?? 0) !== (sheetW[id] ?? 0) ? `가중치 ${S.weights[id]} (기준표 ${sheetW[id]})` : ""].filter(Boolean).join("; ") }));
    R.custom.rules.forEach((c) => out.push({ 구분: "추가조건", 규칙ID: c.id, 점검항목: c.name, 사용: P.off.includes(c.id) ? "사용 안 함" : "사용", 가중치: S.weights[c.id] ?? c.weight, "기준표 가중치": "", 판단기준: ruleDesc(c.id), "기본값과 다른 설정": S.customRows ? "화면에서 수정" : "" }));
    T_IDS.forEach((id) => out.push({ 구분: "기간", 규칙ID: id, 점검항목: ENGINE.T_NAMES[id], 사용: P.off.includes(id) ? "사용 안 함" : "사용", 가중치: "점수 없음", "기준표 가중치": "", 판단기준: ruleDesc(id), "기본값과 다른 설정": diff(id) }));
    return out;
  }
  $("#editBtn").addEventListener("click", () => openEditor());
  $("#ruleEdit").addEventListener("click", () => { closeRules(); openEditor(T_IDS.includes(S.ruleFocus) ? "T" : S.ruleFocus?.startsWith("U") ? "U" : "R"); });

  // ── 거래 조회 ──
  const B = { period: "month", q: "", gl: "", ind: "", dept: "", min: "", max: "", sel: "", hol: false, cancel: false };
  function renderBrowseFilters() {
    const all = R.prep.rows;
    const opts = (key) => ["", ...[...new Set(all.map((r) => r[key]).filter(Boolean))].sort()].map((v) => `<option value="${esc(v)}" ${B[{ glName: "gl", industry: "ind", dept: "dept" }[key]] === v ? "selected" : ""}>${v ? esc(v) : "전체"}</option>`).join("");
    $("#b_gl").innerHTML = opts("glName"); $("#b_ind").innerHTML = opts("industry"); $("#b_dept").innerHTML = opts("dept");
  }
  function browseRows() {
    const q = B.q.trim();
    return R.prep.rows.filter((r) => {
      if (B.period === "month" && r.month !== S.month) return false;
      if (!B.cancel && r.excluded) return false;
      if (B.gl && r.glName !== B.gl) return false;
      if (B.ind && r.industry !== B.ind) return false;
      if (B.dept && r.dept !== B.dept) return false;
      if (B.min !== "" && r.amount < unComma(B.min)) return false;
      if (B.max !== "" && r.amount > unComma(B.max)) return false;
      if (B.hol && !r.holiday) return false;
      const flagged = !r.excluded && r.score > 0;
      if (B.sel === "y" && !flagged) return false;
      if (B.sel === "n" && flagged) return false;
      if (q && ![r.merchant, r.memo, r.emp, r.title, r.no, r.card].some((x) => String(x).includes(q))) return false;
      return true;
    }).sort((a, b) => a.dt - b.dt);
  }
  function renderBrowse() {
    const rows = browseRows();
    const total = rows.reduce((s, r) => s + r.amount, 0);
    const shown = rows.slice(0, 500);
    $("#b_summary").innerHTML = `<b>${fmt(rows.length)}건</b> · 합계 ${won(total)}${rows.length > 500 ? " · 화면에는 앞 500건만 표시하고, 엑셀 저장에는 전부 들어갑니다" : ""}`;
    $("#browseTbl").innerHTML = `<thead><tr><th>등급</th><th>규칙</th><th>승인일시</th><th>직원 · 부서</th><th>가맹점</th><th>계정</th><th style="text-align:right">금액</th><th>적요</th></tr></thead><tbody>` +
      (shown.length ? shown.map((r) => `<tr class="${r.grade ? "sev-" + r.grade : ""}"><td>${r.excluded ? '<span class="muted">상계</span>' : r.grade ? `<span class="g g-${r.grade}">${r.grade}</span>` : ""}</td>
        <td><div class="chips">${r.excluded ? "" : ruleChips(r)}</div></td><td class="mono" style="white-space:nowrap">${r.dkey.slice(5)} ${r.time}<div class="muted">${r.holName || r.wdName}</div></td>
        <td style="white-space:nowrap">${esc(who(r))}<div class="muted">${esc(r.dept)}</div></td><td>${esc(r.merchant)}<div class="muted">${esc(r.industry)}</div></td>
        <td>${esc(r.glName)}</td><td class="n">${fmt(r.amount)}</td><td class="reason">${esc(r.memo)}</td></tr>`).join("")
        : `<tr><td colspan="8" class="empty">조건에 맞는 거래가 없습니다.</td></tr>`) + "</tbody>";
  }
  function browseExport() {
    const rows = browseRows();
    const cols = R.columns;
    const data = rows.map((r) => {
      const o = {};
      for (const c of cols) o[c] = r.raw[c] ?? "";
      o["등급"] = r.excluded ? "취소상계" : r.grade || "";
      o["위험점수"] = r.excluded ? "" : r.score;
      o["걸린 규칙"] = r.excluded ? "" : [...r.hits].sort().map((h) => `${h} ${ENGINE.ruleName(h)}`).concat([...r.tHits].map((h) => `${h} ${ENGINE.T_NAMES[h]}`)).join(" / ");
      return o;
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, commaSheet(XLSX.utils.json_to_sheet(data.length ? data : [{ 안내: "조건에 맞는 거래가 없습니다" }])), "거래조회");
    const cond = [["조회 조건"], ["기간", B.period === "month" ? S.month : "전체 기간"], ["검색어", B.q], ["계정", B.gl || "전체"], ["업종", B.ind || "전체"], ["부서", B.dept || "전체"],
      ["금액", `${B.min ? fmt(unComma(B.min)) : "-"} ~ ${B.max ? fmt(unComma(B.max)) : "-"}`], ["선별", { "": "전체", y: "선별된 건만", n: "선별 안 된 건만" }[B.sel]], ["휴일만", B.hol ? "예" : "아니오"], ["취소·상계 포함", B.cancel ? "예" : "아니오"], ["건수", fmt(rows.length)]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cond), "조회조건");
    saveFile(wb, `card-audit_browse_${B.period === "month" ? S.month : "all"}.xlsx`);
    flash("#b_xlsx", "저장했습니다");
  }
  $("#browseFilters").addEventListener("input", (e) => {
    const k = e.target.dataset.b; if (!k) return;
    if (e.target.dataset.money !== undefined) commaInput(e.target);
    B[k] = e.target.type === "checkbox" ? e.target.checked : e.target.value;
    renderBrowse();
  });
  $("#b_reset").addEventListener("click", () => {
    Object.assign(B, { period: "month", q: "", gl: "", ind: "", dept: "", min: "", max: "", sel: "", hol: false, cancel: false });
    document.querySelectorAll("#browseFilters [data-b]").forEach((el) => { if (el.type === "checkbox") el.checked = false; else el.value = B[el.dataset.b]; });
    renderBrowseFilters(); renderBrowse();
  });
  $("#b_xlsx").addEventListener("click", browseExport);

  // ── 이벤트 ──
  $("#tabs").addEventListener("click", (e) => { const b = e.target.closest(".tab"); if (!b) return; S.tab = b.dataset.tab; store.set("cc_tab", S.tab); renderTabs(); });
  $("#month").addEventListener("change", (e) => { S.month = e.target.value; S.sel = null; run(); });
  $("#reviewFilter").addEventListener("click", (e) => { const b = e.target.closest("button[data-r]"); if (!b) return; S.review = b.dataset.r; renderList(); });
  $("#gradeFilter").addEventListener("click", (e) => { const b = e.target.closest("button"); if (!b) return; if (b.dataset.clearRule) { S.ruleFilter = null; renderList(); renderRuleDrawer(); return; } S.grade = b.dataset.g; renderList(); });
  $("#listTbl").addEventListener("click", (e) => { const tr = e.target.closest("tr[data-no]"); if (!tr) return; S.sel = S.sel === tr.dataset.no ? null : tr.dataset.no; renderList(); if (S.sel) $("#detail").scrollIntoView({ block: "nearest", behavior: "smooth" }); });
  $("#listTbl").addEventListener("keydown", (e) => { if (e.key === "Enter") e.target.closest("tr[data-no]")?.click(); });
  $("#detail").addEventListener("click", (e) => {
    const b = e.target.closest("#seg button");
    if (b) { const rv = (S.reviews[S.sel] ||= {}); rv.result = rv.result === b.dataset.v ? "" : b.dataset.v; store.set("cc_reviews", S.reviews); renderList(); return; }
  });
  $("#detail").addEventListener("input", (e) => { if (e.target.id !== "opinion") return; (S.reviews[S.sel] ||= {}).opinion = e.target.value; store.set("cc_reviews", S.reviews); });
  $("#uploadBtn").addEventListener("click", () => $("#file").click());
  $("#resetBtn").addEventListener("click", () => {
    if ($("#resetBtn").dataset.confirm !== "1") { $("#resetBtn").dataset.confirm = "1"; $("#resetBtn").textContent = "한 번 더 누르면 초기화"; setTimeout(() => { $("#resetBtn").dataset.confirm = ""; $("#resetBtn").textContent = "불러온 파일 비우기"; }, 3000); return; }
    $("#resetBtn").dataset.confirm = ""; $("#resetBtn").textContent = "불러온 파일 비우기";
    Object.assign(S, { source: "none", files: [], raw: new Map(), sheetRows: 0, readCount: 0, refRaw: {}, refFiles: {}, answers: null, month: null, sel: null, weights: null });
    closeEditor();
    R = null; renderEmpty();
  });
  $("#templateBtn").addEventListener("click", () => {
    const wb = XLSX.utils.book_new();
    const guide = [["기준표 양식 안내"], ["이 파일의 값은 교육용 가상값입니다. 내부망에서 실제 사내 기준으로 바꿔 쓰세요."], ["시트 이름은 바꾸지 마세요. 컬럼 순서는 바뀌어도 되지만 컬럼명은 그대로 두세요."],
      ["계정코드: G/L 계정, G/L 계정명, 허용 가맹점업종(쉼표 구분), 건당 한도(원)"], ["업종코드: 업종코드, 가맹점업종, 제한구분(금지/사전승인/주의/일반)"], ["공휴일: 일자, 명칭 (주말은 자동 계산)"],
      ["점검기준: 규칙ID, 점검항목, 위험가중치"], ["직원마스터: 사번, 이름, 부서, 근무지 (R13 근무지 판단)"], ["휴가내역: 사번, 이름, 휴가구분, 시작일, 종료일 (R17)"],
      [""], ["추가조건: 반입 없이 내부망에서 점검 조건을 추가하는 시트"], ["- 조건ID는 U01, U02… 처럼 적고, 같은 조건ID의 행은 모두 만족해야 걸립니다(AND). 조건명·가중치는 그 ID의 첫 행에 적습니다."],
      ["- 컬럼: SAP 엑셀의 컬럼명 그대로(예: 가맹점업종, 전표금액, 전표 품목 텍스트, 결재문서제목) 또는 계산 컬럼(" + ENGINE.DERIVED.join(", ") + ")"],
      ["- 비교: " + ENGINE.OPS.join(", ")], ["- 목록중하나는 값에 쉼표로 여러 개를 적습니다(예: 토,일). 승인시는 0~23 숫자, 휴일은 Y/N입니다."],
      ["- 예시 2개는 실제 기준이 아니니 지우거나 바꿔 쓰세요."]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(guide), "안내");
    for (const [k, name] of Object.entries(REF_SHEETS)) XLSX.utils.book_append_sheet(wb, commaSheet(XLSX.utils.json_to_sheet(DEFAULT_REF[k])), name);
    saveFile(wb, "card-audit_reference-template.xlsx");
  });
  $("#file").addEventListener("change", (e) => { if (e.target.files.length) loadFiles([...e.target.files]); e.target.value = ""; });
  $("#copyBtn").addEventListener("click", async () => {
    const tsv = exportRows().map((r) => Object.values(r).map((v) => String(v ?? "").replace(/[\t\n]/g, " ")).join("\t"));
    const text = [Object.keys(exportRows()[0] || {}).join("\t"), ...tsv].join("\n");
    try { await navigator.clipboard.writeText(text); flash("#copyBtn", "복사했습니다"); } catch { flash("#copyBtn", "복사할 수 없습니다"); }
  });
  // ── 소명 요청 메일: 검토결과 '소명요청' 건을 사람별로 묶어 메일 제목·본문을 만듦 ──
  const WD = ["일", "월", "화", "수", "목", "금", "토"];
  const pad2 = (n) => String(n).padStart(2, "0");
  const localYmd = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  const withWd = (s) => { const [y, m, d] = s.split("-").map(Number); return s ? `${s}(${WD[new Date(y, m - 1, d).getDay()]})` : ""; };
  // 다음 영업일: 주말과 기준표 공휴일 시트의 날을 건너뜀
  function nextBizDay(from) {
    const d = new Date(from.getFullYear(), from.getMonth(), from.getDate());
    do d.setDate(d.getDate() + 1); while (d.getDay() === 0 || d.getDay() === 6 || R.ref.holidays.has(localYmd(d)));
    return localYmd(d);
  }
  function mailGroups() {
    const map = new Map();
    for (const r of R.list) {
      if (S.reviews[r.no]?.result !== "소명요청") continue;
      // 직원명 공란(공용카드)은 전표 작성자(참조)에게 보내므로 작성자·부서별로 묶음
      const key = r.emp ? `e:${r.emp}` : `w:${r.writer}|${r.dept}`;
      if (!map.has(key)) map.set(key, { key, emp: r.emp || "", dept: r.dept, items: [] });
      map.get(key).items.push(r);
    }
    return [...map.values()].sort((a, b) => (!a.emp - !b.emp) || a.emp.localeCompare(b.emp, "ko") || a.dept.localeCompare(b.dept, "ko"));
  }
  const M = { groups: [], sel: 0, edits: {}, done: new Set(), due: "", today: "" };
  const monthKo = () => `${S.month.slice(0, 4)}년 ${Number(S.month.slice(5))}월`;
  const mailName = (g) => g.emp || g.dept; // 제목 괄호: 직원명, 공란이면 부서명
  const mailSubject = (g) => `[감사] ${monthKo()} 법인카드 사용내역 소명 요청 (${mailName(g)}, ${g.items.length}건)`;
  // 메일 프로그램에 붙여도 모양이 유지되도록 서식은 모두 인라인으로 넣음
  function mailBody(g) {
    const sender = esc(store.get("cc_sender", "") || "○○○");
    const F = "font-family:'맑은 고딕','Malgun Gothic',sans-serif;font-size:10pt;color:#1c2320;";
    const th = `style="${F}background:#e2edf0;color:#1d5a6e;font-weight:bold;text-align:left;padding:6px 8px;border:1px solid #b9c4c0"`;
    const td = (x = "") => `style="${F}padding:6px 8px;border:1px solid #b9c4c0;vertical-align:top;${x}"`;
    const p = `style="${F}margin:0 0 12px;line-height:1.6"`;
    const rows = g.items.map((r, i) => `<tr><td ${td("text-align:center")}>${i + 1}</td><td ${td("white-space:nowrap")}>${r.dkey} ${r.time.slice(0, 5)}</td><td ${td()}>${esc(who(r))}</td><td ${td()}>${esc(r.merchant)}</td><td ${td("text-align:right;white-space:nowrap")}>${fmt(r.amount)}</td><td ${td()}>${esc(r.glName)}</td><td ${td()}>${r.memo ? esc(r.memo) : "(공란)"}</td><td ${td("min-width:130px")}>${esc(ENGINE.reasonShort(r))}</td><td ${td("min-width:130px")}>${r.checks.map(esc).join("<br>")}</td><td ${td("min-width:60px")}></td></tr>`).join("");
    return `<p ${p}>안녕하세요,<br>감사부 ${sender}입니다.</p>`
      + `<p ${p}>${monthKo()} 법인카드 사용 적정성 점검 결과, 아래 ${g.items.length}건의 사용 내역에 대해 확인이 필요하여 소명을 요청드립니다.</p>`
      + `<table style="border-collapse:collapse;width:100%;margin:4px 0 14px" cellpadding="0" cellspacing="0"><tr>${["No", "승인일시", "직원(참조)", "가맹점", "금액(원)", "계정명", "전표 품목 텍스트", "선별 사유", "확인 요청 사항", "비고"].map((h) => `<th ${th}>${h}</th>`).join("")}</tr>${rows}</table>`
      + `<p ${p}>• <b>회신 기한</b>: ${withWd(M.due)}<br>• <b>회신 방법</b>: 건별 확인 요청 사항에 대한 답변 및 증빙(영수증, 품의서 등)을 첨부하여 메일 회신 요청드립니다.</p>`
      + `<p ${p}>감사합니다.<br>${sender} 드림</p>`;
  }
  function renderMailBtn() {
    const g = mailGroups(), n = g.reduce((a, x) => a + x.items.length, 0);
    const b = $("#mailBtn");
    b.textContent = `✉ 소명 요청 메일 만들기 (${g.length}명 · ${n}건)`;
    b.disabled = !n;
    b.title = n ? "검토결과 '소명요청' 건을 사람별로 묶어 메일 제목·본문을 만듭니다" : "검토결과를 '소명요청'으로 고른 건이 없습니다";
  }
  function openMail() {
    M.groups = mailGroups(); if (!M.groups.length) return;
    const now = new Date();
    M.sel = 0; M.edits = {}; M.done = new Set(); M.today = localYmd(now); M.due = nextBizDay(now);
    const n = M.groups.reduce((a, x) => a + x.items.length, 0);
    $("#mailModal").innerHTML = `<div class="mm" role="dialog" aria-modal="true" aria-labelledby="mmTitle">
      <div class="mm-h"><h2 id="mmTitle">소명 요청 메일 만들기</h2><span class="sub">${esc(S.month)} 점검 · 검토결과 '소명요청' ${n}건을 사용자별로 묶었습니다</span><span class="sp"></span><button class="btn" id="mmClose" type="button">닫기</button></div>
      <div class="mm-opt"><span>회신 기한 <input type="date" id="mmDue" value="${M.due}"> <span class="muted">(요청일 ${M.today} + 1영업일, 주말·공휴일 제외)</span></span><span>보내는 사람 <input id="mmSender" value="${esc(store.get("cc_sender", ""))}" placeholder="이름" style="width:110px"> <span class="muted">(인사말·서명에 들어감, 이 PC에 기억)</span></span></div>
      <div class="mm-b"><div class="mm-l" id="mmList"></div>
        <div class="mm-r">
          <div class="mm-row"><label for="mmSubj">제목</label><input class="mm-subj" id="mmSubj"><button class="btn" id="mmSubjCopy" type="button">제목 복사</button></div>
          <div class="mm-edit"><span>✎ 제목·본문을 눌러 바로 고칠 수 있습니다 (창을 닫기 전까지 사람별로 유지)</span><button class="btn" id="mmReset" type="button">처음 문구로 되돌리기</button></div>
          <div class="mm-paper" id="mmBody" contenteditable="true"></div>
        </div></div>
      <div class="mm-f"><span class="hint"></span><label><input type="checkbox" id="mmRecord" checked> <span id="mmRecordTxt"></span></label><button class="btn primary" id="mmCopy" type="button">본문 복사</button></div>
    </div>`;
    $("#mailModal").hidden = false;
    renderMail();
    $("#mmClose").focus();
  }
  function renderMail() {
    $("#mmList").innerHTML = M.groups.map((g, i) => `<button type="button" class="who ${i === M.sel ? "on" : ""} ${M.done.has(g.key) ? "done" : ""} ${g.emp ? "" : "warn"}" data-i="${i}"><span><span class="nm">${esc(`${g.emp || "부서공용"}(${[...new Set(g.items.map((r) => r.writer).filter(Boolean))].join(", ") || "참조 없음"})`)}</span><small>${esc(g.dept)}</small></span><span class="n">${g.items.length}건</span></button>`).join("");
    const g = M.groups[M.sel], e = M.edits[g.key] || {};
    $("#mmSubj").value = e.subj ?? mailSubject(g);
    $("#mmBody").innerHTML = e.body ?? mailBody(g);
    $("#mmRecordTxt").textContent = "본문 복사시 결과 엑셀에 소명요청일·회신기한 자동 입력";
  }
  const closeMail = () => { $("#mailModal").hidden = true; $("#mailModal").innerHTML = ""; $("#mailBtn").focus(); };
  async function copyRich(el) {
    const html = `<div style="font-family:'맑은 고딕','Malgun Gothic',sans-serif;font-size:10pt">${el.innerHTML}</div>`;
    try {
      await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([el.innerText], { type: "text/plain" }) })]);
      return true;
    } catch {
      // ClipboardItem을 못 쓰는 브라우저: 본문을 선택해 복사
      try { const sel = getSelection(), rg = document.createRange(); rg.selectNodeContents(el); sel.removeAllRanges(); sel.addRange(rg); const ok = document.execCommand("copy"); sel.removeAllRanges(); return ok; } catch { return false; }
    }
  }
  $("#mailBtn").addEventListener("click", openMail);
  $("#mailModal").addEventListener("click", async (e) => {
    const t = e.target;
    const who = t.closest(".who"); if (who) { M.sel = Number(who.dataset.i); renderMail(); return; }
    if (t.id === "mmClose") { closeMail(); return; }
    const g = M.groups[M.sel];
    if (t.id === "mmReset") { delete M.edits[g.key]; renderMail(); return; }
    if (t.id === "mmSubjCopy") { try { await navigator.clipboard.writeText($("#mmSubj").value); flash("#mmSubjCopy", "복사했습니다"); } catch { flash("#mmSubjCopy", "복사할 수 없습니다"); } return; }
    if (t.id === "mmCopy") {
      if (!(await copyRich($("#mmBody")))) { flash("#mmCopy", "복사할 수 없습니다"); return; }
      if ($("#mmRecord").checked) {
        for (const r of g.items) { const rv = (S.reviews[r.no] ||= {}); rv.sm = { ...(rv.sm || {}), 소명요청일: M.today, 회신기한: M.due }; }
        store.set("cc_reviews", S.reviews);
      }
      M.done.add(g.key); renderMail(); flash("#mmCopy", "복사했습니다");
    }
  });
  $("#mailModal").addEventListener("input", (e) => {
    const t = e.target, g = M.groups[M.sel];
    if (t.id === "mmSubj") (M.edits[g.key] ||= {}).subj = t.value;
    else if (t.id === "mmBody") (M.edits[g.key] ||= {}).body = t.innerHTML;
    else if (t.id === "mmSender") { store.set("cc_sender", t.value.trim()); if (!M.edits[g.key]?.body) $("#mmBody").innerHTML = mailBody(g); }
    else if (t.id === "mmDue" && t.value) { M.due = t.value; if (!M.edits[g.key]?.body) $("#mmBody").innerHTML = mailBody(g); }
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !$("#mailModal").hidden) closeMail(); });
  $("#xlsxBtn").addEventListener("click", () => {
    const wb = XLSX.utils.book_new();
    const rows = exportRows();
    XLSX.utils.book_append_sheet(wb, summarySheet(rows), "요약");
    XLSX.utils.book_append_sheet(wb, commaSheet(XLSX.utils.json_to_sheet(rows)), "점검대상목록");
    const per = [];
    for (const [k, list] of Object.entries(R.period)) for (const x of list) per.push({ 패턴: `${k} ${ENGINE.T_NAMES[k]}`, 대상: x.who, 부서: x.dept || "", 가맹점: x.merchant || "", 기준월: x.month || "", 지표: k === "T-R1" ? `3개월 ${x.n}회, ${fmt(x.amount)}원 ${x.focus || ""}` : k === "T-R2" ? `${fmt(x.cur)}원 / 직전 6개월 평균 ${fmt(x.avg)}원 (×${x.ratio.toFixed(1)})` : k === "T-R3" ? `월말 5영업일 ${(x.share * 100).toFixed(0)}%, 부서 평균의 ${(x.total / x.avg).toFixed(1)}배` : `${fmt(x.amount)}원 ${x.months}개월 연속`, 관련건수: x.rows.length });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(per.length ? per : [{ 패턴: "해당 없음" }]), "기간패턴");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(R.recs.map((r) => ({ 추천항목: r.title, 근거: r.basis, 이유: r.why, 규칙안: r.rule }))), "추가점검항목");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(appliedRows()), "적용기준");
    const ci = Object.keys(rows[0] || {}).indexOf("소명결과");
    saveFile(wb, `card-audit_result_${S.month}.xlsx`, (zip) => { styleSummary(zip, 0); if (ci >= 0) addListValidation(zip, 1, XLSX.utils.encode_col(ci), rows.length + 1, SM_RESULTS, "소명요청 건의 회신 결과를 고르세요: 소명 인정 / 부적정 전환 / 미회신"); });
    flash("#xlsxBtn", "저장했습니다");
  });
  // 소명 관리 열: 결과 엑셀에 빈칸으로 만들어 두고, 다시 불러오면 적은 내용을 되살림
  const SM_COLS = ["소명요청일", "회신기한", "회신일", "소명결과", "소명 요지", "증빙번호"];
  const SM_RESULTS = ["소명 인정", "부적정 전환", "미회신"];
  // SheetJS 무료판은 데이터 유효성 검사를 쓰지 못하므로, 저장된 xlsx(zip) 안의 시트 XML에 목록 검사를 직접 넣는다
  function addListValidation(zip, sheetIdx, col, lastRow, items, prompt) {
    const ent = XLSX.CFB.find(zip, `/xl/worksheets/sheet${sheetIdx + 1}.xml`);
    if (!ent) return;
    let xml = new TextDecoder().decode(ent.content);
    const dv = `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorStyle="warning" promptTitle="소명결과" prompt="${prompt}" errorTitle="소명결과" error="${items.join(", ")} 중에서 고르세요." sqref="${col}2:${col}${Math.max(lastRow, 2)}"><formula1>"${items.join(",")}"</formula1></dataValidation></dataValidations>`;
    // 스키마 순서상 sheetData·mergeCells보다 뒤, hyperlinks·pageMargins·ignoredErrors 등보다 앞에 와야 엑셀이 '복구' 경고 없이 연다
    const at = xml.search(/<(hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[\s>/]/);
    xml = at >= 0 ? xml.slice(0, at) + dv + xml.slice(at) : xml.replace("</worksheet>", dv + "</worksheet>");
    setXml(ent, xml);
  }
  const getXml = (ent) => new TextDecoder().decode(ent.content);
  const setXml = (ent, xml) => { ent.content = new TextEncoder().encode(xml); ent.size = ent.content.length; };
  // 요약 시트 서식: SheetJS 무료판은 글꼴·색·테두리를 저장하지 못하므로 styles.xml에 서식을 덧붙이고 셀에 번호(s)를 단다
  let SUMMARY_KINDS = [];
  const C = { ink: "1C2320", sub: "6B7570", accent: "1D5A6E", headBg: "E2EDF0", headLine: "B9CDD3", line: "E3E6E1", ng: "B3261E", wa: "9A5B00", ok: "2F6B3A", mu: "8A948F",
    rv: { 적정: ["EEF6EF", "2F6B3A"], 소명요청: ["EAF2F5", "1D5A6E"], 부적정: ["FBECEB", "B3261E"], 미검토: ["FFFFFF", "4E5A55"] } };
  function summaryStyle(kind, col, tag) {
    const num = col > 0, rv = kind.startsWith("rv-") ? C.rv[kind.slice(3)] : null;
    const font = { sz: 10, color: C.ink, b: false };
    let fill = null, border = {}, h = num ? "right" : "left", fmt = tag === "pct" ? "pct" : num ? "num" : null;
    if (kind === "title") Object.assign(font, { sz: 15, b: true });
    else if (kind === "sub" || kind === "note") Object.assign(font, { sz: 9, color: C.sub });
    else if (kind === "sec") Object.assign(font, { sz: 11.5, b: true, color: C.accent });
    else if (kind === "head") { Object.assign(font, { b: true, color: C.accent }); fill = C.headBg; border = { bottom: ["thin", C.headLine] }; fmt = null; }
    else if (kind === "total") { font.b = true; border = { top: ["medium", C.ink] }; }
    else { border = { bottom: ["thin", C.line] }; if (rv) { fill = rv[0]; if (!num) Object.assign(font, { b: true, color: rv[1] }); } }
    if (kind === "body" || rv) {
      if (tag === "fl") Object.assign(font, { b: true, color: C.accent });
      else if (tag === "ng") Object.assign(font, { b: true, color: C.ng });
      else if (tag === "wa") Object.assign(font, { b: true, color: C.wa });
      else if (tag === "mu") font.color = C.mu;
    }
    if (kind === "title" || kind === "sub" || kind === "note" || kind === "sec") h = "left";
    return { font, fill, border, h, fmt };
  }
  function styleSummary(zip, sheetIdx) {
    const st = XLSX.CFB.find(zip, "/xl/styles.xml"), sh = XLSX.CFB.find(zip, `/xl/worksheets/sheet${sheetIdx + 1}.xml`);
    if (!st || !sh) return;
    let css = getXml(st), xml = getXml(sh);
    const count = (tag, re) => ((css.match(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`)) || [, ""])[1].match(re) || []).length;
    const base = { font: count("fonts", /<font[ >/]/g), fill: count("fills", /<fill[ >/]/g), border: count("borders", /<border[ >/]/g), xf: count("cellXfs", /<xf /g) };
    const fonts = [], fills = [], borders = [], xfs = [], memo = {};
    const idx = (arr, x) => { const k = arr.indexOf(x); return k >= 0 ? k : arr.push(x) - 1; };
    const NUM = { num: 200, pct: 201 };
    const xfOf = (s) => {
      const key = JSON.stringify(s);
      if (key in memo) return memo[key];
      const f = base.font + idx(fonts, `<font>${s.font.b ? "<b/>" : ""}<sz val="${s.font.sz}"/><color rgb="FF${s.font.color}"/><name val="맑은 고딕"/><family val="3"/><charset val="129"/></font>`);
      const fi = s.fill ? base.fill + idx(fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${s.fill}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
      const side = (n) => (s.border[n] ? `<${n} style="${s.border[n][0]}"><color rgb="FF${s.border[n][1]}"/></${n}>` : `<${n}/>`);
      const bo = Object.keys(s.border).length ? base.border + idx(borders, `<border>${side("left")}${side("right")}${side("top")}${side("bottom")}<diagonal/></border>`) : 0;
      const nf = s.fmt ? NUM[s.fmt] : 0;
      return (memo[key] = base.xf + idx(xfs, `<xf numFmtId="${nf}" fontId="${f}" fillId="${fi}" borderId="${bo}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="${s.h}" vertical="center"/></xf>`));
    };
    // 셀마다 서식 번호를 붙임
    xml = xml.replace(/<c r="([A-Z]+)(\d+)"( s="\d+")?/g, (m, c, r, sOld) => {
      const [kind, tags] = SUMMARY_KINDS[r - 1] || ["body", []];
      if (kind === "blank") return m;
      const col = XLSX.utils.decode_col(c);
      return `<c r="${c}${r}" s="${xfOf(summaryStyle(kind, col, tags[col] || ""))}"`;
    });
    xml = xml.replace("<sheetView workbookViewId=\"0\"/>", "<sheetView workbookViewId=\"0\" showGridLines=\"0\"/>");
    // 인쇄할 때 A4 세로 한 장 너비에 맞춤
    xml = xml.replace(/(<worksheet[^>]*>)/, '$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    const ps = '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/>';
    xml = /<ignoredErrors/.test(xml) ? xml.replace("<ignoredErrors", ps + "<ignoredErrors") : xml.replace("</worksheet>", ps + "</worksheet>");
    const addTo = (tag, items) => { if (!items.length) return; css = css.replace(new RegExp(`<${tag} count="(\\d+)">([\\s\\S]*?)</${tag}>`), (m, n, body) => `<${tag} count="${+n + items.length}">${body}${items.join("")}</${tag}>`); };
    addTo("fonts", fonts); addTo("fills", fills); addTo("borders", borders); addTo("cellXfs", xfs);
    const nfs = `<numFmt numFmtId="200" formatCode="#,##0;-#,##0;&quot;–&quot;"/><numFmt numFmtId="201" formatCode="0%;-0%;&quot;–&quot;"/>`;
    css = /<numFmts count="\d+">/.test(css) ? css.replace(/<numFmts count="(\d+)">/, (m, n) => `<numFmts count="${+n + 2}">${nfs}`) : css.replace(/(<styleSheet[^>]*>)/, `$1<numFmts count="2">${nfs}</numFmts>`);
    setXml(st, css); setXml(sh, xml);
  }
  // 결과 엑셀 '요약' 시트: 점검대상목록을 COUNTIFS·SUMIFS로 세는 수식표라, 엑셀에서 검토결과·소명결과를 고치면 숫자가 바로 바뀜
  // (SheetJS 무료판은 피벗 테이블을 만들지 못함). 행 순서는 저장할 때의 지적 건수 기준으로 고정
  function summarySheet(rows) {
    const heads = Object.keys(rows[0] || {});
    const N = rows.length, last = N + 1;
    const rng = (h) => `점검대상목록!$${XLSX.utils.encode_col(heads.indexOf(h))}$2:$${XLSX.utils.encode_col(heads.indexOf(h))}$${last}`;
    const RV = rng("검토결과"), GR = rng("등급"), AM = rng("금액"), RU = rng("걸린 규칙"), DP = rng("부서"), SR = rng("소명결과");
    const q = (v) => `"${String(v).replace(/"/g, '""')}"`;
    const flagged = (r) => r.검토결과 === "소명요청" || r.검토결과 === "부적정";
    const aoa = [];
    const kinds = [];  // 행마다 [종류, 열별 강조] → 저장 후 서식으로 바꿈(styleSummary)
    const put = (row, kind = row.length ? "body" : "blank", tags = []) => { aoa.push(row); kinds.push([kind, tags]); };
    // 셀: [수식, 미리 계산한 값, 표시형식]
    const F = (f, v, z) => ({ t: "n", f, v, ...(z ? { z } : {}) });
    const cnt = (pred) => rows.filter(pred).length, sum = (pred) => rows.filter(pred).reduce((a, r) => a + (Number(r.금액) || 0), 0);
    put([`${S.month} 법인카드 점검결과 요약`], "title");
    put([`점검대상목록 ${N}건 기준. 점검대상목록 시트에서 검토결과·소명결과를 고치면 아래 숫자가 바로 바뀝니다. 지적 = 소명요청 + 부적정. 행 순서는 저장할 때 기준입니다.`], "sub");
    put([]);
    // ① 검토결과별
    put(["① 검토결과별 건수·금액"], "sec");
    put(["검토결과", "상", "중", "하", "건수", "비율", "금액(원)"], "head");
    const top = aoa.length + 1;
    for (const v of ["미검토", "소명요청", "부적정", "적정"]) {
      const key = v === "미검토" ? "" : v, m = (r) => (r.검토결과 || "") === key, rr = aoa.length + 1;
      put([v, ...["상", "중", "하"].map((g) => F(`COUNTIFS(${RV},${q(key)},${GR},${q(g)})`, cnt((r) => m(r) && r.등급 === g))),
        F(`SUM(B${rr}:D${rr})`, cnt(m)), F(`IF(${N}=0,0,E${rr}/${N})`, N ? cnt(m) / N : 0, "0%"), F(`SUMIFS(${AM},${RV},${q(key)})`, sum(m), "#,##0")], `rv-${v}`, ["", "", "", "", "", "pct"]);
    }
    const tr = aoa.length + 1, col = (c) => `SUM(${c}${top}:${c}${tr - 1})`;
    put(["합계", ...["B", "C", "D", "E"].map((c, i) => F(col(c), i < 3 ? cnt((r) => r.등급 === ["상", "중", "하"][i]) : N)), F(col("F"), 1, "0%"), F(col("G"), sum(() => true), "#,##0")], "total", ["", "", "", "", "", "pct"]);
    put([]);
    // ② 규칙별
    put(["② 규칙별 검토결과 (지적 건수 → 지적 금액 순)"], "sec");
    put(["점검 규칙", "선별", "지적", "지적 금액(원)", "적정", "미검토"], "head");
    const codes = [...new Set(rows.flatMap((r) => String(r["걸린 규칙"] || "").split(" / ").filter(Boolean)))];
    const ruleRows = codes.map((label) => {
      const code = label.split(" ")[0], pat = `*${code} *`, has = (r) => ` ${r["걸린 규칙"]} `.includes(` ${code} `) || String(r["걸린 규칙"]).startsWith(`${code} `);
      return { label, pat, has, f: cnt((r) => has(r) && flagged(r)), fa: sum((r) => has(r) && flagged(r)) };
    }).sort((a, b) => b.f - a.f || b.fa - a.fa || a.label.localeCompare(b.label));
    for (const x of ruleRows) {
      const P = `${RU},${q(x.pat)}`;
      put([x.label, F(`COUNTIFS(${P})`, cnt(x.has)), F(`COUNTIFS(${P},${RV},"소명요청")+COUNTIFS(${P},${RV},"부적정")`, x.f),
        F(`SUMIFS(${AM},${P},${RV},"소명요청")+SUMIFS(${AM},${P},${RV},"부적정")`, x.fa, "#,##0"),
        F(`COUNTIFS(${P},${RV},"적정")`, cnt((r) => x.has(r) && r.검토결과 === "적정")), F(`COUNTIFS(${P},${RV},"")`, cnt((r) => x.has(r) && !r.검토결과))], "body", ["", "", "fl", "fl", "", "mu"]);
    }
    put(["한 건이 여러 규칙에 걸리면 규칙마다 한 번씩 세므로 규칙별 합계는 전체 건수보다 많습니다."], "note");
    put([]);
    // ③ 부서별
    put(["③ 부서별 검토결과 (지적 건수 → 지적 금액 순)"], "sec");
    put(["예산부서", "선별", "지적", "지적 금액(원)", "부적정", "적정", "미검토"], "head");
    const depts = [...new Set(rows.map((r) => r.부서 || ""))].map((d) => ({ d, f: cnt((r) => r.부서 === d && flagged(r)), fa: sum((r) => r.부서 === d && flagged(r)) }))
      .sort((a, b) => b.f - a.f || b.fa - a.fa || a.d.localeCompare(b.d));
    for (const { d, f, fa } of depts) {
      const P = `${DP},${q(d)}`, m = (r) => r.부서 === d;
      put([d, F(`COUNTIFS(${P})`, cnt(m)), F(`COUNTIFS(${P},${RV},"소명요청")+COUNTIFS(${P},${RV},"부적정")`, f),
        F(`SUMIFS(${AM},${P},${RV},"소명요청")+SUMIFS(${AM},${P},${RV},"부적정")`, fa, "#,##0"),
        F(`COUNTIFS(${P},${RV},"부적정")`, cnt((r) => m(r) && r.검토결과 === "부적정")), F(`COUNTIFS(${P},${RV},"적정")`, cnt((r) => m(r) && r.검토결과 === "적정")), F(`COUNTIFS(${P},${RV},"")`, cnt((r) => m(r) && !r.검토결과))], "body", ["", "", "fl", "fl", "ng", "", "mu"]);
    }
    put([]);
    // ④ 소명 처리 현황
    put(["④ 소명 처리 현황 (검토결과가 소명요청인 건)"], "sec");
    put(["예산부서", "소명요청", ...SM_RESULTS, "결과 미입력"], "head");
    const s4 = aoa.length + 1, ask = (r) => r.검토결과 === "소명요청";
    for (const { d } of depts) {
      const P = `${DP},${q(d)},${RV},"소명요청"`, m = (r) => r.부서 === d && ask(r);
      put([d, F(`COUNTIFS(${P})`, cnt(m)), ...SM_RESULTS.map((x) => F(`COUNTIFS(${P},${SR},${q(x)})`, cnt((r) => m(r) && r.소명결과 === x))), F(`COUNTIFS(${P},${SR},"")`, cnt((r) => m(r) && !r.소명결과))], "body", ["", "", "", "ng", "wa", "mu"]);
    }
    const e4 = aoa.length;
    put(["합계", ...["B", "C", "D", "E", "F"].map((c, i) => F(`SUM(${c}${s4}:${c}${e4})`, i === 0 ? cnt(ask) : i <= 3 ? cnt((r) => ask(r) && r.소명결과 === SM_RESULTS[i - 1]) : cnt((r) => ask(r) && !r.소명결과)))], "total");
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 30 }, { wch: 11 }, { wch: 11 }, { wch: 15 }, { wch: 11 }, { wch: 11 }, { wch: 15 }];
    ws["!rows"] = kinds.map(([k]) => ({ hpt: { title: 30, sub: 18, sec: 26, head: 22, blank: 10, note: 18 }[k] || 20 }));
    SUMMARY_KINDS = kinds;
    return ws;
  }
  function saveFile(wb, name, post) {
    // 파일(file://)로 연 페이지에서는 blob 주소와 한글 파일명이 무시될 수 있어 data 주소 + 영문 파일명으로 저장
    let b64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
    if (post) {
      try {
        const zip = XLSX.CFB.read(new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" })), { type: "array" });
        post(zip);
        const u8 = new Uint8Array(XLSX.CFB.write(zip, { fileType: "zip", type: "array" }));
        let bin = ""; for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
        b64 = btoa(bin);
      } catch (e) { console.error("서식·목록 추가 실패, 기본 형태로 저장", e); }
    }
    const a = document.createElement("a"); a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + b64; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
  }
  function flash(sel, msg) { const b = $(sel); const t = b.textContent; b.textContent = msg; b.disabled = true; setTimeout(() => { b.textContent = t; b.disabled = false; }, 1600); }
  function exportRows() {
    return R.list.map((r) => ({
      승인번호: r.no, 승인일시: `${r.dkey} ${r.time}`, 직원: who(r), 부서: r.dept, 가맹점: r.merchant, 업종: r.industry, 계정: r.glName, 금액: r.amount,
      "걸린 규칙": [...r.hits].sort().map((h) => `${h} ${ENGINE.ruleName(h)}`).concat([...r.tHits].map((h) => `${h} ${ENGINE.T_NAMES[h]}`)).join(" / "),
      위험점수: r.score, 등급: r.grade, "선별 사유": r.reason, 추가확인사항: r.checks.join("; "),
      검토결과: S.reviews[r.no]?.result || "", 검토의견: S.reviews[r.no]?.opinion || "",
      ...Object.fromEntries(SM_COLS.map((c) => [c, S.reviews[r.no]?.sm?.[c] ?? ""])),
    }));
  }

  // ── 데이터 불러오기 ──
  const SHEETS = { custom: "추가조건", tx: "전체", accounts: "계정코드", industries: "업종코드", holidays: "공휴일", rules: "점검기준(가상)", employees: "직원마스터", vacations: "휴가내역", answers: "정답지" };
  function mergeRef(key, rows, file) {
    const r = ENGINE.mergeRef(key, S.refRaw[key] || [], rows);
    S.refRaw[key] = r.rows;
    (S.refFiles[key] = S.refFiles[key] || []).push({ file, n: r.added });
  }
  async function loadFiles(files) {
    const t0 = performance.now();
    if (typeof XLSX === "undefined") { $("#srcLine").innerHTML = '<span class="ng-txt">엑셀 읽기 라이브러리를 불러오지 못했습니다. 페이지를 새로 고친 뒤 다시 올려 주세요.</span>'; return; }
    for (const f of files) {
      try {
        const wb = XLSX.read(await f.arrayBuffer(), { cellDates: true, ...(/\.csv$/i.test(f.name) ? { codepage: 949 } : {}) });
        const find = (name) => wb.SheetNames.find((n) => n.replace(/\s/g, "") === name.replace(/\s/g, "")) || wb.SheetNames.find((n) => n.startsWith(name.replace(/\(.*$/, "")));
        const read = (n) => XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: "", raw: true });
        let txName = wb.SheetNames.includes("점검대상목록") ? null : find(SHEETS.tx) || wb.SheetNames.find((n) => { const h = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1 })[0] || []; return h.includes("승인번호") && h.includes("전표금액"); });
        if (txName) {
          const rows = read(txName);
          S.sheetRows += rows.length;
          let added = 0;
          for (const r of rows) { const k = String(r["승인번호"] ?? "").trim(); if (!k) continue; if (!S.raw.has(k)) added++; S.raw.set(k, r); }
          S.readCount += rows.filter((r) => String(r["승인번호"] ?? "").trim()).length;
        }
        for (const key of ["accounts", "industries", "holidays", "rules", "employees", "vacations", "custom"]) {
          const n = find(SHEETS[key]); if (n) mergeRef(key, read(n), f.name);
        }
        const an = find(SHEETS.answers); if (an) S.answers = read(an);
        if (wb.SheetNames.includes("규칙설정")) readConfigSheet(read("규칙설정"));
        // 이전에 저장한 결과 엑셀이면 검토결과·의견을 되살림
        const prev = wb.SheetNames.includes("점검대상목록") ? read("점검대상목록") : [];
        let restored = 0;
        const smVal = (v) => (v instanceof Date ? `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}` : String(v ?? "").trim());
        for (const p of prev) {
          const k = String(p["승인번호"] ?? "").trim();
          const sm = Object.fromEntries(SM_COLS.map((c) => [c, smVal(p[c])]).filter(([, v]) => v));
          if (k && (p["검토결과"] || p["검토의견"] || Object.keys(sm).length)) { S.reviews[k] = { result: p["검토결과"] || "", opinion: p["검토의견"] || "", ...(Object.keys(sm).length ? { sm } : {}) }; restored++; }
        }
        if (restored) store.set("cc_reviews", S.reviews);
        S.files.push(f.name);
      } catch (err) {
        console.error(f.name, err); S.errors = (S.errors || []).concat(f.name);
      }
    }
    S.source = "upload"; S.month = null; S.sel = null; S.weights = null;
    closeEditor();
    S.loadMs = performance.now() - t0;
    run();
    if (S.errors?.length) { $("#srcLine").insertAdjacentHTML("beforeend", ` · <span class="ng-txt">읽지 못한 파일: ${S.errors.map(esc).join(", ")}</span>`); S.errors = []; }
  }

  renderEmpty();
})();
