// ───────────── 화면 ─────────────
(() => {
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const won = ENGINE.won;
  const fmt = (n) => Math.round(n).toLocaleString("ko-KR");
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  };

  const REF_SHEETS = { accounts: "계정코드", industries: "업종코드", holidays: "공휴일", rules: "점검기준", employees: "직원마스터", vacations: "휴가내역", custom: "추가조건" };
  const REF_LABELS = REF_SHEETS;
  const S = {
    source: "none", files: [], raw: new Map(), sheetRows: 0, readCount: 0, refRaw: {}, refFiles: {}, answers: null, loadMs: 0,
    params: { ...ENGINE.DEFAULT_PARAMS }, weights: null, month: null, tab: store.get("cc_tab", "list"), grade: "전체", review: "전체", sel: null,
    reviews: store.get("cc_reviews", {}), ms: 0,
  };
  let R = null; // 계산 결과

  // ── 계산 파이프라인 ──
  function run() {
    if (!S.raw.size || !S.refRaw.accounts) { renderEmpty(); return; }
    document.body.classList.remove("is-empty");
    const t0 = performance.now();
    const ref = ENGINE.buildRef(S.refRaw);
    const columns = [...new Set([...S.raw.values()].slice(0, 50).flatMap((r) => Object.keys(r)))];
    const custom = ENGINE.parseCustom(S.refRaw.custom || [], columns);
    for (const c of custom.rules) ref.weights[c.id] = c.weight;
    if (S.weights) Object.assign(ref.weights, S.weights);
    S.weights = { ...ref.weights };
    const prep = ENGINE.prepare([...S.raw.values()], ref);
    const rows = ENGINE.applyRules(prep.rows, ref, S.params, custom.rules);
    const months = [...new Set(rows.map((r) => r.month))].sort();
    if (!S.month || !months.includes(S.month)) S.month = months[months.length - 1];
    const period = ENGINE.periodRules(rows, ref, S.params, S.month);
    const graded = S.answers && S.answers.length ? ENGINE.grade(rows, S.answers) : null;
    const list = ENGINE.selectMonth(rows, S.month);
    // 기간 규칙에 걸린 대상월 건도 목록에 표시 (단건 점수 없음 → T 표시만)
    const recs = ENGINE.recommend(rows, ref, S.month);
    const summary = ENGINE.monthlySummary(rows);
    R = { ref, prep, rows, months, period, graded, list, recs, summary, custom, columns };
    render();
    S.ms = S.loadMs + (performance.now() - t0);
    renderScore();
  }

  // ── 렌더 ──
  function render() {
    $("#month").innerHTML = R.months.map((m) => `<option value="${m}" ${m === S.month ? "selected" : ""}>${m.replace("-", "년 ")}월</option>`).join("");
    const refTxt = Object.entries(REF_LABELS).map(([k, v]) => `${v} ${S.refRaw[k]?.length ? "✓" : "–"}`).join(" · ");
    const srcTxt = `<span class="tag">업로드</span> ${S.files.map(esc).join(", ")} · 승인번호 기준 중복 제거 후 ${fmt(S.raw.size)}건 · 기준표: ${refTxt}`;
    $("#srcLine").innerHTML = srcTxt + ` · 분석 대상 ${fmt(R.rows.length)}건 (취소 ${R.prep.stats.cancels}건 상계 제외 ${R.prep.stats.matched * 2}건${R.prep.stats.unmatched ? `, 미매칭 ${R.prep.stats.unmatched}건` : ""})`;
    renderTabs(); renderList(); renderAnswers(); renderPeriod(); renderRecs(); renderSettings(); renderBrowseFilters(); renderBrowse();
    if (!$("#ruleDrawer").hidden) renderRuleDrawer();
  }

  function renderScore() {
    const g = R.graded;
    const cnt = { 상: 0, 중: 0, 하: 0 }; R.list.forEach((r) => cnt[r.grade]++);
    const crit = [
      { lbl: "데이터 읽기", val: `${fmt(S.readCount)}<small> / ${fmt(S.sheetRows)}행</small>`, tgt: "원본 행수와 읽은 건수 일치", ok: S.readCount === S.sheetRows && S.sheetRows > 0 },
      { lbl: "업로드 → 목록 표시", val: `${(S.ms / 1000).toFixed(2)}<small> 초</small>`, tgt: "1분 이내", ok: S.ms < 60000 },
      { lbl: `${S.month.replace("-", "년 ")}월 선별`, val: `${R.list.length}<small> 건</small>`, tgt: `상 ${cnt.상} · 중 ${cnt.중} · 하 ${cnt.하}`, ok: undefined },
    ];
    if (g) crit.push(
      { lbl: "정답지 단건 탐지", val: `${g.sHit}<small> / ${g.S.length}종</small>`, tgt: "전부 탐지", ok: g.S.length > 0 && g.sHit === g.S.length },
      { lbl: "정답지 기간 패턴", val: `${g.tHit}<small> / ${g.T.length}종</small>`, tgt: "전부 탐지", ok: g.T.length > 0 && g.tHit === g.T.length },
      { lbl: "정상거래 '상' 오분류", val: `${g.fpHigh}<small> / ${g.fpN}건</small>`, tgt: "0건", ok: g.fpN > 0 && g.fpHigh === 0 });
    else crit.push(
      { lbl: "분석 대상", val: `${fmt(R.rows.length)}<small> 건</small>`, tgt: `취소 ${R.prep.stats.cancels}건 상계${R.prep.stats.unmatched ? `, 미매칭 ${R.prep.stats.unmatched}건` : ""}`, ok: undefined },
      { lbl: "기간 패턴", val: `${Object.values(R.period).reduce((s, x) => s + x.length, 0)}<small> 건</small>`, tgt: "T-R1~T-R4 누적 분석", ok: undefined },
      { lbl: "추가 점검항목 추천", val: `${R.recs.length}<small> 건</small>`, tgt: "최근 3개월 집계 기반", ok: undefined });
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
    $("#cntList").textContent = R.list.length;
    $("#cntAns").textContent = R.graded ? `${R.graded.sHit + R.graded.tHit}/${R.graded.S.length + R.graded.T.length}` : "";
    $("#cntPer").textContent = Object.values(R.period).reduce((s, a) => s + a.length, 0);
    $("#cntRec").textContent = R.recs.length;
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
      R01: "승인일이 토·일 또는 공휴일 시트에 있는 날",
      R02: `승인 시각이 ${P.nightStart}시 이후 또는 ${P.nightEnd}시 이전`,
      R03: "가맹점 업종의 제한구분이 금지 또는 사전승인 (금지 업종은 점수와 관계없이 '상')",
      R04: `같은 카드·가맹점·날짜에 ${P.splitGap}분 간격 이내로 2건 이상, 합계가 계정 건당 한도 초과`,
      R05: "전표금액이 계정코드 시트의 건당 한도 초과",
      R06: `적요가 비었거나 ${P.memoMin}자 미만, 또는 접대비인데 상대방(ㅇㅇㅇ·직함·n인) 표기 없음`,
      R07: "가맹점 업종이 계정코드 시트의 허용 가맹점업종에 없음",
      R09: "같은 카드·가맹점·금액·날짜가 2건 이상 (취소로 상계된 건 제외)",
      R11: "국내해외코드명이 해외",
      R12: `금액이 ${w(P.round)}원 단위로 떨어지고 ${w(P.round)}원 이상`,
      R13: "결재문서제목의 비용 항목(하이패스·항공·숙박 등)과 업종이 다르거나, 근무지 안 식사·회식을 여비교통비로 처리",
      R14: `같은 카드의 연속 2건이 시·도가 다르고 ${P.distGap}분 이내 (온라인·교통 업종 제외)`,
      R15: `회계일자 − 승인일이 ${P.settleDays}일 초과`,
      R16: `공용카드이고 직원명이 비었으며 금액이 ${w(P.sharedMin)}원 이상`,
      R17: "휴가내역 시트의 휴가 기간 중 본인 카드 사용",
      "T-R1": `같은 직원·가맹점이 최근 3개월 ${P.t1Count}회 이상`,
      "T-R2": `직원 월 사용액이 직전 6개월 평균의 ${P.t2Ratio}배 초과 (월 ${w(P.t2Floor)}원 이상)`,
      "T-R3": `분기말 월에 부서의 월말 5영업일 사용액이 ${Math.round(P.t3Share * 100)}% 초과, 월 사용액이 부서 평균의 ${P.t3Mult}배 이상`,
      "T-R4": `같은 직원·가맹점·금액이 ${P.t4Months}개월 이상 연속`,
    };
    if (D[id]) return D[id];
    const c = R.custom.rules.find((x) => x.id === id);
    return c ? c.conds.map((x) => `${x.col} ${x.op}${x.val !== "" ? " " + x.val : ""}`).join(" 그리고 ") : "";
  }
  function renderRuleDrawer() {
    if (!R) return;
    const cnt = {};
    R.list.forEach((r) => { r.hits.forEach((h) => (cnt[h] = (cnt[h] || 0) + 1)); r.tHits.forEach((h) => (cnt[h] = (cnt[h] || 0) + 1)); });
    const row = (id, name, weight, kind) => `<li class="rule${S.ruleFocus === id ? " focus" : ""}" id="rule-${esc(id)}">
      <div class="rh"><b class="mono">${esc(id)}</b> <span>${esc(name)}</span>${weight !== null ? `<span class="w">가중치 ${weight}</span>` : ""}</div>
      <div class="rd">${esc(ruleDesc(id))}</div>
      <button type="button" class="rf" data-filter="${esc(id)}" ${cnt[id] ? "" : "disabled"}>${S.month} 선별 중 ${cnt[id] || 0}건${cnt[id] ? (S.ruleFilter === id ? " · 보는 중" : " · 이 규칙 건만 보기") : ""}</button></li>`;
    const sec = (title, items) => items.length ? `<h3>${title}</h3><ul>${items.join("")}</ul>` : "";
    $("#ruleBody").innerHTML =
      `<p class="muted">위험점수 = 걸린 규칙 가중치 합 · 6점 이상 <span class="g g-상">상</span> 3~5점 <span class="g g-중">중</span> 1~2점 <span class="g g-하">하</span> · 금지 업종은 무조건 상 · 휴일·심야만 걸리고 적요에 사전승인 문서번호가 있으면 규칙당 1점 감점</p>` +
      sec("단건·다건 규칙", ENGINE.RULE_IDS.map((id) => row(id, ENGINE.RULE_NAMES[id], S.weights[id] ?? 0))) +
      sec(`추가조건 <small>기준표 '추가조건' 시트</small>`, R.custom.rules.map((c) => row(c.id, c.name, S.weights[c.id] ?? c.weight))) +
      sec("기간 규칙 <small>점수 없이 표시만</small>", Object.keys(ENGINE.T_NAMES).map((id) => row(id, ENGINE.T_NAMES[id], null)));
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
    const b = e.target.closest("[data-filter]"); if (!b) return;
    S.ruleFilter = S.ruleFilter === b.dataset.filter ? null : b.dataset.filter; S.grade = "전체"; S.tab = "list";
    renderTabs(); renderList(); renderRuleDrawer();
  });

  function renderList() {
    const counts = { 전체: R.list.length, 상: 0, 중: 0, 하: 0 };
    R.list.forEach((r) => counts[r.grade]++);
    $("#gradeFilter").innerHTML = `<span class="flabel">등급</span>` + Object.entries(counts).map(([k, v]) => `<button type="button" data-g="${k}" aria-pressed="${S.grade === k}">${k} <span class="mono">${v}</span></button>`).join("");
    const reviewOf = (r) => S.reviews[r.no]?.result || "미검토";
    const rcount = { 전체: R.list.length, 미검토: 0, 적정: 0, 소명요청: 0, 부적정: 0 };
    R.list.forEach((r) => rcount[reviewOf(r)]++);
    if (!(S.review in rcount)) S.review = "전체";
    $("#reviewFilter").innerHTML = `<span class="flabel">검토</span>` + Object.entries(rcount).map(([k, v]) => `<button type="button" data-r="${k}" aria-pressed="${S.review === k}">${k} <span class="mono">${v}</span></button>`).join("");
    if (S.ruleFilter && !R.list.some((r) => r.hits.has(S.ruleFilter) || r.tHits.has(S.ruleFilter))) S.ruleFilter = null;
    if (S.ruleFilter) $("#gradeFilter").insertAdjacentHTML("beforeend", `<button type="button" class="rulef" data-clear-rule="1" aria-pressed="true">${esc(S.ruleFilter)} ${esc(ENGINE.ruleName(S.ruleFilter) === S.ruleFilter ? ENGINE.T_NAMES[S.ruleFilter] || "" : ENGINE.ruleName(S.ruleFilter))} 만 보는 중 ✕</button>`);
    const rows = R.list.filter((r) => (S.grade === "전체" || r.grade === S.grade) && (S.review === "전체" || reviewOf(r) === S.review) && (!S.ruleFilter || r.hits.has(S.ruleFilter) || r.tHits.has(S.ruleFilter)));
    const done = R.list.filter((r) => S.reviews[r.no]?.result).length;
    $("#reviewProgress").textContent = `검토 입력 ${done}/${R.list.length}건`;
    $("#listTbl").innerHTML = `<thead><tr><th>등급</th><th>점수</th><th>걸린 규칙</th><th>승인일시</th><th>직원 · 부서</th><th>가맹점 · 적요</th><th style="text-align:right">금액</th><th>선별 사유</th><th>검토</th></tr></thead><tbody>` +
      (rows.length ? rows.map((r) => { const m = spotMarker(r); return `<tr class="click sev-${r.grade} rvrow-${S.reviews[r.no]?.result || "미검토"} ${S.sel === r.no ? "on" : ""}" data-no="${r.no}" tabindex="0">
        <td><span class="g g-${r.grade}">${r.grade}</span></td><td class="n">${r.score}${r.adjust ? `<div class="muted" title="사전승인 기재로 조정">${r.ruleScore}→${r.score}</div>` : ""}</td>
        <td><div class="chips">${ruleChips(r)}</div></td>
        <td class="mono" style="white-space:nowrap">${r.dkey.slice(5)} ${m("time", r.time)}<div class="muted">${m("day", (r.holName || r.wdName) + (r.hits.has("R17") ? " · 휴가 중" : ""))}</div>${r.hits.has("R15") ? `<div class="muted">${m("settle", `정산 ${Math.round((r.acctDate - r.date) / 86400000)}일`)}</div>` : ""}</td>
        <td style="white-space:nowrap">${m("who", esc(r.emp) || '<span class="ng-txt">미기재</span>')}<div class="muted">${esc(r.dept)}</div></td>
        <td class="mer">${m("mer", esc(r.merchant))}${r.overseas ? ` <span class="muted">(${esc(r.currency || "해외")})</span>` : ""}<div class="muted">${m("ind", esc(r.industry))} · ${m("gl", esc(r.glName))}</div>
          <div class="memo-line"><span class="lbl">적요</span><span>${m("memo", r.memo ? esc(r.memo) : '<span class="ng-txt">적요 공란</span>')}</span></div>${m.has("title") ? `<div class="memo-line"><span class="lbl">제목</span><span>${m("title", esc(r.title))}</span></div>` : ""}</td>
        <td class="n">${m("amt", fmt(r.amount))}${r.limit && (r.hits.has("R05") || r.hits.has("R04")) ? `<div class="muted">한도 ${fmt(r.limit)}</div>` : ""}</td>
        <td class="reason">${esc(ENGINE.reasonShort(r))}</td><td style="white-space:nowrap">${reviewPill(r)}</td></tr>`; }).join("")
        : `<tr><td colspan="9" class="empty">선택한 조건(등급·검토결과·규칙)에 해당하는 선별 건이 없습니다.</td></tr>`) + "</tbody>";
    renderDetail();
  }

  function renderDetail() {
    const r = R.list.find((x) => x.no === S.sel);
    if (!r) { $("#detail").innerHTML = `<p class="empty">행을 누르면 건별 상세, AI 추가 확인사항, 검토결과 입력란이 열립니다.</p>`; return; }
    const rv = S.reviews[r.no] || {};
    $("#detail").innerHTML = `<div class="detail">
      <div class="panel"><h2><span class="g g-${r.grade}">${r.grade}</span> 승인번호 <span class="mono">${r.no}</span> <small>위험점수 ${r.score}${r.adjust ? ` (규칙 ${r.ruleScore}점, AI 조정 ${r.adjust})` : ""}</small></h2>
        <dl class="kv">
          <dt>승인일시</dt><dd class="mono">${r.dkey} ${r.time} (${r.holName || r.wdName + "요일"})</dd>
          <dt>카드 · 구분</dt><dd class="mono">${esc(r.card)} · ${esc(r.kind)}</dd>
          <dt>직원 · 부서</dt><dd>${esc(r.emp) || "미기재"} · ${esc(r.dept)}${r.work ? ` (근무지 ${esc(r.work)})` : ""}</dd>
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
    $("#ansSummary").innerHTML = `단건 <b>${g.sHit}/${g.S.length}</b> · 기간 패턴 <b>${g.tHit}/${g.T.length}</b> · 정상거래(FP) ${g.fpN}건 중 '상' <b>${g.fpHigh}건</b>. 정답지는 채점에만 쓰고 선별 로직에는 넣지 않습니다.`;
    const fpGroup = g.list.filter((o) => o.type === "FP");
    const body = g.list.filter((o) => o.type !== "FP").map((o) => `<tr>
      <td class="mono">${o.id}</td><td>${esc(o.title)}</td><td class="mono">${esc(o.expect)}</td>
      <td>${o.items.map((i) => i.row ? `<div class="chips" style="margin-bottom:2px">${o.type === "S" ? `<span class="g g-${i.row.grade}" style="width:20px;height:18px;font-size:11px">${i.row.grade || "–"}</span>` : ""}${o.type === "S" ? ruleChips({ hits: i.row.hits, tHits: new Set() }) : [...i.row.tHits].map((h) => `<span class="chip t">${h}</span>`).join("")}</div>` : `<div class="ng-txt">승인번호 ${i.no} 없음</div>`).slice(0, o.type === "T" ? 1 : 9).join("")}${o.type === "T" ? `<div class="muted">${o.items.length}건 묶음</div>` : ""}</td>
      <td>${o.detected ? `<span class="ok-txt">탐지</span>` : `<span class="ng-txt">미탐</span>`}<div class="muted">${esc(o.note)}</div></td></tr>`).join("");
    const fpRow = fpGroup.length ? `<tr><td class="mono">FP01–${fpGroup[fpGroup.length - 1].id.slice(2)}</td><td>${esc(fpGroup[0].title)}</td><td class="mono">${esc(fpGroup[0].expect)}</td>
      <td>${Object.entries(g.fpGrades).map(([k, v]) => `${k} ${v}건`).join(" · ")}</td><td>${g.fpHigh === 0 ? `<span class="ok-txt">상 0건</span>` : `<span class="ng-txt">상 ${g.fpHigh}건</span>`}<div class="muted">사전승인 문서번호로 저위험 조정</div></td></tr>` : "";
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
    for (let v = 0; v <= top; v += step) svg += `<line x1="${pl}" x2="${W - pr}" y1="${y(v)}" y2="${y(v)}" stroke="var(--rule-2)"/><text x="${pl - 6}" y="${y(v) + 4}" text-anchor="end">${(v / 1e6).toFixed(0)}백만</text>`;
    sm.forEach((m, i) => {
      const x = pl + i * bw + bw * 0.18, w = bw * 0.64;
      svg += `<rect x="${x}" y="${y(m.amount)}" width="${w}" height="${y(0) - y(m.amount)}" rx="2" fill="${m.month === S.month ? "var(--accent)" : "var(--rule)"}"><title>${m.month}: ${won(m.amount)} (${m.n}건)</title></rect>`;
      svg += `<text x="${pl + i * bw + bw / 2}" y="${H - 10}" text-anchor="middle">${m.month.slice(2).replace("-", ".")}</text>`;
    });
    const pts = sm.map((m, i) => `${pl + i * bw + bw / 2},${y(m.ma3)}`).join(" ");
    svg += `<polyline points="${pts}" fill="none" stroke="var(--mid)" stroke-width="2"/>`;
    const last = sm[sm.length - 1];
    if (last) svg += `<circle cx="${pl + (sm.length - 1) * bw + bw / 2}" cy="${y(last.ma3)}" r="3.5" fill="var(--mid)"/>`;
    $("#chart").innerHTML = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="월별 사용액 막대 차트">${svg}</svg>`;

    $("#monthTbl").innerHTML = `<thead><tr><th>월</th><th style="text-align:right">건수</th><th style="text-align:right">금액</th><th style="text-align:right">건당평균</th><th style="text-align:right">전월 대비</th><th style="text-align:right">선별</th></tr></thead><tbody>` +
      sm.map((m, i) => { const p = sm[i - 1]; const d = p ? ((m.amount - p.amount) / p.amount) * 100 : null;
        return `<tr${m.month === S.month ? ' class="on"' : ""}><td class="mono">${m.month}</td><td class="n">${m.n}</td><td class="n">${fmt(m.amount)}</td><td class="n">${fmt(m.avg)}</td><td class="n">${d === null ? "–" : (d > 0 ? "+" : "") + d.toFixed(1) + "%"}</td><td class="n">${m.flagged}</td></tr>`; }).join("") + "</tbody>";

    const P = S.params;
    const defs = {
      "T-R1": `직원×가맹점 최근 3개월(${ENGINE.addMonths(S.month, -2)}~${S.month}) ${P.t1Count}회 이상`,
      "T-R2": `직원 월 사용액 ÷ 직전 6개월 평균 > ${P.t2Ratio} (월 ${fmt(P.t2Floor)}원 이상, 최근 3개월)`,
      "T-R3": `분기말 월 부서 월말 5영업일 비중 > ${P.t3Share * 100}% 그리고 월 사용액 ≥ 부서 월평균 × ${P.t3Mult}`,
      "T-R4": `동일 직원·가맹점·금액 ${P.t4Months}개월 이상 연속`,
    };
    const item = {
      "T-R1": (x) => `<tr><td>${esc(x.who)}<div class="muted">${esc(x.dept)}</div></td><td>${esc(x.merchant)}</td><td class="n">${x.n}회</td><td class="n">${fmt(x.amount)}</td><td>${esc(x.focus) || "–"}</td></tr>`,
      "T-R2": (x) => `<tr><td>${esc(x.who)}<div class="muted">${esc(x.dept)}</div></td><td class="mono">${x.month}</td><td class="n">${fmt(x.cur)}</td><td class="n">${fmt(x.avg)}</td><td class="n"><b>×${x.ratio.toFixed(1)}</b></td><td>${spark(x.series)}</td></tr>`,
      "T-R3": (x) => `<tr><td>${esc(x.who)}</td><td class="mono">${x.month}</td><td class="n">${(x.share * 100).toFixed(0)}%</td><td class="n">${fmt(x.total)}</td><td class="n">×${(x.total / x.avg).toFixed(1)}</td></tr>`,
      "T-R4": (x) => `<tr><td>${esc(x.who)}<div class="muted">${esc(x.dept)}</div></td><td>${esc(x.merchant)}</td><td class="n">${fmt(x.amount)}</td><td class="n">${x.months}개월</td></tr>`,
    };
    const heads = { "T-R1": ["직원", "가맹점", "3개월 건수", "금액", "요일 집중"], "T-R2": ["직원", "기준월", "월 사용액", "직전 6개월 평균", "배율", "12개월 추이"], "T-R3": ["부서", "월", "월말 5영업일 비중", "월 사용액", "부서 평균 대비"], "T-R4": ["직원", "가맹점", "금액", "연속"] };
    $("#patterns").innerHTML = Object.keys(ENGINE.T_NAMES).map((k) => `<div class="panel pattern"><div class="head"><b>${k}</b><h2 style="margin:0">${ENGINE.T_NAMES[k]}</h2><span class="tag">${R.period[k].length}건</span></div>
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

  function renderSettings() {
    const labels = { nightStart: "R02 심야 시작(시, 이상)", nightEnd: "R02 심야 종료(시, 미만)", splitGap: "R04 분할결제 간격(분)", memoMin: "R06 적요 최소 글자수", settleDays: "R15 정산 기한(일)", round: "R12 라운드 단위(원)", sharedMin: "R16 공용카드 기준금액(원)", distGap: "R14 원거리 간격(분)", t1Count: "T-R1 3개월 반복 횟수", t2Ratio: "T-R2 급증 배율", t2Floor: "T-R2 최소 월 사용액(원)", t3Share: "T-R3 월말 비중(0~1)", t3Mult: "T-R3 부서 평균 배수", t4Months: "T-R4 연속 개월" };
    $("#params").innerHTML = Object.keys(labels).map((k) => `<label for="p_${k}">${labels[k]}<input id="p_${k}" data-p="${k}" type="number" step="any" value="${S.params[k]}"></label>`).join("");
    $("#weights").innerHTML = ENGINE.RULE_IDS.concat(R.custom.rules.map((c) => c.id)).map((k) => `<label for="w_${k}">${k} ${esc(ENGINE.ruleName(k))}<input id="w_${k}" data-w="${k}" type="number" min="0" max="10" value="${S.weights[k]}"></label>`).join("");
  }

  function renderCustom() {
    const { rules, errors } = R.custom;
    const hitN = (id) => R.rows.filter((r) => r.hits.has(id)).length;
    const monthN = (id) => R.list.filter((r) => r.hits.has(id)).length;
    $("#customList").innerHTML = (errors.length ? `<div class="notice" style="border-left-color:var(--high)">${errors.map(esc).join("<br>")}</div>` : "") +
      (rules.length ? `<div class="tbl-wrap"><table><thead><tr><th>ID</th><th>조건명</th><th>조건 (모두 만족)</th><th style="text-align:right">가중치</th><th style="text-align:right">전체 적중</th><th style="text-align:right">${S.month} 선별</th></tr></thead><tbody>` +
        rules.map((c) => `<tr><td class="mono">${c.id}</td><td>${esc(c.name)}</td><td>${c.conds.map((x) => `<span class="chip">${esc(x.col)} ${esc(x.op)}${x.val !== "" ? " " + esc(x.val) : ""}</span>`).join(" ")}</td><td class="n">${S.weights[c.id]}</td><td class="n">${hitN(c.id)}</td><td class="n">${monthN(c.id)}</td></tr>`).join("") + "</tbody></table></div>"
        : `<p class="empty">기준표에 '추가조건' 시트가 없거나 비어 있습니다. 기준표 양식을 내려받으면 예시가 들어 있습니다.</p>`);
  }

  // ── 거래 조회 ──
  const B = { period: "month", q: "", gl: "", ind: "", dept: "", min: "", max: "", sel: "", hol: false, cancel: false };
  function renderBrowseFilters() {
    const all = R.prep.rows;
    const opts = (key) => ["", ...[...new Set(all.map((r) => r[key]).filter(Boolean))].sort()].map((v) => `<option value="${esc(v)}" ${B[{ glName: "gl", industry: "ind", dept: "dept" }[key]] === v ? "selected" : ""}>${v ? esc(v) : "전체"}</option>`).join("");
    $("#b_gl").innerHTML = opts("glName"); $("#b_ind").innerHTML = opts("industry"); $("#b_dept").innerHTML = opts("dept");
    renderCustom();
  }
  function browseRows() {
    const q = B.q.trim();
    return R.prep.rows.filter((r) => {
      if (B.period === "month" && r.month !== S.month) return false;
      if (!B.cancel && r.excluded) return false;
      if (B.gl && r.glName !== B.gl) return false;
      if (B.ind && r.industry !== B.ind) return false;
      if (B.dept && r.dept !== B.dept) return false;
      if (B.min !== "" && r.amount < Number(B.min)) return false;
      if (B.max !== "" && r.amount > Number(B.max)) return false;
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
        <td style="white-space:nowrap">${esc(r.emp) || '<span class="muted">미기재</span>'}<div class="muted">${esc(r.dept)}</div></td><td>${esc(r.merchant)}<div class="muted">${esc(r.industry)}</div></td>
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
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data.length ? data : [{ 안내: "조건에 맞는 거래가 없습니다" }]), "거래조회");
    const cond = [["조회 조건"], ["기간", B.period === "month" ? S.month : "전체 기간"], ["검색어", B.q], ["계정", B.gl || "전체"], ["업종", B.ind || "전체"], ["부서", B.dept || "전체"],
      ["금액", `${B.min || "-"} ~ ${B.max || "-"}`], ["선별", { "": "전체", y: "선별된 건만", n: "선별 안 된 건만" }[B.sel]], ["휴일만", B.hol ? "예" : "아니오"], ["취소·상계 포함", B.cancel ? "예" : "아니오"], ["건수", rows.length]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cond), "조회조건");
    saveFile(wb, `card-audit_browse_${B.period === "month" ? S.month : "all"}.xlsx`);
    flash("#b_xlsx", "저장했습니다");
  }
  $("#browseFilters").addEventListener("input", (e) => {
    const k = e.target.dataset.b; if (!k) return;
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
  $("#params").addEventListener("change", (e) => { const k = e.target.dataset.p; if (!k) return; const v = Number(e.target.value); if (!isFinite(v)) return; S.params[k] = v; run(); });
  $("#weights").addEventListener("change", (e) => { const k = e.target.dataset.w; if (!k) return; S.weights[k] = Math.max(0, Number(e.target.value) || 0); run(); });
  $("#resetParams").addEventListener("click", () => { S.params = { ...ENGINE.DEFAULT_PARAMS }; S.weights = null; run(); });
  $("#uploadBtn").addEventListener("click", () => $("#file").click());
  $("#resetBtn").addEventListener("click", () => {
    if ($("#resetBtn").dataset.confirm !== "1") { $("#resetBtn").dataset.confirm = "1"; $("#resetBtn").textContent = "한 번 더 누르면 초기화"; setTimeout(() => { $("#resetBtn").dataset.confirm = ""; $("#resetBtn").textContent = "불러온 파일 비우기"; }, 3000); return; }
    $("#resetBtn").dataset.confirm = ""; $("#resetBtn").textContent = "불러온 파일 비우기";
    Object.assign(S, { source: "none", files: [], raw: new Map(), sheetRows: 0, readCount: 0, refRaw: {}, answers: null, month: null, sel: null, weights: null });
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
    for (const [k, name] of Object.entries(REF_SHEETS)) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(DEFAULT_REF[k]), name);
    saveFile(wb, "card-audit_reference-template.xlsx");
  });
  $("#file").addEventListener("change", (e) => { if (e.target.files.length) loadFiles([...e.target.files]); e.target.value = ""; });
  $("#copyBtn").addEventListener("click", async () => {
    const tsv = exportRows().map((r) => Object.values(r).map((v) => String(v ?? "").replace(/[\t\n]/g, " ")).join("\t"));
    const text = [Object.keys(exportRows()[0] || {}).join("\t"), ...tsv].join("\n");
    try { await navigator.clipboard.writeText(text); flash("#copyBtn", "복사했습니다"); } catch { flash("#copyBtn", "복사할 수 없습니다"); }
  });
  $("#xlsxBtn").addEventListener("click", () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(exportRows()), "점검대상목록");
    const per = [];
    for (const [k, list] of Object.entries(R.period)) for (const x of list) per.push({ 패턴: `${k} ${ENGINE.T_NAMES[k]}`, 대상: x.who, 부서: x.dept || "", 가맹점: x.merchant || "", 기준월: x.month || "", 지표: k === "T-R1" ? `3개월 ${x.n}회, ${fmt(x.amount)}원 ${x.focus || ""}` : k === "T-R2" ? `${fmt(x.cur)}원 / 직전 6개월 평균 ${fmt(x.avg)}원 (×${x.ratio.toFixed(1)})` : k === "T-R3" ? `월말 5영업일 ${(x.share * 100).toFixed(0)}%, 부서 평균의 ${(x.total / x.avg).toFixed(1)}배` : `${fmt(x.amount)}원 ${x.months}개월 연속`, 관련건수: x.rows.length });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(per.length ? per : [{ 패턴: "해당 없음" }]), "기간패턴");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(R.recs.map((r) => ({ 추천항목: r.title, 근거: r.basis, 이유: r.why, 규칙안: r.rule }))), "추가점검항목");
    saveFile(wb, `card-audit_result_${S.month}.xlsx`);
    flash("#xlsxBtn", "저장했습니다");
  });
  function saveFile(wb, name) {
    // 파일(file://)로 연 페이지에서는 blob 주소와 한글 파일명이 무시될 수 있어 data 주소 + 영문 파일명으로 저장
    const b64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
    const a = document.createElement("a"); a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + b64; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
  }
  function flash(sel, msg) { const b = $(sel); const t = b.textContent; b.textContent = msg; b.disabled = true; setTimeout(() => { b.textContent = t; b.disabled = false; }, 1600); }
  function exportRows() {
    return R.list.map((r) => ({
      승인번호: r.no, 승인일시: `${r.dkey} ${r.time}`, 직원: r.emp || "미기재", 부서: r.dept, 가맹점: r.merchant, 업종: r.industry, 계정: r.glName, 금액: r.amount,
      "걸린 규칙": [...r.hits].sort().map((h) => `${h} ${ENGINE.ruleName(h)}`).concat([...r.tHits].map((h) => `${h} ${ENGINE.T_NAMES[h]}`)).join(" / "),
      위험점수: r.score, 등급: r.grade, "선별 사유": r.reason, 추가확인사항: r.checks.join("; "),
      검토결과: S.reviews[r.no]?.result || "", 검토의견: S.reviews[r.no]?.opinion || "",
    }));
  }

  // ── 데이터 불러오기 ──
  const SHEETS = { custom: "추가조건", tx: "전체", accounts: "계정코드", industries: "업종코드", holidays: "공휴일", rules: "점검기준(가상)", employees: "직원마스터", vacations: "휴가내역", answers: "정답지" };
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
          const n = find(SHEETS[key]); if (n) { S.refRaw[key] = read(n); }
        }
        const an = find(SHEETS.answers); if (an) S.answers = read(an);
        // 이전에 저장한 결과 엑셀이면 검토결과·의견을 되살림
        const prev = wb.SheetNames.includes("점검대상목록") ? read("점검대상목록") : [];
        let restored = 0;
        for (const p of prev) { const k = String(p["승인번호"] ?? "").trim(); if (k && (p["검토결과"] || p["검토의견"])) { S.reviews[k] = { result: p["검토결과"] || "", opinion: p["검토의견"] || "" }; restored++; } }
        if (restored) store.set("cc_reviews", S.reviews);
        S.files.push(f.name);
      } catch (err) {
        console.error(f.name, err); S.errors = (S.errors || []).concat(f.name);
      }
    }
    S.source = "upload"; S.month = null; S.sel = null; S.weights = null;
    S.loadMs = performance.now() - t0;
    run();
    if (S.errors?.length) { $("#srcLine").insertAdjacentHTML("beforeend", ` · <span class="ng-txt">읽지 못한 파일: ${S.errors.map(esc).join(", ")}</span>`); S.errors = []; }
  }

  renderEmpty();
})();
