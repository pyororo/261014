// 사용: node offline/build_artifact.js offline/dist/법인카드점검_v2.0.html data/법인카드_가상데이터_202509-202608.xlsx <출력.html>
// v2.0 오프라인 HTML → 아티팩트용: 문서 껍데기·CSP 제거, 가상데이터 자동 불러오기, 엑셀 저장은 downloads 기능 사용
const fs = require('fs');
const [src, data, out] = process.argv.slice(2);
let h = fs.readFileSync(src, 'utf8');
const must = (cond, msg) => { if (!cond) throw new Error(msg); };
h = h.replace(/^<!doctype html>\s*<html[^>]*>\s*<head>\s*/i, '');
h = h.replace(/<meta charset="utf-8">\s*<meta name="viewport"[^>]*>\s*/, '');
h = h.replace(/<!-- 외부 통신 전면 차단[^>]*-->\s*<meta http-equiv="Content-Security-Policy"[^>]*>\s*/, '');
must(!/Content-Security-Policy/.test(h), 'csp');
h = h.replace('<title>법인카드 점검 도구</title>', '<title>법인카드 점검 PoC</title>');
h = h.replace(/<\/head>\s*<body class="is-empty">/, '<script>document.body.classList.add("is-empty")</script>');
must(!/<\/head>|<body/.test(h.slice(0, 60000)), 'body');
h = h.replace(/<\/body>\s*<\/html>\s*$/, '');
const oldSave = 'const a = document.createElement("a"); a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + b64; a.download = name;\n    document.body.appendChild(a); a.click(); a.remove();';
must(h.includes(oldSave), 'save');
h = h.replace(oldSave, `const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    (window.claude?.use ? window.claude.use("downloads") : Promise.resolve(null)).then((dl) => {
      if (!dl) { $("#srcLine").insertAdjacentHTML("beforeend", ' · <span class="ng-txt">이 화면에서는 파일 저장을 쓸 수 없습니다.</span>'); return; }
      return dl.save({ filename: name, data: bytes }).catch((e) => { if (e?.code !== "declined") console.warn("save", e); });
    });`);
const tail = '  renderEmpty();\n})();';
const i = h.lastIndexOf(tail); must(i > 0, 'tail');
const b64 = fs.readFileSync(data).toString('base64');
h = h.slice(0, i) + `  renderEmpty();
  // 아티팩트: 교육용 가상데이터를 미리 불러와 첫 화면부터 점검 결과를 보여 줌 (다른 파일을 올리면 함께 합쳐 읽음, '불러온 파일 비우기'로 초기화)
  const SAMPLE_XLSX = "${b64}";
  loadFiles([new File([Uint8Array.from(atob(SAMPLE_XLSX), (c) => c.charCodeAt(0))], "법인카드_가상데이터_202509-202608.xlsx (가상)")]);
})();` + h.slice(i + tail.length);
h = h.replace(/\uFFFD/g, '\\uFFFD'); // 아티팩트 게시가 U+FFFD 글자를 거부하므로 JS 이스케이프로 바꿈
fs.writeFileSync(out, h);
console.log('ok', out, h.length);
