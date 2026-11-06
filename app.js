/* 등록서류 작성 도우미 - 메인 로직 */
(function () {
  "use strict";
  pdfjsLib.GlobalWorkerOptions.workerSrc = "vendor/pdf.worker.min.js";

  const TEXT_BUDGET = 320000;      // AI에 보낼 본문 최대 글자 수
  const SCANNED_CHARS = 60;        // 쪽당 평균 글자 수가 이보다 적으면 스캔본으로 판단
  const PARALLEL = 2;              // 동시에 작성할 항목 수
  const TYPES = [                  // 화면 열 순서
    { key: "제출", label: "국문<br>요약서", doc: "시험요약서" },
    { key: "면제", label: "면제<br>사유서", doc: "면제사유서" },
    { key: "생략", label: "생략<br>사유서", doc: "생략사유서" },
  ];
  const DOCNAME = { 제출: "국문요약서", 면제: "면제사유서", 생략: "생략사유서" };
  const FILESHORT = { phys: "물화", human: "인체", env: "환경" };
  const $ = (id) => document.getElementById(id);
  const store = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} },
  };

  let CONFIG = null;
  let files = [];               // {id, name, bytes, numPages, pages[], scanned}
  let fileSeq = 0;
  const selection = new Map();  // 항목번호 → "제출" | "면제" | "생략"
  let tasks = [];               // 작성 결과

  // ---------------- 초기화 ----------------
  async function init() {
    const k = store.get("regdoc.apiKey");
    if (k) { $("apiKey").value = k; $("rememberKey").checked = true; }
    const m = store.get("regdoc.model"); if (m) $("model").value = m;
    const t = store.get("regdoc.tonnage"); if (t) $("tonnage").value = t;
    $("apiKey").addEventListener("change", saveKey);
    $("rememberKey").addEventListener("change", saveKey);
    $("model").addEventListener("change", () => store.set("regdoc.model", $("model").value.trim()));
    $("tonnage").addEventListener("change", () => { store.set("regdoc.tonnage", $("tonnage").value); renderPanels(); });

    CONFIG = await (await fetch("config/templates.json", { cache: "no-cache" })).json();
    renderPanels();

    const drop = $("drop");
    drop.addEventListener("click", () => $("fileInput").click());
    drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") $("fileInput").click(); });
    drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
    drop.addEventListener("dragleave", () => drop.classList.remove("over"));
    drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("over"); addFiles(e.dataTransfer.files); });
    $("fileInput").addEventListener("change", (e) => { addFiles(e.target.files); e.target.value = ""; });
    $("runBtn").addEventListener("click", run);
    $("dlAll").addEventListener("click", downloadAll);
  }

  function saveKey() {
    if ($("rememberKey").checked) store.set("regdoc.apiKey", $("apiKey").value.trim());
    else store.del("regdoc.apiKey");
  }

  // ---------------- 시험항목 선택 패널 ----------------
  function renderPanels() {
    const tier = +$("tonnage").value;
    // 톤수 범위를 벗어난 선택은 해제
    for (const no of [...selection.keys()]) {
      const c = CONFIG.catalog.find((x) => x.no === no);
      if (!c || c.tier > tier) selection.delete(no);
    }
    $("panels").innerHTML = CONFIG.groups.map((g) => {
      const rows = CONFIG.catalog.filter((c) => c.file === g.file && c.tier <= tier).map((c) => {
        const has = Object.keys(c.types).length > 0;
        const cells = TYPES.map((t) => c.types[t.key]
          ? `<td class="opt"><input type="radio" name="sel-${c.no}" value="${t.key}" data-no="${c.no}" title="${esc(c.name)} · ${DOCNAME[t.key]}"${selection.get(c.no) === t.key ? " checked" : ""}></td>`
          : `<td class="opt"><span class="dash">–</span></td>`).join("");
        return `<tr class="${has ? "" : "none"}"><td class="iname"><span class="no">${c.no}</span>${esc(c.name)}</td>${cells}</tr>`;
      }).join("");
      return `<div class="panel"><table>
        <thead><tr><th class="pname">${esc(g.label)}</th>${TYPES.map((t) => `<th>${t.label}</th>`).join("")}</tr></thead>
        <tbody>${rows}</tbody></table>
        <div class="panel-actions"><button type="button" data-clear="${g.file}">선택 해제</button></div></div>`;
    }).join("");

    $("panels").querySelectorAll("input[type=radio]").forEach((r) => {
      r.addEventListener("mousedown", () => { r.dataset.was = r.checked ? "1" : ""; });
      r.addEventListener("keydown", () => { r.dataset.was = ""; });
      r.addEventListener("click", () => {
        const no = +r.dataset.no;
        if (r.dataset.was === "1") { r.checked = false; selection.delete(no); }
        else selection.set(no, r.value);
        r.dataset.was = "";
        updateSummary();
      });
    });
    $("panels").querySelectorAll("[data-clear]").forEach((b) => b.addEventListener("click", () => {
      for (const c of CONFIG.catalog) if (c.file === b.dataset.clear) selection.delete(c.no);
      renderPanels();
    }));
    updateSummary();
  }

  function updateSummary() {
    const n = selection.size;
    const cnt = { 제출: 0, 면제: 0, 생략: 0 };
    for (const v of selection.values()) cnt[v]++;
    $("selSummary").textContent = n
      ? `선택 ${n}개 — 국문요약서 ${cnt.제출} · 면제사유서 ${cnt.면제} · 생략사유서 ${cnt.생략}`
      : "아직 선택한 항목이 없습니다.";
    updateRunBtn();
  }
  function updateRunBtn() {
    $("runBtn").disabled = !(selection.size && files.some((f) => !f.loading && !f.error));
    $("runBtn").textContent = selection.size ? `등록서류 작성 (${selection.size}개 항목)` : "등록서류 작성";
  }

  // ---------------- 파일 ----------------
  async function addFiles(fileList) {
    const list = Array.from(fileList);   // 입력창을 비우면 FileList가 비워지므로 먼저 복사
    for (const f of list) {
      const ext = (f.name.match(/\.([a-z0-9]+)$/i) || [])[1]?.toLowerCase();
      if (["doc", "xls", "ppt", "hwp"].includes(ext)) {
        alert(`${f.name}\n예전 형식(.${ext})은 읽을 수 없습니다. ${ext === "hwp" ? "PDF로 저장" : `.${ext}x 형식(또는 PDF)으로 다시 저장`}해서 올려 주세요.`);
        continue;
      }
      if (!["pdf", "docx", "xlsx", "pptx"].includes(ext)) { alert(`PDF·워드·엑셀·파워포인트 파일만 올릴 수 있습니다: ${f.name}`); continue; }
      const entry = { id: ++fileSeq, name: f.name, kind: ext, bytes: await f.arrayBuffer(), numPages: 0, pages: [], names: [], scanned: false, loading: true };
      files.push(entry);
      renderFiles();
      if (ext !== "pdf") {
        try {
          const r = await OfficeText.extract(entry.bytes.slice(0), ext);
          entry.pages = r.pages; entry.names = r.names; entry.approx = r.approx; entry.numPages = r.pages.length;
          if (!entry.pages.join("").trim()) entry.error = "글자를 찾지 못했습니다(그림만 있는 파일). PDF로 저장해 올려 주세요.";
        } catch (e) {
          entry.error = "파일을 읽지 못했습니다(암호 또는 손상): " + e.message;
        }
        entry.loading = false;
        renderFiles();
        continue;
      }
      try {
        const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(entry.bytes.slice(0)) }).promise;
        entry.numPages = pdf.numPages;
        for (let p = 1; p <= pdf.numPages; p++) {
          const page = await pdf.getPage(p);
          const tc = await page.getTextContent();
          let s = "", lastY = null;
          for (const it of tc.items) {
            const y = it.transform ? it.transform[5] : null;
            if (lastY !== null && y !== null && Math.abs(y - lastY) > 2) s += "\n"; else if (s && !s.endsWith(" ")) s += " ";
            s += it.str; lastY = y;
          }
          entry.pages.push(s.replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim());
        }
        const total = entry.pages.reduce((a, b) => a + b.length, 0);
        entry.scanned = total / Math.max(1, entry.numPages) < SCANNED_CHARS;
      } catch (e) {
        entry.error = "PDF를 읽지 못했습니다(암호 또는 손상): " + e.message;
      }
      entry.loading = false;
      renderFiles();
    }
  }
  function renderFiles() {
    $("fileList").innerHTML = files.map((f) => `<li><span>📄 ${esc(f.name)}</span><span class="meta">${
      f.loading ? "읽는 중…" : f.error ? `<span style="color:var(--err)">${esc(f.error)}</span>` :
      unitText(f)}</span><button title="삭제" data-id="${f.id}">✕</button></li>`).join("");
    $("fileList").querySelectorAll("button").forEach((b) => b.addEventListener("click", () => {
      files = files.filter((f) => f.id !== +b.dataset.id); renderFiles();
    }));
    updateRunBtn();
  }

  const KINDNAME = { pdf: "PDF", docx: "워드", xlsx: "엑셀", pptx: "파워포인트" };
  const UNIT = { pdf: "쪽", docx: "쪽", xlsx: "시트", pptx: "슬라이드" };
  function unitText(f) {
    if (f.kind === "pdf") return `PDF · ${f.numPages}쪽${f.scanned ? " · 스캔본(이미지)" : ""}`;
    if (f.kind === "docx") return f.approx === "chunk" ? `워드 · ${f.numPages}구간(쪽 정보 없음)` : `워드 · 약 ${f.numPages}쪽`;
    return `${KINDNAME[f.kind]} · ${f.numPages}${UNIT[f.kind]}`;
  }
  function unitHeader(f) {
    if (f.kind === "pdf") return `PDF, 총 ${f.numPages}쪽 — [p.N]은 PDF의 N번째 쪽`;
    if (f.kind === "docx") return f.approx === "chunk" ? `워드 문서, 쪽 정보가 없어 ${f.numPages}개 구간으로 나눔 — [p.N]은 N번째 구간` : `워드 문서, 약 ${f.numPages}쪽 — [p.N]은 N번째 쪽(근사치)`;
    if (f.kind === "xlsx") return `엑셀 파일, 시트 ${f.numPages}개 — [p.N]은 N번째 시트(셀주소: 값 형식)`;
    return `파워포인트, 슬라이드 ${f.numPages}장 — [p.N]은 N번째 슬라이드`;
  }

  // ---------------- 진행 로그 ----------------
  function log(msg, cls) {
    $("sec-progress").classList.remove("hidden");
    const li = document.createElement("li"); li.textContent = msg; if (cls) li.className = cls;
    $("log").appendChild(li); li.scrollIntoView({ block: "nearest" });
  }

  // ---------------- 실행 ----------------
  async function run() {
    const key = $("apiKey").value.trim();
    if (!key) { alert("1번 설정에 Anthropic API 키를 입력하세요."); $("apiKey").focus(); return; }
    const ok = files.filter((f) => !f.loading && !f.error);
    tasks = [...selection.entries()].sort((a, b) => a[0] - b[0]).map(([no, type]) => {
      const cat = CONFIG.catalog.find((c) => c.no === no);
      const item = CONFIG.items.find((i) => i.key === cat.types[type]);
      return { id: `t${no}`, no, type, cat, item, sheet: CONFIG.sheets[item.key], status: "wait", values: {}, evidence: {}, notes: "" };
    });
    $("log").innerHTML = "";
    $("runBtn").disabled = true;
    renderResults(ok);
    $("sec-result").classList.remove("hidden");
    try {
      const keywords = [...new Set(tasks.flatMap((t) => t.item.keywords))];
      const { content, sentInfo } = buildDocuments(ok, keywords);
      sentInfo.forEach((s) => log(s));
      // 같은 자료를 여러 항목에 반복 전송하므로 캐시 표시(비용·시간 절감)
      if (content.length) content[content.length - 1].cache_control = { type: "ephemeral" };
      log(`${tasks.length}개 항목 작성을 시작합니다. 항목당 보통 1~3분 걸립니다.`);
      const queue = [...tasks];
      const worker = async () => {
        while (queue.length) {
          const t = queue.shift();
          await runTask(t, key, content, ok);
        }
      };
      await Promise.all(Array.from({ length: Math.min(PARALLEL, tasks.length) }, worker));
      const done = tasks.filter((t) => t.status === "ok").length;
      log(`완료: ${done}/${tasks.length}개 항목. 아래에서 확인 후 다운로드하세요.`, done === tasks.length ? "ok" : "err");
    } catch (e) {
      console.error(e);
      log("오류: " + e.message, "err");
    } finally {
      updateRunBtn();
    }
  }

  async function runTask(t, key, docContent, ok) {
    t.status = "run"; updateTaskHead(t);
    const t0 = Date.now();
    try {
      const prompt = buildPrompt(t, ok);
      const text = await callClaude(key, $("model").value.trim(), SYSTEM_PROMPT, [...docContent, { type: "text", text: prompt }]);
      const json = parseJson(text);
      for (const f of t.sheet.fields) {
        let v = json.fields && Object.prototype.hasOwnProperty.call(json.fields, f.cell) ? json.fields[f.cell] : undefined;
        if (v === undefined || v === null) v = f.default ?? "";
        t.values[f.cell] = String(v);
        if (f.label === "화학물질명" && $("chemName").value.trim()) t.values[f.cell] = $("chemName").value.trim();
        if (f.label === "CAS No." && $("cas").value.trim()) t.values[f.cell] = $("cas").value.trim();
      }
      for (const f of ok) t.evidence[f.id] = new Set();
      for (const e of json.evidence || []) {
        const file = ok[(e.file || 1) - 1];
        if (!file) continue;
        for (const p of e.pages || []) if (p >= 1 && p <= file.numPages) t.evidence[file.id].add(+p);
      }
      t.notes = json.notes || "";
      t.status = "ok";
      log(`✔ ${t.no}. ${t.cat.name} (${DOCNAME[t.type]}) — ${Math.round((Date.now() - t0) / 1000)}초`, "ok");
    } catch (e) {
      console.error(e);
      t.status = "err"; t.error = e.message;
      log(`✖ ${t.no}. ${t.cat.name}: ${e.message}`, "err");
    }
    renderTaskBody(t, ok);
  }

  function buildDocuments(list, keywordList) {
    const kw = [...keywordList, $("chemName").value.trim(), $("cas").value.trim()].filter(Boolean).map((k) => k.toLowerCase());
    const content = [], sentInfo = [];
    const textFiles = list.filter((f) => !f.scanned);
    const totalChars = textFiles.reduce((a, f) => a + f.pages.reduce((x, y) => x + y.length, 0), 0);
    const perFileBudget = (f) => Math.floor(TEXT_BUDGET * (f.pages.reduce((x, y) => x + y.length, 0) / Math.max(1, totalChars)));
    list.forEach((f, i) => {
      const no = i + 1;
      if (f.scanned) {
        if (f.bytes.byteLength > 30 * 1024 * 1024 || f.numPages > 100) {
          sentInfo.push(`⚠ 문서 ${no} "${f.name}"은 스캔본이며 너무 커서(30MB/100쪽 초과) 제외했습니다. 필요한 부분만 잘라 올려 주세요.`);
          return;
        }
        content.push({ type: "text", text: `<문서 ${no}: ${f.name} (총 ${f.numPages}쪽, 스캔본 PDF 원본 첨부 — 쪽 번호는 PDF의 물리적 쪽 순서)>` });
        content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: b64(f.bytes) } });
        sentInfo.push(`문서 ${no} "${f.name}": 스캔본이라 PDF 이미지 그대로 전송 (${f.numPages}쪽)`);
        return;
      }
      const pages = f.pages.map((t, idx) => ({ p: idx + 1, t }));
      const total = pages.reduce((a, b) => a + b.t.length, 0);
      let selected = pages;
      if (totalChars > TEXT_BUDGET) {
        const budget = perFileBudget(f);
        const scored = pages.map((x) => {
          const low = x.t.toLowerCase();
          let s = 0; for (const k of kw) { if (k.length < 2) continue; let pos = 0; while ((pos = low.indexOf(k, pos)) !== -1) { s++; pos += k.length; } }
          return { ...x, s: x.p <= 2 ? 1e9 : s };
        }).sort((a, b) => b.s - a.s);
        selected = []; let used = 0;
        for (const x of scored) { if (x.s <= 0 && x.p > 2) break; if (used + x.t.length > budget) continue; selected.push(x); used += x.t.length; }
        selected.sort((a, b) => a.p - b.p);
        sentInfo.push(`문서 ${no} "${f.name}": ${f.numPages}${UNIT[f.kind]} 중 관련도 높은 ${selected.length}${UNIT[f.kind]}만 전송 (분량 제한)`);
      } else {
        sentInfo.push(`문서 ${no} "${f.name}": ${KINDNAME[f.kind]} ${f.numPages}${UNIT[f.kind]} 전체 전송 (${total.toLocaleString()}자)`);
      }
      const body = selected.map((x) => `[p.${x.p}]${f.names[x.p - 1] && f.kind !== "pdf" ? ` (${f.names[x.p - 1]})` : ""}\n${x.t}`).join("\n\n");
      content.push({ type: "text", text: `<문서 ${no}: ${f.name} (${unitHeader(f)})>\n${body}\n</문서 ${no}>` });
    });
    return { content, sentInfo };
  }

  function b64(buf) {
    const bytes = new Uint8Array(buf); let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  const SYSTEM_PROMPT = `너는 한국 「화학물질의 등록 및 평가 등에 관한 법률」(화평법) 등록신청자료 작성 전문가다.
업로드된 자료(국가·국제기구 평가보고서, 논문, 시험보고서)를 근거로 지정된 엑셀 양식의 입력칸을 채운다.

[작성 원칙]
1. 반드시 제공된 문서에 있는 내용만 쓴다. 문서에서 확인되지 않는 값은 빈 문자열("")로 둔다. 추측하거나 지어내지 않는다.
2. '선택지'가 있는 칸은 선택지 중 하나를 글자 그대로 고른다. 맞는 것이 없고 선택지에 'other'가 있으면 "other: 원문 내용" 형식으로 쓴다.
3. '기본값'은 양식에 이미 들어 있는 값이다. 고시명·'면제'·'자료 면제 (Data waiving)' 같은 고정 문구는 그대로 둔다. 'AAA'는 화학물질명으로 바꾼다. '해당없음' 기본값은 해당 조건이 없으면 그대로 둔다.
4. '작성예시'는 다른 물질의 작성 사례다. 문체·형식·서술 순서만 참고하고 그 내용(물질명, 수치, 출처)을 절대 옮기지 않는다.
5. 결론은 한두 문장으로 "~함." 또는 "~이다." 체로, 요약/본문은 시험방법·조건·결과·판정을 서술형 국문 문단으로 쓴다. 영문 원문은 정확히 번역하고, 시험지침명(OECD TG 등)·학명·단위는 원문 표기를 유지한다.
6. 수치는 원문의 값·단위·부등호를 그대로 쓴다(예: > 2000 mg/kg bw).
7. 출처 표기는 "저자 또는 기관(연도)., 제목, 쪽p." 형식으로 쓴다.
8. 면제사유서는 면제조건 해당 여부를, 생략사유서는 생략조건(시행령 제13조)과 근거자료를 문서에서 확인해 작성한다.
9. 근거로 사용한 위치를 evidence에 모두 적는다. 번호는 문서 본문의 [p.N] 표기의 N을 쓴다(PDF는 쪽, 워드는 쪽 또는 구간, 엑셀은 시트, 파워포인트는 슬라이드 순번). PDF 표지는 자동으로 포함되므로 적지 않아도 된다.

[출력 형식] 아래 JSON 하나만 출력한다. 다른 설명이나 코드블록 표시는 쓰지 않는다.
{"fields": {"셀주소": "값", ...}, "evidence": [{"file": 문서번호, "pages": [쪽번호, ...]}], "notes": "검토자가 확인해야 할 사항(자료에서 못 찾은 항목, 판단이 애매한 부분 등)을 한국어로 간단히"}`;

  function buildPrompt(t, list) {
    const name = $("chemName").value.trim(), cas = $("cas").value.trim();
    const ton = $("tonnage").selectedOptions[0].textContent;
    const lines = t.sheet.fields.map((f) => {
      const ctx = [f.section, f.sub].filter(Boolean).join(" > ");
      let s = `- ${f.cell} | ${ctx ? ctx + " | " : ""}${f.label}${f.tall ? " (서술형)" : ""}`;
      if (f.options && f.options.length) {
        const opts = f.options.slice(0, 80).map((o) => o.length > 90 ? o.slice(0, 90) + "…" : o);
        s += ` | 선택지: ${opts.join(" / ")}${f.options.length > 80 ? ` (외 ${f.options.length - 80}개)` : ""}`;
      }
      if (f.default) s += ` | 기본값: "${clip(f.default, 400)}"`;
      if (f.example) s += ` | 작성예시: "${clip(f.example, 300)}"`;
      return s;
    }).join("\n");
    return `[작성 대상]
- 화학물질명: ${name || "(미입력 — 문서에서 확인)"}
- CAS No.: ${cas || "(미입력 — 문서에서 확인)"}
- 톤수범주: ${ton}
- 시험항목: ${t.no}. ${t.cat.name} (${t.cat.group})
- 작성 서류: ${DOCNAME[t.type]}
- 양식 시트: ${t.item.sheet} (제목: ${t.sheet.title})
- 업로드 문서: ${list.map((f, i) => `문서 ${i + 1} "${f.name}"`).join(", ")}

[입력칸 목록] (셀주소 | 위치 | 항목 | 선택지 | 기본값 | 작성예시)
${lines}

위 입력칸을 모두 검토해 JSON으로 답하라. fields에는 위 목록의 셀주소를 키로 쓴다. 이 시험항목과 관련 없는 문서 내용은 쓰지 않는다.`;
  }
  const clip = (s, n) => { s = String(s).replace(/\n/g, "⏎"); return s.length > n ? s.slice(0, n) + "…" : s; };

  async function callClaude(key, model, system, content) {
    let res;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        res = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers: {
            "content-type": "application/json",
            "x-api-key": key,
            "anthropic-version": "2023-06-01",
            "anthropic-dangerous-direct-browser-access": "true",
          },
          body: JSON.stringify({ model: model || "claude-sonnet-5-5", max_tokens: 16000, system, messages: [{ role: "user", content }] }),
        });
      } catch (e) {
        throw new Error("Anthropic API에 연결하지 못했습니다. 인터넷/사내 방화벽을 확인하세요. (" + e.message + ")");
      }
      if ((res.status === 429 || res.status === 529 || res.status >= 500) && attempt < 2) {
        await new Promise((r) => setTimeout(r, 8000 * (attempt + 1)));
        continue;
      }
      break;
    }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      const m = (data.error && data.error.message) || res.statusText;
      if (res.status === 401) throw new Error("API 키가 올바르지 않습니다. (" + m + ")");
      if (res.status === 404) throw new Error("모델 이름을 확인하세요: " + model + " (" + m + ")");
      if (res.status === 429 || res.status === 529) throw new Error("요청이 많아 잠시 후 다시 시도해 주세요. (" + m + ")");
      throw new Error(`API 오류 ${res.status}: ${m}`);
    }
    if (data.stop_reason === "max_tokens") log("⚠ 응답이 길어 일부가 잘렸을 수 있습니다.", "err");
    return (data.content || []).filter((c) => c.type === "text").map((c) => c.text).join("");
  }

  function parseJson(text) {
    const a = text.indexOf("{"), b = text.lastIndexOf("}");
    if (a < 0 || b < 0) throw new Error("AI 응답에서 결과(JSON)를 찾지 못했습니다.");
    try { return JSON.parse(text.slice(a, b + 1)); }
    catch (e) { throw new Error("AI 응답 형식이 올바르지 않습니다. 다시 시도해 주세요."); }
  }

  // ---------------- 결과 화면 ----------------
  const STATUS = { wait: ["대기", "wait"], run: ["작성 중…", "wait"], ok: ["완료", "ok"], err: ["오류", "err"] };

  function renderResults(ok) {
    $("results").innerHTML = tasks.map((t) => `<details class="res" id="${t.id}">
      <summary><strong>${t.no}. ${esc(t.cat.name)}</strong><span>${DOCNAME[t.type]}</span>
        <span class="hint" style="margin:0">${FILESHORT[t.cat.file]} › ${esc(t.item.sheet)}</span><span class="st"></span></summary>
      <div class="body"></div></details>`).join("");
    tasks.forEach(updateTaskHead);
  }
  function updateTaskHead(t) {
    const el = document.querySelector(`#${t.id} .st`);
    if (!el) return;
    const [txt, cls] = STATUS[t.status];
    el.textContent = txt; el.className = "st " + cls;
  }

  function renderTaskBody(t, ok) {
    updateTaskHead(t);
    const body = document.querySelector(`#${t.id} .body`);
    if (t.status === "err") { body.innerHTML = `<p style="color:var(--err)">${esc(t.error)}</p>`; return; }
    let rows = "", lastSec = null;
    t.sheet.fields.forEach((f, i) => {
      const sec = f.section || "";
      if (sec && sec !== lastSec) { rows += `<tr class="sec"><td colspan="3">${esc(sec)}</td></tr>`; lastSec = sec; }
      const v = t.values[f.cell] ?? "";
      const lid = `${t.id}-dl${i}`;
      const dl = f.options && f.options.length ? `<datalist id="${lid}">${f.options.map((o) => `<option value="${esc(o)}">`).join("")}</datalist>` : "";
      const input = f.options && f.options.length && !f.tall && v.length < 120 && !v.includes("\n")
        ? `<input type="text" list="${lid}" data-cell="${f.cell}" value="${esc(v)}">${dl}`
        : `<textarea data-cell="${f.cell}" rows="${Math.min(10, Math.max(1, Math.ceil(v.length / 60) + (v.match(/\n/g) || []).length))}">${esc(v)}</textarea>`;
      rows += `<tr><td class="cell">${f.cell}</td><td class="lbl">${esc(f.label)}${f.sub ? `<small>${esc(f.sub)}</small>` : ""}</td><td>${input}</td></tr>`;
    });
    const ev = ok.map((f) => {
      const pages = [...(t.evidence[f.id] || [])].sort((a, b) => a - b);
      const unit = f.kind === "docx" && f.approx === "chunk" ? "구간" : UNIT[f.kind];
      const note = f.kind === "pdf" ? "" : " — 원본 파일이 첨부되고 인용 위치는 발췌안내.txt에 적힘";
      return `<label style="margin-top:6px">${esc(f.name)} (${KINDNAME[f.kind]}, 총 ${f.numPages}${unit}${note})<input type="text" data-file="${f.id}" value="${pages.join(", ")}" placeholder="인용하지 않은 문서는 비워 두면 제외"></label>`;
    }).join("");
    body.innerHTML = `
      ${t.notes ? `<div class="notes">AI 검토 메모\n${esc(t.notes)}</div>` : ""}
      <div class="item-dl"><button data-act="xlsx">이 항목 엑셀</button><button data-act="pdf">이 항목 증빙자료</button></div>
      <div class="evidence"><strong>증빙자료 발췌 위치</strong> (PDF는 표지 1쪽 자동 포함, 수정 가능 · 예: 1, 12-14)${ev}</div>
      <div class="table-wrap"><table><thead><tr><th>셀</th><th>항목</th><th>값</th></tr></thead><tbody>${rows}</tbody></table></div>`;
    body.querySelectorAll("[data-cell]").forEach((el) => { el.addEventListener("input", () => checkOption(t, el)); checkOption(t, el); });
    body.querySelector('[data-act="xlsx"]').addEventListener("click", () => downloadTaskXlsx(t));
    body.querySelector('[data-act="pdf"]').addEventListener("click", () => downloadTaskPdf(t));
  }
  function checkOption(t, el) {
    const f = t.sheet.fields.find((x) => x.cell === el.dataset.cell);
    const v = el.value.trim();
    const bad = f.options && f.options.length && v && !f.options.includes(v) && !/^other/i.test(v) && !/^기타/.test(v);
    el.classList.toggle("warn", !!bad);
  }

  function collectValues(t) {
    const vals = {};
    document.querySelectorAll(`#${t.id} [data-cell]`).forEach((el) => {
      const f = t.sheet.fields.find((x) => x.cell === el.dataset.cell);
      if (el.value !== (f.default ?? "")) vals[f.cell] = el.value;
    });
    return vals;
  }

  // ---------------- 다운로드 ----------------
  const safe = (s) => String(s || "").replace(/[\\/:*?"<>|\s]+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
  const substId = () => safe($("cas").value.trim() || $("chemName").value.trim() || "물질");
  const taskName = (t) => `${t.no}.${safe(t.cat.name)}_${DOCNAME[t.type]}`;
  const tplCache = {};
  async function template(file) {
    if (!tplCache[file]) tplCache[file] = await (await fetch(CONFIG.files[file], { cache: "no-cache" })).arrayBuffer();
    return tplCache[file].slice(0);
  }

  async function workbookFor(file, list) {
    const map = {};
    for (const t of list) map[t.item.sheet] = collectValues(t);
    return XlsxFill.fillWorkbook(await template(file), map);
  }

  async function downloadTaskXlsx(t) {
    try { save(await workbookFor(t.cat.file, [t]), `${substId()}_${taskName(t)}.xlsx`); }
    catch (e) { alert("엑셀 생성 실패: " + e.message); console.error(e); }
  }

  function parsePages(s, max) {
    const out = new Set();
    for (const part of String(s).split(/[,\s]+/)) {
      const m = /^(\d+)(?:-(\d+))?$/.exec(part.trim());
      if (!m) continue;
      const a = +m[1], b = m[2] ? +m[2] : a;
      for (let p = Math.min(a, b); p <= Math.max(a, b); p++) if (p >= 1 && p <= max) out.add(p);
    }
    return [...out].sort((a, b) => a - b);
  }

  // 증빙자료: PDF는 표지+인용 쪽을 발췌해 하나로 합치고, 워드·엑셀·파워포인트는 원본 + 인용 위치 안내
  async function evidence(t) {
    const out = await PDFLib.PDFDocument.create();
    let added = 0;
    const originals = [], guide = [];
    for (const el of document.querySelectorAll(`#${t.id} [data-file]`)) {
      const f = files.find((x) => x.id === +el.dataset.file);
      if (!f) continue;
      const pages = parsePages(el.value, f.numPages);
      if (!pages.length) continue;
      if (f.kind === "pdf") {
        if (!pages.includes(1)) pages.unshift(1);
        const src = await PDFLib.PDFDocument.load(f.bytes.slice(0), { ignoreEncryption: true });
        const copied = await out.copyPages(src, pages.map((p) => p - 1));
        copied.forEach((pg) => out.addPage(pg)); added += copied.length;
        guide.push(`- ${f.name}: ${pages.join(", ")}쪽 (증빙 PDF에 발췌됨)`);
      } else {
        originals.push(f);
        guide.push(`- ${f.name}: ${pages.map((p) => f.names[p - 1] || p).join(", ")} (원본 파일 첨부 — 해당 부분을 PDF로 저장해 제출하세요)`);
      }
    }
    const pdf = added ? new Blob([await out.save()], { type: "application/pdf" }) : null;
    const text = `[${t.no}. ${t.cat.name} · ${DOCNAME[t.type]}] 증빙자료 인용 위치\r\n${guide.join("\r\n")}\r\n`;
    return { pdf, originals, text, empty: !pdf && !originals.length };
  }

  async function downloadTaskPdf(t) {
    try {
      const ev = await evidence(t);
      if (ev.empty) { alert("발췌할 위치가 없습니다. 쪽 번호를 입력하세요."); return; }
      if (!ev.originals.length) { save(ev.pdf, `증빙_${substId()}_${taskName(t)}.pdf`); return; }
      const zip = new JSZip();
      addEvidence(zip, "", t, ev);
      save(await zip.generateAsync({ type: "blob" }), `증빙_${substId()}_${taskName(t)}.zip`);
    } catch (e) { alert("증빙자료 생성 실패: " + e.message); console.error(e); }
  }

  function addEvidence(zip, prefix, t, ev) {
    if (!ev.originals.length) { if (ev.pdf) zip.file(`${prefix}${taskName(t)}.pdf`, ev.pdf); return; }
    const dir = `${prefix}${taskName(t)}/`;
    if (ev.pdf) zip.file(`${dir}증빙_발췌.pdf`, ev.pdf);
    for (const f of ev.originals) zip.file(`${dir}원본_${f.name}`, f.bytes.slice(0));
    zip.file(`${dir}발췌안내.txt`, "\ufeff" + ev.text);
  }

  async function downloadAll() {
    const done = tasks.filter((t) => t.status === "ok");
    if (!done.length) { alert("완료된 항목이 없습니다."); return; }
    try {
      const zip = new JSZip();
      for (const g of CONFIG.groups) {
        const list = done.filter((t) => t.cat.file === g.file);
        if (list.length) zip.file(`${substId()}_등록서류_${FILESHORT[g.file]}.xlsx`, await workbookFor(g.file, list));
      }
      for (const t of done) {
        const ev = await evidence(t);
        if (!ev.empty) addEvidence(zip, "증빙자료/", t, ev);
      }
      save(await zip.generateAsync({ type: "blob" }), `${substId()}_등록서류.zip`);
    } catch (e) { alert("다운로드 파일 생성 실패: " + e.message); console.error(e); }
  }

  function save(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  function esc(s) { return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

  init().catch((e) => { log("초기화 실패: " + e.message, "err"); });
})();
