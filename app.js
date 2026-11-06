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
  const conds = new Map();      // 면제 항목번호 → 선택한 면제조건 문구 ("" = AI가 판단)

  // ---------------- 초기화 ----------------
  async function init() {
    const k = store.get("regdoc.apiKey");
    if (k) { $("apiKey").value = k; $("rememberKey").checked = true; }
    const gk = store.get("regdoc.geminiKey");
    if (gk) { $("geminiKey").value = gk; $("rememberKey").checked = true; }
    const m = store.get("regdoc.model"); if (m) $("model").value = m;
    const gm = store.get("regdoc.geminiModel"); if (gm) $("geminiModel").value = gm;
    const mode = store.get("regdoc.mode") || "anthropic";
    (document.querySelector(`input[name=mode][value="${mode}"]`) || document.querySelector("input[name=mode]")).checked = true;
    document.querySelectorAll("input[name=mode]").forEach((r) => r.addEventListener("change", () => { store.set("regdoc.mode", getMode()); showMode(); }));
    showMode();
    $("geminiKey").addEventListener("change", saveKey);
    $("geminiModel").addEventListener("change", () => store.set("regdoc.geminiModel", $("geminiModel").value.trim()));
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
    if ($("rememberKey").checked) {
      store.set("regdoc.apiKey", $("apiKey").value.trim());
      store.set("regdoc.geminiKey", $("geminiKey").value.trim());
    } else { store.del("regdoc.apiKey"); store.del("regdoc.geminiKey"); }
  }
  const getMode = () => (document.querySelector("input[name=mode]:checked") || {}).value || "anthropic";
  function showMode() {
    const m = getMode();
    document.querySelectorAll(".mode-pane").forEach((p) => p.classList.toggle("hidden", p.dataset.pane !== m));
    document.querySelectorAll(".key-only").forEach((p) => p.classList.toggle("hidden", m === "manual"));
  }

  // ---------------- 시험항목 선택 패널 ----------------
  const catOf = (no) => CONFIG.catalog.find((x) => x.no === no);
  function exemptOptions(no) {
    const key = (catOf(no) || {}).types?.["면제"];
    if (!key) return [];
    const f = CONFIG.sheets[key].fields.find((x) => x.label === "면제조건");
    return (f && f.options) || [];
  }
  function linkFor(no) {
    if (selection.get(no) !== "면제" || !conds.get(no)) return null;
    return (CONFIG.links || []).find((l) => l.target === no && l.option === conds.get(no)) || null;
  }
  function linkedSources() {           // 근거 항목번호 → 이를 근거로 쓰는 면제 항목들
    const m = new Map();
    for (const no of selection.keys()) {
      const l = linkFor(no);
      if (l) for (const s of l.sources) { if (!m.has(s)) m.set(s, []); m.get(s).push(no); }
    }
    return m;
  }
  function unmetLinks() {
    const out = [];
    for (const no of selection.keys()) {
      const l = linkFor(no);
      if (!l) continue;
      const usable = l.sources.filter((s) => Object.keys(catOf(s).types).some((t) => t !== "면제"));
      const chosen = l.sources.filter((s) => selection.has(s) && selection.get(s) !== "면제");
      const ok = l.mode === "any" ? chosen.length > 0 : usable.every((s) => chosen.includes(s));
      if (!ok) out.push({ no, need: l.mode === "any" ? usable : usable.filter((s) => !chosen.includes(s)), mode: l.mode });
    }
    return out;
  }
  const short = (s, n) => (s.length > n ? s.slice(0, n) + "…" : s);

  function renderPanels() {
    const tier = +$("tonnage").value;
    // 톤수 범위를 벗어난 선택은 해제(면제 근거로 연계된 항목은 유지)
    for (let pass = 0; pass < 3; pass++) {
      const linked = linkedSources();
      for (const no of [...selection.keys()]) {
        const c = catOf(no);
        if (!c || (c.tier > tier && !linked.has(no))) { selection.delete(no); conds.delete(no); }
      }
    }
    for (const no of [...conds.keys()]) if (selection.get(no) !== "면제") conds.delete(no);
    const linked = linkedSources();
    $("panels").innerHTML = CONFIG.groups.map((g) => {
      const rows = CONFIG.catalog.filter((c) => c.file === g.file && (c.tier <= tier || linked.has(c.no))).map((c) => {
        const has = Object.keys(c.types).length > 0;
        const cells = TYPES.map((t) => c.types[t.key]
          ? `<td class="opt"><input type="radio" name="sel-${c.no}" value="${t.key}" data-no="${c.no}" title="${esc(c.name)} · ${DOCNAME[t.key]}"${selection.get(c.no) === t.key ? " checked" : ""}></td>`
          : `<td class="opt"><span class="dash">–</span></td>`).join("");
        const badge = linked.has(c.no)
          ? `<span class="link-badge" title="면제 근거로 연계된 항목">근거 → ${linked.get(c.no).join(", ")}번</span>` : "";
        let html = `<tr class="${has ? "" : "none"}${linked.has(c.no) ? " linked" : ""}"><td class="iname"><span class="no">${c.no}</span>${esc(c.name)}${badge}</td>${cells}</tr>`;
        if (selection.get(c.no) === "면제") {
          const opts = exemptOptions(c.no);
          const cur = conds.get(c.no) || "";
          const optHtml = [`<option value="">면제조건: AI가 자료를 보고 판단</option>`].concat(opts.map((o) => {
            const l = (CONFIG.links || []).find((x) => x.target === c.no && x.option === o);
            const tag = l ? ` [근거: ${l.sources.join(l.mode === "any" ? " 또는 " : "+")}번]` : "";
            return `<option value="${esc(o)}" title="${esc(o)}"${o === cur ? " selected" : ""}>${esc(short(o, 46))}${tag}</option>`;
          })).join("");
          const l = linkFor(c.no);
          const hint = l ? `<div class="cond-hint">근거 항목 ${l.sources.map((s) => `${s}.${esc(catOf(s).name)}`).join(l.mode === "any" ? " 또는 " : ", ")}의 서류 종류를 선택하세요. 근거 서류를 먼저 작성한 뒤 그 결과로 면제사유서를 씁니다.</div>` : "";
          html += `<tr class="cond"><td colspan="4">${opts.length ? `<select data-cond="${c.no}">${optHtml}</select>` : `<span class="hint">면제조건 목록 없음 — AI가 자료를 보고 작성</span>`}${hint}</td></tr>`;
        }
        return html;
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
        renderPanels();
      });
    });
    $("panels").querySelectorAll("select[data-cond]").forEach((sel) => sel.addEventListener("change", () => {
      const no = +sel.dataset.cond;
      conds.set(no, sel.value);
      const l = linkFor(no);
      if (l) {   // 근거 항목 기본 선택: 국문요약서(없으면 생략사유서)
        const pick = (s) => { const t = catOf(s).types; if (!selection.has(s) || selection.get(s) === "면제") selection.set(s, t["제출"] ? "제출" : "생략"); };
        const usable = l.sources.filter((s) => catOf(s).types["제출"] || catOf(s).types["생략"]);
        if (l.mode === "any") { if (!usable.some((s) => selection.has(s) && selection.get(s) !== "면제") && usable[0]) pick(usable[0]); }
        else usable.forEach(pick);
      }
      renderPanels();
    }));
    $("panels").querySelectorAll("[data-clear]").forEach((b) => b.addEventListener("click", () => {
      for (const c of CONFIG.catalog) if (c.file === b.dataset.clear) { selection.delete(c.no); conds.delete(c.no); }
      renderPanels();
    }));
    updateSummary();
  }

  function updateSummary() {
    const n = selection.size;
    const cnt = { 제출: 0, 면제: 0, 생략: 0 };
    for (const v of selection.values()) cnt[v]++;
    const warn = unmetLinks().map((u) => `⚠ ${u.no}번 면제 근거로 ${u.need.join(u.mode === "any" ? " 또는 " : ", ")}번의 국문요약서/생략사유서를 선택하세요.`);
    $("selSummary").innerHTML = (n
      ? `선택 ${n}개 — 국문요약서 ${cnt.제출} · 면제사유서 ${cnt.면제} · 생략사유서 ${cnt.생략}`
      : "아직 선택한 항목이 없습니다.") + (warn.length ? `<br><span style="color:var(--err)">${warn.join("<br>")}</span>` : "");
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
    const mode = getMode();
    const key = mode === "gemini" ? $("geminiKey").value.trim() : $("apiKey").value.trim();
    if (mode === "anthropic" && !key) { alert("1번 설정에 Anthropic API 키를 입력하세요."); $("apiKey").focus(); return; }
    if (mode === "gemini" && !key) { alert("1번 설정에 Gemini API 키를 입력하세요."); $("geminiKey").focus(); return; }
    const ok = files.filter((f) => !f.loading && !f.error);
    const unmet = unmetLinks();
    if (unmet.length && !confirm(unmet.map((u) => `${u.no}번 면제의 근거 항목(${u.need.join(", ")}번)이 선택되지 않았습니다.`).join("\n") + "\n\n근거 없이 계속할까요?")) return;
    tasks = [...selection.entries()].sort((a, b) => a[0] - b[0]).map(([no, type]) => {
      const cat = CONFIG.catalog.find((c) => c.no === no);
      const item = CONFIG.items.find((i) => i.key === cat.types[type]);
      return { id: `t${no}`, no, type, cat, item, sheet: CONFIG.sheets[item.key], status: "wait", values: {}, evidence: {}, notes: "",
               cond: type === "면제" ? (conds.get(no) || "") : "", link: linkFor(no) };
    });
    for (const t of tasks) t.sources = t.link ? tasks.filter((x) => t.link.sources.includes(x.no) && x.type !== "면제") : [];
    const level = (t, seen = new Set()) => (seen.has(t.no) ? 0 : (seen.add(t.no), t.sources.length ? 1 + Math.max(...t.sources.map((x) => level(x, seen))) : 0));
    tasks.forEach((t) => { t.level = level(t); });
    $("log").innerHTML = "";
    const oldBox = $("manualBox"); if (oldBox) oldBox.remove();
    $("runBtn").disabled = true;
    renderResults(ok);
    $("sec-result").classList.remove("hidden");
    try {
      const keywords = [...new Set(tasks.flatMap((t) => t.item.keywords))];
      const { content, sentInfo, manualAttach } = buildDocuments(ok, keywords, mode);
      sentInfo.forEach((s) => log(s));
      if (mode === "manual") { renderManual(content, ok, manualAttach); return; }
      // 같은 자료를 여러 항목에 반복 전송하므로 캐시 표시(비용·시간 절감)
      if (content.length) content[content.length - 1].cache_control = { type: "ephemeral" };
      log(`${tasks.length}개 항목 작성을 시작합니다. 항목당 보통 1~3분 걸립니다.` + (tasks.some((t) => t.sources.length) ? " 면제 근거 항목을 먼저 작성합니다." : ""));
      const maxLevel = Math.max(...tasks.map((t) => t.level));
      for (let lv = 0; lv <= maxLevel; lv++) {
        const queue = tasks.filter((t) => t.level === lv);
        const worker = async () => {
          while (queue.length) await runTask(queue.shift(), mode, key, content, ok);
        };
        await Promise.all(Array.from({ length: Math.min(mode === "gemini" ? 1 : PARALLEL, queue.length) }, worker));
      }
      const done = tasks.filter((t) => t.status === "ok").length;
      log(`완료: ${done}/${tasks.length}개 항목. 아래에서 확인 후 다운로드하세요.`, done === tasks.length ? "ok" : "err");
    } catch (e) {
      console.error(e);
      log("오류: " + e.message, "err");
    } finally {
      updateRunBtn();
    }
  }

  async function runTask(t, mode, key, docContent, ok) {
    t.status = "run"; updateTaskHead(t);
    const t0 = Date.now();
    try {
      const prompt = buildPrompt(t, ok);
      const msg = [...docContent, { type: "text", text: prompt }];
      const text = mode === "gemini"
        ? await callGemini(key, $("geminiModel").value.trim(), SYSTEM_PROMPT, msg)
        : await callClaude(key, $("model").value.trim(), SYSTEM_PROMPT, msg);
      applyJson(t, parseJson(text), ok);
      log(`✔ ${t.no}. ${t.cat.name} (${DOCNAME[t.type]}) — ${Math.round((Date.now() - t0) / 1000)}초`, "ok");
    } catch (e) {
      console.error(e);
      t.status = "err"; t.error = e.message;
      log(`✖ ${t.no}. ${t.cat.name}: ${e.message}`, "err");
    }
    renderTaskBody(t, ok);
  }

  function applyJson(t, json, ok) {
      t.values = {}; t.evidence = {};
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
      if (t.type === "면제") {
        const fc = t.sheet.fields.find((f) => f.label === "면제조건");
        if (fc && t.cond) t.values[fc.cell] = t.cond;
        const fa = t.sheet.fields.find((f) => f.label === "첨부 1");
        const srcDone = (t.sources || []).filter((x) => x.status === "ok");
        if (fa && srcDone.length) t.values[fa.cell] = srcDone.map(srcLabel).join(", ");
        for (const x of srcDone) for (const [fid, set] of Object.entries(x.evidence || {})) {
          if (!t.evidence[fid]) t.evidence[fid] = new Set();
          set.forEach((p) => t.evidence[fid].add(p));
        }
      }
  }
  const srcLabel = (x) => `${x.sheet.title || x.cat.name} ${DOCNAME[x.type]}`;

  // 면제사유서에 넘길 근거 서류 요약(채워진 칸만)
  function sourceSummary(x) {
    const lines = [];
    for (const f of x.sheet.fields) {
      const v = (x.values[f.cell] || "").trim();
      if (!v || v === String(f.default || "").trim()) continue;
      lines.push(`  - ${f.label}: ${v.replace(/\n/g, " ")}`);
    }
    let s = lines.join("\n");
    if (s.length > 3500) s = s.slice(0, 3500) + " …";
    return `■ ${x.no}. ${srcLabel(x)}\n${s}`;
  }

  function buildDocuments(list, keywordList, mode) {
    const manualAttach = [];
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
        if (mode === "manual") {
          manualAttach.push(f.name);
          content.push({ type: "text", text: `<문서 ${no}: ${f.name} (총 ${f.numPages}쪽, 스캔본 PDF — 이 대화에 파일로 첨부됨. 쪽 번호는 PDF의 물리적 쪽 순서)>` });
          sentInfo.push(`문서 ${no} "${f.name}": 스캔본이라 글자를 읽을 수 없습니다. claude.ai 대화창에 이 PDF를 직접 첨부하세요.`);
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
    return { content, sentInfo, manualAttach };
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

  function buildPrompt(t, list, manual) {
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

${exemptBlock(t, manual)}
위 입력칸을 모두 검토해 JSON으로 답하라. fields에는 위 목록의 셀주소를 키로 쓴다. 이 시험항목과 관련 없는 문서 내용은 쓰지 않는다.`;
  }
  function exemptBlock(t, manual) {
    if (t.type !== "면제" || (!t.cond && !(t.sources || []).length)) return "";
    let s = "\n[면제 작성 조건]\n";
    if (t.cond) s += `- 면제조건은 다음 문구로 확정되어 있다(그대로 사용): "${t.cond}"\n`;
    if ((t.sources || []).length) {
      const names = t.sources.map((x) => `${x.no}. ${srcLabel(x)}`).join(", ");
      s += `- 면제 근거 자료: ${names}. 근거자료 칸에는 이 서류(시험자료의 출처 포함)를 적고, 결론·요약에는 근거 자료의 시험방법과 주요 결과(예: NOAEL, 판정)를 인용해 면제 사유를 서술한다. 첨부 1에는 "${t.sources.map(srcLabel).join(", ")}"을 쓴다.\n`;
      if (manual) s += `- 근거 자료의 내용은 이 답변의 항목 ${t.sources.map((x) => x.no).join(", ")} 작성 내용과 일치시킨다.\n`;
      else {
        const done = t.sources.filter((x) => x.status === "ok");
        if (done.length) s += "\n[근거 서류 작성 결과]\n" + done.map(sourceSummary).join("\n") + "\n";
        else s += "- 근거 서류 작성에 실패했으므로 업로드 문서에서 직접 근거를 찾아 작성하고 notes에 그 사실을 적는다.\n";
      }
    }
    return s;
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

  async function callGemini(key, model, system, content) {
    const parts = content.map((c) => c.type === "document"
      ? { inline_data: { mime_type: c.source.media_type, data: c.source.data } }
      : { text: c.text });
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model || "gemini-3.8-flash")}:generateContent`;
    let res;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        res = await fetch(url, {
          method: "POST",
          headers: { "content-type": "application/json", "x-goog-api-key": key },
          body: JSON.stringify({
            systemInstruction: { parts: [{ text: system }] },
            contents: [{ role: "user", parts }],
            generationConfig: { responseMimeType: "application/json", maxOutputTokens: 16000, temperature: 0.2 },
          }),
        });
      } catch (e) {
        throw new Error("Gemini API에 연결하지 못했습니다. 인터넷/사내 방화벽을 확인하세요. (" + e.message + ")");
      }
      if ((res.status === 429 || res.status === 503 || res.status >= 500) && attempt < 2) {
        log(`Gemini 사용량 제한/혼잡으로 ${20 * (attempt + 1)}초 뒤 다시 시도합니다…`);
        await new Promise((r) => setTimeout(r, 20000 * (attempt + 1)));
        continue;
      }
      break;
    }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      const m = (data.error && data.error.message) || res.statusText;
      if (res.status === 400 && /API key/i.test(m)) throw new Error("Gemini API 키가 올바르지 않습니다. (" + m + ")");
      if (res.status === 403) throw new Error("Gemini API 키 권한 오류입니다. (" + m + ")");
      if (res.status === 404) throw new Error("Gemini 모델 이름을 확인하세요: " + model + " (" + m + ")");
      if (res.status === 429) throw new Error("Gemini 무료 사용량을 초과했습니다. 잠시 후(또는 내일) 다시 시도하세요. (" + m + ")");
      throw new Error(`Gemini 오류 ${res.status}: ${m}`);
    }
    const cand = (data.candidates || [])[0];
    if (!cand) throw new Error("Gemini가 응답을 거부했습니다. " + JSON.stringify(data.promptFeedback || {}));
    if (cand.finishReason === "MAX_TOKENS") log("⚠ 응답이 길어 일부가 잘렸을 수 있습니다.", "err");
    return ((cand.content && cand.content.parts) || []).map((p) => p.text || "").join("");
  }

  // ---------------- Claude 앱(구독) 모드: 복사 → 붙여넣기 ----------------
  function buildManualText(content, ok) {
    const sys = SYSTEM_PROMPT.split("[출력 형식]")[0];
    const docs = content.filter((c) => c.type === "text").map((c) => c.text).join("\n\n");
    const items = [...tasks].sort((a, b) => a.level - b.level || a.no - b.no).map((t) => `\n==================== 항목 ${t.no} ====================\n${buildPrompt(t, ok, true).replace(/\n위 입력칸을 모두 검토해[\s\S]*$/, "")}`).join("\n");
    return `${sys}
[출력 형식] 아래 JSON 하나만 출력한다. 다른 설명이나 코드블록 표시는 쓰지 않는다. 항목이 많아 답이 길어지면 끝까지 이어서 출력한다.
{"items": {"항목번호": {"fields": {"셀주소": "값", ...}, "evidence": [{"file": 문서번호, "pages": [번호, ...]}], "notes": "검토 필요사항"}, ...}}
작성할 항목번호: ${tasks.map((t) => t.no).join(", ")}

==================== 자료 ====================
${docs}
${items}

모든 항목을 위 [출력 형식]의 JSON 하나로 답하라.`;
  }

  function renderManual(content, ok, attach) {
    const text = buildManualText(content, ok);
    tasks.forEach((t) => { t.status = "wait"; updateTaskHead(t); });
    const box = document.createElement("div");
    box.className = "manual-box"; box.id = "manualBox";
    box.innerHTML = `<h3>Claude 앱으로 작성 (${tasks.length}개 항목 · 요청문 ${text.length.toLocaleString()}자)</h3>
      <div class="steps">
        <button class="primary" id="mCopy">① 요청문 복사</button>
        <a href="https://claude.ai/new" target="_blank" rel="noopener"><button type="button">② claude.ai 새 대화 열기</button></a>
        <button id="mSave">요청문 파일로 저장(.txt)</button>
      </div>
      ${attach.length ? `<p class="warn-list">⚠ 스캔본이라 글자를 못 읽은 파일 — claude.ai 대화창에 직접 첨부하세요: ${attach.map(esc).join(", ")}</p>` : ""}
      ${text.length > 150000 ? `<p class="warn-list">⚠ 요청문이 깁니다. 붙여넣기가 안 되면 ‘파일로 저장’한 .txt를 대화창에 첨부하세요.</p>` : ""}
      <label>③ Claude의 답을 전부 복사해 여기에 붙여넣기
        <textarea id="mAnswer" placeholder='{"items": {...}} 형태의 답변'></textarea></label>
      <div class="steps" style="margin-top:8px"><button class="primary" id="mApply">④ 적용</button><span class="hint" id="mMsg"></span></div>`;
    const old = $("manualBox"); if (old) old.remove();
    $("results").before(box);
    $("mCopy").addEventListener("click", async () => {
      try { await navigator.clipboard.writeText(text); $("mCopy").textContent = "✔ 복사됨 — claude.ai에 붙여넣으세요"; }
      catch (e) { save(new Blob(["\ufeff" + text], { type: "text/plain" }), "작성요청문.txt"); alert("클립보드 복사가 막혀 있어 파일로 저장했습니다. 이 파일을 대화창에 첨부하세요."); }
    });
    $("mSave").addEventListener("click", () => save(new Blob(["\ufeff" + text], { type: "text/plain" }), `작성요청문_${substId()}.txt`));
    $("mApply").addEventListener("click", () => {
      try {
        const json = parseJson($("mAnswer").value);
        const items = json.items || (tasks.length === 1 && json.fields ? { [tasks[0].no]: json } : null);
        if (!items) throw new Error("답변에 items가 없습니다. Claude의 답 전체를 붙여넣었는지 확인하세요.");
        let n = 0;
        for (const t of [...tasks].sort((a, b) => a.level - b.level)) {
          const it = items[t.no] || items[String(t.no)];
          if (it) { applyJson(t, it, ok); n++; }
          else if (t.status !== "ok") { t.status = "err"; t.error = "답변에 이 항목이 없습니다. Claude에게 '항목 " + t.no + "도 이어서 작성해줘'라고 요청한 뒤 다시 붙여넣으세요."; }
          renderTaskBody(t, ok);
        }
        $("mMsg").textContent = `${n}/${tasks.length}개 항목 적용됨`;
        log(`Claude 앱 답변 적용: ${n}/${tasks.length}개 항목`, n === tasks.length ? "ok" : "err");
      } catch (e) { $("mMsg").textContent = ""; alert(e.message); }
    });
    log("요청문을 만들었습니다. 4번 결과 칸의 ①~④ 순서대로 진행하세요.", "ok");
    $("sec-result").scrollIntoView({ behavior: "smooth" });
  }

  function parseJson(text) {
    const a = text.indexOf("{"), b = text.lastIndexOf("}");
    if (a < 0 || b < 0) throw new Error("AI 응답에서 결과(JSON)를 찾지 못했습니다.");
    try { return JSON.parse(text.slice(a, b + 1)); }
    catch (e) { throw new Error("AI 응답 형식이 올바르지 않습니다. 다시 시도해 주세요."); }
  }

  // ---------------- 결과 화면 ----------------
  const STATUS = { wait: ["대기 중", "wait"], run: ["작성 중…", "wait"], ok: ["완료", "ok"], err: ["오류", "err"] };

  // 결과는 내용 없이 항목별 다운로드 버튼만 표시
  function renderResults(ok) {
    $("results").innerHTML = tasks.map((t) => `<div class="res" id="${t.id}">
      <div class="res-name"><strong>${t.no}. ${esc(t.cat.name)}</strong> <span class="res-doc">${DOCNAME[t.type]}</span>${t.sources.length ? `<span class="link-badge">근거: ${t.sources.map((x) => x.no + "번").join(", ")}</span>` : ""}</div>
      <span class="st"></span>
      <div class="res-act"><button data-act="xlsx" disabled>엑셀</button><button data-act="pdf" disabled>증빙자료</button></div>
    </div>`).join("");
    for (const t of tasks) {
      const row = $(t.id);
      row.querySelector('[data-act="xlsx"]').addEventListener("click", () => downloadTaskXlsx(t));
      row.querySelector('[data-act="pdf"]').addEventListener("click", () => downloadTaskPdf(t));
      updateTaskHead(t);
    }
  }
  function updateTaskHead(t) {
    const row = $(t.id);
    if (!row) return;
    const el = row.querySelector(".st");
    const [txt, cls] = STATUS[t.status];
    el.textContent = txt; el.className = "st " + cls;
    el.title = t.status === "err" ? t.error || "" : t.status === "ok" && t.notes ? "AI 검토 메모: " + t.notes : "";
    if (t.status === "ok" && t.notes) el.textContent = "완료 · 검토 메모 있음";
    const done = t.status === "ok";
    row.querySelector('[data-act="xlsx"]').disabled = !done;
    row.querySelector('[data-act="pdf"]').disabled = !done || !hasEvidence(t);
  }
  function renderTaskBody(t) { updateTaskHead(t); }
  const hasEvidence = (t) => Object.values(t.evidence || {}).some((set) => set.size);

  function collectValues(t) {
    const vals = {};
    for (const f of t.sheet.fields) {
      const v = t.values[f.cell];
      if (v !== undefined && v !== (f.default ?? "")) vals[f.cell] = v;
    }
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
    for (const [fid, set] of Object.entries(t.evidence || {})) {
      const f = files.find((x) => x.id === +fid);
      if (!f) continue;
      const pages = [...set].filter((p) => p >= 1 && p <= f.numPages).sort((a, b) => a - b);
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
      const memo = done.filter((t) => t.notes).map((t) => `[${t.no}. ${t.cat.name} · ${DOCNAME[t.type]}]\r\n${t.notes}\r\n`);
      if (memo.length) zip.file("AI_검토메모.txt", "\ufeff" + memo.join("\r\n"));
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
