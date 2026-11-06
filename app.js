/* 등록서류 작성 도우미 - 메인 로직 */
(function () {
  "use strict";
  pdfjsLib.GlobalWorkerOptions.workerSrc = "vendor/pdf.worker.min.js";

  const TEXT_BUDGET = 320000;      // AI에 보낼 본문 최대 글자 수
  const SCANNED_CHARS = 60;        // 쪽당 평균 글자 수가 이보다 적으면 스캔본으로 판단
  const $ = (id) => document.getElementById(id);
  const store = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} },
  };

  let CONFIG = null;
  let files = [];         // {id, name, bytes, numPages, pages[], scanned}
  let result = null;      // {item, sheet, fields, values, evidence}
  let fileSeq = 0;

  // ---------------- 초기화 ----------------
  async function init() {
    const k = store.get("regdoc.apiKey");
    if (k) { $("apiKey").value = k; $("rememberKey").checked = true; }
    const m = store.get("regdoc.model"); if (m) $("model").value = m;
    $("apiKey").addEventListener("change", saveKey);
    $("rememberKey").addEventListener("change", saveKey);
    $("model").addEventListener("change", () => store.set("regdoc.model", $("model").value.trim()));

    CONFIG = await (await fetch("config/templates.json", { cache: "no-cache" })).json();
    const groups = [...new Set(CONFIG.items.map((i) => i.group))];
    $("group").innerHTML = groups.map((g) => `<option>${g}</option>`).join("");
    $("group").addEventListener("change", fillItems);
    $("item").addEventListener("change", fillTypes);
    $("type").addEventListener("change", showSheet);
    fillItems();

    const drop = $("drop");
    drop.addEventListener("click", () => $("fileInput").click());
    drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") $("fileInput").click(); });
    drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
    drop.addEventListener("dragleave", () => drop.classList.remove("over"));
    drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("over"); addFiles(e.dataTransfer.files); });
    $("fileInput").addEventListener("change", (e) => { addFiles(e.target.files); e.target.value = ""; });
    $("runBtn").addEventListener("click", run);
    $("dlXlsx").addEventListener("click", downloadXlsx);
    $("dlPdf").addEventListener("click", downloadPdf);
  }

  function saveKey() {
    if ($("rememberKey").checked) store.set("regdoc.apiKey", $("apiKey").value.trim());
    else store.del("regdoc.apiKey");
  }

  function fillItems() {
    const g = $("group").value;
    const names = [...new Set(CONFIG.items.filter((i) => i.group === g).map((i) => i.name))];
    $("item").innerHTML = names.map((n) => `<option>${n}</option>`).join("");
    fillTypes();
  }
  function fillTypes() {
    const g = $("group").value, n = $("item").value;
    const types = CONFIG.items.filter((i) => i.group === g && i.name === n).map((i) => i.type);
    $("type").innerHTML = types.map((t) => `<option value="${t}">${t === "제출" ? "제출 (시험요약서)" : t === "생략" ? "생략 (생략사유서)" : "면제 (면제사유서)"}</option>`).join("");
    showSheet();
  }
  function currentItem() {
    return CONFIG.items.find((i) => i.group === $("group").value && i.name === $("item").value && i.type === $("type").value);
  }
  function showSheet() {
    const it = currentItem();
    if (it) $("sheetInfo").textContent = `사용 양식: ${fileLabel(it.file)} › 시트 "${it.sheet}" (입력칸 ${CONFIG.sheets[it.key].fields.length}개)`;
  }
  const fileLabel = (k) => ({ phys: "시험요약서_물화", human: "시험요약서_인체", env: "시험요약서_환경" }[k] || k);

  // ---------------- 파일 ----------------
  async function addFiles(list) {
    for (const f of list) {
      if (!/\.pdf$/i.test(f.name) && f.type !== "application/pdf") { alert(`PDF만 올릴 수 있습니다: ${f.name}`); continue; }
      const entry = { id: ++fileSeq, name: f.name, bytes: await f.arrayBuffer(), numPages: 0, pages: [], scanned: false, loading: true };
      files.push(entry);
      renderFiles();
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
      `${f.numPages}쪽${f.scanned ? " · 스캔본(이미지)" : ""}`}</span><button title="삭제" data-id="${f.id}">✕</button></li>`).join("");
    $("fileList").querySelectorAll("button").forEach((b) => b.addEventListener("click", () => {
      files = files.filter((f) => f.id !== +b.dataset.id); renderFiles();
    }));
    $("runBtn").disabled = !files.some((f) => !f.loading && !f.error);
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
    const it = currentItem();
    const sheet = CONFIG.sheets[it.key];
    const ok = files.filter((f) => !f.loading && !f.error);
    $("log").innerHTML = ""; $("sec-result").classList.add("hidden");
    $("runBtn").disabled = true;
    try {
      log(`양식: ${fileLabel(it.file)} › ${it.sheet}`);
      const { content, sentInfo } = buildDocuments(ok, it);
      sentInfo.forEach((s) => log(s));
      const prompt = buildPrompt(it, sheet, ok);
      log("AI가 자료를 분석하고 있습니다… (보통 1~3분)");
      const t0 = Date.now();
      const text = await callClaude(key, $("model").value.trim(), SYSTEM_PROMPT, [...content, { type: "text", text: prompt }]);
      log(`AI 응답 수신 (${Math.round((Date.now() - t0) / 1000)}초)`, "ok");
      const json = parseJson(text);
      result = { item: it, sheet, values: {}, evidence: {}, notes: json.notes || "" };
      for (const f of sheet.fields) {
        let v = json.fields && Object.prototype.hasOwnProperty.call(json.fields, f.cell) ? json.fields[f.cell] : undefined;
        if (v === undefined || v === null) v = f.default ?? "";
        result.values[f.cell] = String(v);
      }
      // 사용자가 입력한 물질명/CAS 우선
      for (const f of sheet.fields) {
        if (f.label === "화학물질명" && $("chemName").value.trim()) result.values[f.cell] = $("chemName").value.trim();
        if (f.label === "CAS No." && $("cas").value.trim()) result.values[f.cell] = $("cas").value.trim();
      }
      for (const f of ok) result.evidence[f.id] = new Set();
      for (const e of json.evidence || []) {
        const file = ok[(e.file || 1) - 1];
        if (!file) continue;
        for (const p of e.pages || []) if (p >= 1 && p <= file.numPages) result.evidence[file.id].add(+p);
      }
      renderResult(ok);
      log("완료되었습니다. 아래에서 확인 후 다운로드하세요.", "ok");
    } catch (e) {
      console.error(e);
      log("오류: " + e.message, "err");
    } finally {
      $("runBtn").disabled = false;
    }
  }

  function buildDocuments(list, it) {
    const kw = [...it.keywords, $("chemName").value.trim(), $("cas").value.trim()].filter(Boolean).map((k) => k.toLowerCase());
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
      let pages = f.pages.map((t, idx) => ({ p: idx + 1, t }));
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
        sentInfo.push(`문서 ${no} "${f.name}": ${f.numPages}쪽 중 관련도 높은 ${selected.length}쪽만 전송 (분량 제한)`);
      } else {
        sentInfo.push(`문서 ${no} "${f.name}": ${f.numPages}쪽 전체 전송 (${total.toLocaleString()}자)`);
      }
      const body = selected.map((x) => `[p.${x.p}]\n${x.t}`).join("\n\n");
      content.push({ type: "text", text: `<문서 ${no}: ${f.name} (총 ${f.numPages}쪽)>\n${body}\n</문서 ${no}>` });
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
8. 근거로 사용한 쪽 번호를 evidence에 모두 적는다. 쪽 번호는 문서 본문의 [p.N] 표기(PDF 물리적 쪽 번호)를 쓴다. 표지는 자동으로 포함되므로 적지 않아도 된다.

[출력 형식] 아래 JSON 하나만 출력한다. 다른 설명이나 코드블록 표시는 쓰지 않는다.
{"fields": {"셀주소": "값", ...}, "evidence": [{"file": 문서번호, "pages": [쪽번호, ...]}], "notes": "검토자가 확인해야 할 사항(자료에서 못 찾은 항목, 판단이 애매한 부분 등)을 한국어로 간단히"}`;

  function buildPrompt(it, sheet, list) {
    const name = $("chemName").value.trim(), cas = $("cas").value.trim();
    const typeDesc = { 제출: "시험자료 국문요약서", 생략: "등록제출자료 생략사유서", 면제: "자료 면제사유서" }[it.type];
    const lines = sheet.fields.map((f) => {
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
- 시험항목: ${it.name} (${it.group})
- 작성 유형: ${it.type} → ${typeDesc}
- 양식 시트: ${it.sheet} (제목: ${sheet.title})
- 업로드 문서: ${list.map((f, i) => `문서 ${i + 1} "${f.name}"`).join(", ")}

[입력칸 목록] (셀주소 | 위치 | 항목 | 선택지 | 기본값 | 작성예시)
${lines}

위 입력칸을 모두 검토해 JSON으로 답하라. fields에는 위 목록의 셀주소를 키로 쓴다.`;
  }
  const clip = (s, n) => { s = String(s).replace(/\n/g, "⏎"); return s.length > n ? s.slice(0, n) + "…" : s; };

  async function callClaude(key, model, system, content) {
    let res;
    try {
      res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-api-key": key,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({ model: model || "claude-sonnet-4-5", max_tokens: 16000, system, messages: [{ role: "user", content }] }),
      });
    } catch (e) {
      throw new Error("Anthropic API에 연결하지 못했습니다. 인터넷/사내 방화벽을 확인하세요. (" + e.message + ")");
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
  function renderResult(list) {
    const { sheet } = result;
    const tb = $("resultTable").querySelector("tbody");
    let html = "", lastSec = null;
    sheet.fields.forEach((f, i) => {
      const sec = f.section || "";
      if (sec && sec !== lastSec) { html += `<tr class="sec"><td colspan="3">${esc(sec)}</td></tr>`; lastSec = sec; }
      const v = result.values[f.cell] ?? "";
      const dl = f.options && f.options.length ? `<datalist id="dl${i}">${f.options.map((o) => `<option value="${esc(o)}">`).join("")}</datalist>` : "";
      const input = f.options && f.options.length && !f.tall && v.length < 120 && !v.includes("\n")
        ? `<input type="text" list="dl${i}" data-cell="${f.cell}" value="${esc(v)}">${dl}`
        : `<textarea data-cell="${f.cell}" rows="${Math.min(10, Math.max(1, Math.ceil(v.length / 60) + (v.match(/\n/g) || []).length))}">${esc(v)}</textarea>`;
      html += `<tr><td class="cell">${f.cell}</td><td class="lbl">${esc(f.label)}${f.sub ? `<small>${esc(f.sub)}</small>` : ""}</td><td>${input}</td></tr>`;
    });
    tb.innerHTML = html;
    tb.querySelectorAll("[data-cell]").forEach((el) => { el.addEventListener("input", checkOption); checkOption({ target: el }); });

    const notes = $("aiNotes");
    if (result.notes) { notes.textContent = "AI 검토 메모\n" + result.notes; notes.classList.remove("hidden"); } else notes.classList.add("hidden");

    $("evidence").innerHTML = "<strong>증빙자료 발췌 쪽</strong> (표지 1쪽 자동 포함, 수정 가능 · 예: 1, 12-14)" + list.map((f) => {
      const pages = [...result.evidence[f.id]].sort((a, b) => a - b);
      return `<label style="margin-top:6px">${esc(f.name)} (총 ${f.numPages}쪽)<input type="text" data-file="${f.id}" value="${pages.join(", ")}" placeholder="인용하지 않은 문서는 비워 두면 제외"></label>`;
    }).join("");
    $("sec-result").classList.remove("hidden");
    $("sec-result").scrollIntoView({ behavior: "smooth" });
  }
  function checkOption(e) {
    const el = e.target, f = result.sheet.fields.find((x) => x.cell === el.dataset.cell);
    const v = el.value.trim();
    const bad = f.options && f.options.length && v && !f.options.includes(v) && !/^other/i.test(v) && !/^기타/.test(v);
    el.classList.toggle("warn", !!bad);
  }

  function collectValues() {
    const vals = {};
    document.querySelectorAll("#resultTable [data-cell]").forEach((el) => {
      const f = result.sheet.fields.find((x) => x.cell === el.dataset.cell);
      const v = el.value;
      if (v !== (f.default ?? "")) vals[f.cell] = v;
    });
    return vals;
  }

  const safe = (s) => String(s || "").replace(/[\\/:*?"<>|\s]+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
  function baseName() {
    const it = result.item;
    const kind = { 제출: "시험요약서", 생략: "생략사유서", 면제: "면제사유서" }[it.type];
    return `${kind}_${safe(it.sheet)}_${safe($("cas").value.trim() || $("chemName").value.trim() || "물질")}`;
  }

  async function downloadXlsx() {
    try {
      const it = result.item;
      const buf = await (await fetch(CONFIG.files[it.file], { cache: "no-cache" })).arrayBuffer();
      const blob = await XlsxFill.fillWorkbook(buf, it.sheet, collectValues());
      save(blob, baseName() + ".xlsx");
    } catch (e) { alert("엑셀 생성 실패: " + e.message); console.error(e); }
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

  async function downloadPdf() {
    try {
      const out = await PDFLib.PDFDocument.create();
      let added = 0;
      for (const el of document.querySelectorAll("#evidence [data-file]")) {
        const f = files.find((x) => x.id === +el.dataset.file);
        if (!f) continue;
        let pages = parsePages(el.value, f.numPages);
        if (!pages.length) continue;
        if (!pages.includes(1)) pages.unshift(1);
        const src = await PDFLib.PDFDocument.load(f.bytes.slice(0), { ignoreEncryption: true });
        const copied = await out.copyPages(src, pages.map((p) => p - 1));
        copied.forEach((pg) => out.addPage(pg)); added += copied.length;
      }
      if (!added) { alert("발췌할 쪽이 없습니다. 쪽 번호를 입력하세요."); return; }
      save(new Blob([await out.save()], { type: "application/pdf" }), "증빙_" + baseName() + ".pdf");
    } catch (e) { alert("PDF 생성 실패: " + e.message); console.error(e); }
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
