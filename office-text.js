/* 워드(.docx)·엑셀(.xlsx)·파워포인트(.pptx)에서 쪽/시트/슬라이드 단위로 텍스트 추출 (JSZip 사용)
   반환: { pages: [텍스트...], names: [단위 이름...] } */
(function (global) {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const parse = (s) => new DOMParser().parseFromString(s, "application/xml");
  const tidy = (s) => s.replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim();

  async function relsOf(zip, path) {
    const dir = path.substring(0, path.lastIndexOf("/") + 1);
    const relPath = dir + "_rels/" + path.substring(path.lastIndexOf("/") + 1) + ".rels";
    const f = zip.file(relPath);
    const map = {};
    if (!f) return map;
    const doc = parse(await f.async("string"));
    for (const r of doc.getElementsByTagName("Relationship")) {
      let t = r.getAttribute("Target");
      if (t.startsWith("/")) t = t.slice(1);
      else {
        const parts = (dir + t).split("/"), out = [];
        for (const p of parts) { if (p === "..") out.pop(); else if (p !== ".") out.push(p); }
        t = out.join("/");
      }
      map[r.getAttribute("Id")] = t;
    }
    return map;
  }

  // ---- Word: 페이지 구분(수동 쪽나눔 + Word가 저장한 마지막 렌더링 쪽 경계) 기준으로 나눔 ----
  async function docx(zip) {
    const doc = parse(await zip.file("word/document.xml").async("string"));
    const body = doc.getElementsByTagNameNS(W, "body")[0];
    const pages = [""];
    const walk = (node) => {
      for (const n of node.childNodes) {
        if (n.nodeType !== 1) continue;
        const ln = n.localName;
        if (n.namespaceURI === W) {
          if (ln === "lastRenderedPageBreak" || (ln === "br" && n.getAttributeNS(W, "type") === "page")) { pages.push(""); continue; }
          if (ln === "t") { pages[pages.length - 1] += n.textContent; continue; }
          if (ln === "tab") { pages[pages.length - 1] += "\t"; continue; }
          if (ln === "br" || ln === "cr") { pages[pages.length - 1] += "\n"; continue; }
          if (ln === "delText" || ln === "instrText") continue;
        }
        walk(n);
        if (n.namespaceURI === W && ln === "p") pages[pages.length - 1] += "\n";
        if (n.namespaceURI === W && ln === "tc") pages[pages.length - 1] = pages[pages.length - 1].replace(/\n+$/, "") + " | ";
        if (n.namespaceURI === W && ln === "tr") pages[pages.length - 1] = pages[pages.length - 1].replace(/ \| $/, "") + "\n";
      }
    };
    walk(body);
    let out = pages.map(tidy);
    // 쪽 경계 정보가 없는 문서(다른 프로그램에서 저장 등)는 약 2,500자 단위로 나눔
    if (out.length === 1 && out[0].length > 4000) {
      const t = out[0], chunks = [];
      let i = 0;
      while (i < t.length) {
        let j = Math.min(t.length, i + 2500);
        const nl = t.lastIndexOf("\n", j);
        if (j < t.length && nl > i + 1000) j = nl;
        chunks.push(t.slice(i, j).trim()); i = j;
      }
      out = chunks;
      return { pages: out, names: out.map((_, k) => `구간 ${k + 1}`), approx: "chunk" };
    }
    while (out.length > 1 && !out[out.length - 1]) out.pop();
    return { pages: out, names: out.map((_, k) => `${k + 1}쪽`), approx: "page" };
  }

  // ---- Excel: 시트 1개 = 1단위, 행은 줄바꿈, 셀은 탭 ----
  async function xlsx(zip) {
    const sst = [];
    if (zip.file("xl/sharedStrings.xml")) {
      const d = parse(await zip.file("xl/sharedStrings.xml").async("string"));
      for (const si of d.getElementsByTagNameNS(S, "si")) {
        sst.push(Array.from(si.getElementsByTagNameNS(S, "t")).filter((t) => t.parentNode.localName !== "rPh").map((t) => t.textContent).join(""));
      }
    }
    const wb = parse(await zip.file("xl/workbook.xml").async("string"));
    const rels = await relsOf(zip, "xl/workbook.xml");
    const pages = [], names = [];
    for (const s of wb.getElementsByTagNameNS(S, "sheet")) {
      const name = s.getAttribute("name");
      const path = rels[s.getAttributeNS(R, "id")];
      if (!path || !zip.file(path)) continue;
      const ws = parse(await zip.file(path).async("string"));
      const lines = [];
      for (const row of ws.getElementsByTagNameNS(S, "row")) {
        const cells = [];
        for (const c of row.getElementsByTagNameNS(S, "c")) {
          const t = c.getAttribute("t");
          let v = "";
          if (t === "s") { const x = c.getElementsByTagNameNS(S, "v")[0]; v = x ? sst[+x.textContent] || "" : ""; }
          else if (t === "inlineStr") v = Array.from(c.getElementsByTagNameNS(S, "t")).map((x) => x.textContent).join("");
          else { const x = c.getElementsByTagNameNS(S, "v")[0]; v = x ? x.textContent : ""; }
          if (v !== "") cells.push(`${c.getAttribute("r")}: ${v.replace(/\s+/g, " ")}`);
        }
        if (cells.length) lines.push(cells.join("\t"));
        if (lines.join("\n").length > 60000) { lines.push("…(이하 생략)"); break; }
      }
      pages.push(tidy(lines.join("\n")));
      names.push(`시트 '${name}'`);
    }
    return { pages, names, approx: "sheet" };
  }

  // ---- PowerPoint: 발표 순서대로 슬라이드 1장 = 1단위 (슬라이드 노트 포함) ----
  async function pptx(zip) {
    const pres = parse(await zip.file("ppt/presentation.xml").async("string"));
    const rels = await relsOf(zip, "ppt/presentation.xml");
    const pages = [], names = [];
    let k = 0;
    for (const id of pres.getElementsByTagNameNS(P, "sldId")) {
      const path = rels[id.getAttributeNS(R, "id")];
      if (!path || !zip.file(path)) continue;
      k++;
      const sl = parse(await zip.file(path).async("string"));
      const paras = Array.from(sl.getElementsByTagNameNS(A, "p")).map((p) =>
        Array.from(p.getElementsByTagNameNS(A, "t")).map((t) => t.textContent).join(""));
      let text = paras.filter(Boolean).join("\n");
      const srels = await relsOf(zip, path);
      const notesPath = Object.values(srels).find((t) => /notesSlide/.test(t));
      if (notesPath && zip.file(notesPath)) {
        const nd = parse(await zip.file(notesPath).async("string"));
        const notes = Array.from(nd.getElementsByTagNameNS(A, "p")).map((p) =>
          Array.from(p.getElementsByTagNameNS(A, "t")).map((t) => t.textContent).join("")).filter(Boolean).join("\n");
        if (notes.trim() && notes.trim() !== String(k)) text += "\n[슬라이드 노트]\n" + notes;
      }
      pages.push(tidy(text));
      names.push(`슬라이드 ${k}`);
    }
    return { pages, names, approx: "slide" };
  }

  async function extract(arrayBuffer, kind) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    if (kind === "docx") return docx(zip);
    if (kind === "xlsx") return xlsx(zip);
    if (kind === "pptx") return pptx(zip);
    throw new Error("지원하지 않는 형식");
  }

  global.OfficeText = { extract };
})(window);
