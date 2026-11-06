/* 엑셀 양식의 서식을 그대로 두고 셀 값만 채우는 모듈 (JSZip 사용) */
(function (global) {
  const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const colNum = (c) => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const splitRef = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return [m[1], +m[2]]; };
  const parse = (s) => new DOMParser().parseFromString(s, "application/xml");
  const ser = (d) => new XMLSerializer().serializeToString(d);
  const kids = (el, name) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === name);

  async function fillWorkbook(arrayBuffer, sheetName, values) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const wb = parse(await zip.file("xl/workbook.xml").async("string"));
    const rels = parse(await zip.file("xl/_rels/workbook.xml.rels").async("string"));
    const relMap = {};
    for (const r of rels.getElementsByTagName("Relationship")) {
      const t = r.getAttribute("Target");
      relMap[r.getAttribute("Id")] = t.startsWith("/") ? t.slice(1) : "xl/" + t;
    }
    const sheetEls = Array.from(wb.getElementsByTagNameNS(NS, "sheet"));
    const target = sheetEls.find((s) => s.getAttribute("name") === sheetName);
    if (!target) throw new Error("양식에 시트가 없습니다: " + sheetName);
    const targetIdx = sheetEls.indexOf(target);

    // 1) 대상 시트 셀 채우기
    const path = relMap[target.getAttributeNS(RNS, "id")];
    const ws = parse(await zip.file(path).async("string"));
    const sheetData = ws.getElementsByTagNameNS(NS, "sheetData")[0];
    let touchedFormula = false;

    function getRow(r) {
      for (const row of kids(sheetData, "row")) {
        const rr = +row.getAttribute("r");
        if (rr === r) return row;
        if (rr > r) { const n = ws.createElementNS(NS, "row"); n.setAttribute("r", r); sheetData.insertBefore(n, row); return n; }
      }
      const n = ws.createElementNS(NS, "row"); n.setAttribute("r", r); sheetData.appendChild(n); return n;
    }
    function getCell(ref) {
      const [col, r] = splitRef(ref);
      const row = getRow(r);
      for (const c of kids(row, "c")) {
        const [cc] = splitRef(c.getAttribute("r"));
        if (cc === col) return c;
        if (colNum(cc) > colNum(col)) { const n = ws.createElementNS(NS, "c"); n.setAttribute("r", ref); row.insertBefore(n, c); return n; }
      }
      const n = ws.createElementNS(NS, "c"); n.setAttribute("r", ref); row.appendChild(n); return n;
    }
    for (const [ref, val] of Object.entries(values)) {
      if (val === undefined || val === null) continue;
      const c = getCell(ref);
      for (const tag of ["f", "v", "is"]) for (const e of kids(c, tag)) { if (tag === "f") touchedFormula = true; c.removeChild(e); }
      const text = String(val);
      if (text === "") { c.removeAttribute("t"); continue; }
      const num = /^-?\d+(\.\d+)?$/.test(text) && text.length < 15 && !/^0\d/.test(text);
      if (num) {
        c.removeAttribute("t");
        const v = ws.createElementNS(NS, "v"); v.textContent = text; c.appendChild(v);
      } else {
        c.setAttribute("t", "inlineStr");
        const is = ws.createElementNS(NS, "is"); const t = ws.createElementNS(NS, "t");
        t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
        t.textContent = text; is.appendChild(t); c.appendChild(is);
      }
    }

    // 2) 탭 선택 상태: 대상 시트만 선택
    for (let i = 0; i < sheetEls.length; i++) {
      const p = relMap[sheetEls[i].getAttributeNS(RNS, "id")];
      const doc = i === targetIdx ? ws : parse(await zip.file(p).async("string"));
      let changed = false;
      for (const sv of doc.getElementsByTagNameNS(NS, "sheetView")) {
        if (i === targetIdx) { sv.setAttribute("tabSelected", "1"); changed = true; }
        else if (sv.hasAttribute("tabSelected")) { sv.removeAttribute("tabSelected"); changed = true; }
      }
      if (i !== targetIdx && changed) zip.file(p, ser(doc));
    }
    zip.file(path, ser(ws));

    // 3) 대상 외 시트 숨김(드롭다운 목록 참조가 깨지지 않도록 삭제하지 않음)
    for (const s of sheetEls) {
      if (s === target) s.removeAttribute("state");
      else s.setAttribute("state", "hidden");
    }
    const view = wb.getElementsByTagNameNS(NS, "workbookView")[0];
    if (view) { view.setAttribute("activeTab", targetIdx); view.setAttribute("firstSheet", targetIdx); }
    // 수식(예: 제목 =F8) 열 때 재계산
    let calc = wb.getElementsByTagNameNS(NS, "calcPr")[0];
    if (!calc) {
      calc = wb.createElementNS(NS, "calcPr");
      const after = ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];
      const anchor = Array.from(wb.documentElement.childNodes).find((n) => n.nodeType === 1 && after.includes(n.localName));
      wb.documentElement.insertBefore(calc, anchor || null);
    }
    calc.setAttribute("fullCalcOnLoad", "1");
    zip.file("xl/workbook.xml", ser(wb));

    if (touchedFormula && zip.file("xl/calcChain.xml")) {
      zip.remove("xl/calcChain.xml");
      let r = await zip.file("xl/_rels/workbook.xml.rels").async("string");
      zip.file("xl/_rels/workbook.xml.rels", r.replace(/<Relationship [^>]*calcChain[^>]*\/>/, ""));
      let ct = await zip.file("[Content_Types].xml").async("string");
      zip.file("[Content_Types].xml", ct.replace(/<Override [^>]*calcChain[^>]*\/>/, ""));
    }
    return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", compression: "DEFLATE" });
  }

  global.XlsxFill = { fillWorkbook };
})(window);
