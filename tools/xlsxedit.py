"""서식을 보존하면서 xlsx 내부 XML을 직접 수정하는 최소 도구."""
import re, zipfile, copy
from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
N = {"m": NS, "r": RNS, "x14": X14, "xm": XM}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

AFTER_DV = ["hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
            "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
            "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]


def col2n(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.match(r"\$?([A-Z]+)\$?(\d+)", ref)
    return m.group(1), int(m.group(2))


def quote(name):
    return "'" + name.replace("'", "''") + "'"


class Book:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.order = z.namelist()
        self.files = {n: z.read(n) for n in self.order}
        self.infos = {i.filename: i for i in z.infolist()}
        self.trees = {}
        self.wb = self.tree("xl/workbook.xml")
        rels = etree.fromstring(self.files["xl/_rels/workbook.xml.rels"])
        self.rid2path = {}
        for r in rels:
            t = r.get("Target")
            self.rid2path[r.get("Id")] = ("xl/" + t) if not t.startswith("/") else t.lstrip("/")
        self.drop_calcchain = False
        self.sst = None
        if "xl/sharedStrings.xml" in self.files:
            self.sst = self.tree("xl/sharedStrings.xml")

    # ---------- 기본 ----------
    def tree(self, name):
        if name not in self.trees:
            self.trees[name] = etree.fromstring(self.files[name])
        return self.trees[name]

    def sheets(self):
        return self.wb.find("m:sheets", N)

    def sheet_names(self):
        return [s.get("name") for s in self.sheets()]

    def sheet_el(self, name):
        for s in self.sheets():
            if s.get("name") == name:
                return s
        raise KeyError(name)

    def ws(self, name):
        rid = self.sheet_el(name).get("{%s}id" % RNS)
        return self.tree(self.rid2path[rid])

    def ws_path(self, name):
        return self.rid2path[self.sheet_el(name).get("{%s}id" % RNS)]

    # ---------- 셀 ----------
    def _row(self, ws, r, create=True):
        sd = ws.find("m:sheetData", N)
        for row in sd.findall("m:row", N):
            rr = int(row.get("r"))
            if rr == r:
                return row
            if rr > r:
                if not create:
                    return None
                new = etree.Element("{%s}row" % NS, r=str(r))
                row.addprevious(new)
                return new
        if not create:
            return None
        new = etree.SubElement(sd, "{%s}row" % NS, r=str(r))
        return new

    def _cell(self, ws, ref, create=True):
        col, r = split_ref(ref)
        row = self._row(ws, r, create)
        if row is None:
            return None
        for c in row.findall("m:c", N):
            cc, _ = split_ref(c.get("r"))
            if cc == col:
                return c
            if col2n(cc) > col2n(col):
                if not create:
                    return None
                new = etree.Element("{%s}c" % NS, r=f"{col}{r}")
                # 같은 행 왼쪽 셀 스타일 상속은 하지 않음(서식 유지)
                c.addprevious(new)
                return new
        if not create:
            return None
        return etree.SubElement(row, "{%s}c" % NS, r=f"{col}{r}")

    def get(self, sheet, ref):
        c = self._cell(self.ws(sheet), ref, create=False)
        if c is None:
            return None
        t = c.get("t")
        if t == "s":
            idx = int(c.find("m:v", N).text)
            si = self.sst.findall("m:si", N)[idx]
            return "".join(x.text or "" for x in si.iter("{%s}t" % NS))
        if t == "inlineStr":
            return "".join(x.text or "" for x in c.iter("{%s}t" % NS))
        v = c.find("m:v", N)
        return v.text if v is not None else None

    def set(self, sheet, ref, text):
        c = self._cell(self.ws(sheet), ref)
        for tag in ("f", "v", "is"):
            for e in c.findall("m:" + tag, N):
                if tag == "f":
                    self.drop_calcchain = True
                c.remove(e)
        if text is None or text == "":
            if "t" in c.attrib:
                del c.attrib["t"]
            return
        c.set("t", "inlineStr")
        is_ = etree.SubElement(c, "{%s}is" % NS)
        t = etree.SubElement(is_, "{%s}t" % NS)
        t.text = text
        t.set(XML_SPACE, "preserve")

    def replace_text(self, pattern, repl, regex=False):
        """공유문자열과 인라인 문자열 전체에서 텍스트 치환."""
        cnt = 0
        targets = []
        if self.sst is not None:
            targets.append(self.sst)
        for s in self.sheet_names():
            targets.append(self.ws(s))
        for tr in targets:
            for t in tr.iter("{%s}t" % NS):
                if t.text:
                    new = re.sub(pattern, repl, t.text) if regex else t.text.replace(pattern, repl)
                    if new != t.text:
                        t.text = new
                        cnt += 1
        return cnt

    # ---------- 데이터 유효성 ----------
    def _dv_elems(self, ws):
        out = []
        dvs = ws.find("m:dataValidations", N)
        if dvs is not None:
            out += [("std", dvs, d) for d in dvs.findall("m:dataValidation", N)]
        for x in ws.iter("{%s}dataValidations" % X14):
            out += [("x14", x, d) for d in x.findall("x14:dataValidation", N)]
        return out

    @staticmethod
    def _sq_get(kind, d):
        if kind == "std":
            return d.get("sqref", "")
        e = d.find("xm:sqref", N)
        return e.text if e is not None else ""

    @staticmethod
    def _sq_set(kind, d, v):
        if kind == "std":
            d.set("sqref", v)
        else:
            d.find("xm:sqref", N).text = v

    def prune_dv(self, sheet, max_row=300, max_col=30):
        ws = self.ws(sheet)
        removed = 0
        for kind, parent, d in self._dv_elems(ws):
            parts = self._sq_get(kind, d).split()
            keep = []
            for p in parts:
                a = p.split(":")
                ok = True
                for x in a:
                    col, r = split_ref(x)
                    if r > max_row or col2n(col) > max_col:
                        ok = False
                if ok:
                    keep.append(p)
            removed += len(parts) - len(keep)
            if keep:
                self._sq_set(kind, d, " ".join(keep))
            else:
                parent.remove(d)
        self._fix_dv_counts(ws)
        return removed

    def _fix_dv_counts(self, ws):
        dvs = ws.find("m:dataValidations", N)
        if dvs is not None:
            n = len(dvs.findall("m:dataValidation", N))
            if n == 0:
                ws.remove(dvs)
            else:
                dvs.set("count", str(n))
        for x in list(ws.iter("{%s}dataValidations" % X14)):
            n = len(x.findall("x14:dataValidation", N))
            if n == 0:
                ext = x.getparent()
                lst = ext.getparent()
                lst.remove(ext)
                if len(lst) == 0:
                    lst.getparent().remove(lst)
            else:
                x.set("count", str(n))

    def _remove_cell_from_dv(self, ws, ref):
        col, r = split_ref(ref)
        for kind, parent, d in self._dv_elems(ws):
            parts = self._sq_get(kind, d).split()
            keep = []
            for p in parts:
                a = p.split(":")
                c1, r1 = split_ref(a[0])
                c2, r2 = split_ref(a[-1])
                if col2n(c1) <= col2n(col) <= col2n(c2) and r1 <= r <= r2:
                    if p == ref:
                        continue
                    if r1 == r2 == r and c1 == c2:
                        continue
                    # 범위를 셀 단위로 풀어서 해당 셀만 제외
                    for rr in range(r1, r2 + 1):
                        for cn in range(col2n(c1), col2n(c2) + 1):
                            cl = ""
                            k = cn
                            while k:
                                k, rem = divmod(k - 1, 26)
                                cl = chr(65 + rem) + cl
                            if not (cl == col and rr == r):
                                keep.append(f"{cl}{rr}")
                else:
                    keep.append(p)
            if keep:
                self._sq_set(kind, d, " ".join(keep))
            else:
                parent.remove(d)
        self._fix_dv_counts(ws)

    def add_list_dv(self, sheet, ref, formula):
        ws = self.ws(sheet)
        self._remove_cell_from_dv(ws, ref)
        dvs = ws.find("m:dataValidations", N)
        if dvs is None:
            dvs = etree.Element("{%s}dataValidations" % NS, count="0")
            anchor = None
            for tag in AFTER_DV:
                anchor = ws.find("m:" + tag, N)
                if anchor is not None:
                    break
            if anchor is not None:
                anchor.addprevious(dvs)
            else:
                ws.append(dvs)
        d = etree.SubElement(dvs, "{%s}dataValidation" % NS, type="list", allowBlank="1",
                             showInputMessage="1", showErrorMessage="1", sqref=ref)
        f = etree.SubElement(d, "{%s}formula1" % NS)
        f.text = formula
        dvs.set("count", str(len(dvs.findall("m:dataValidation", N))))

    def replace_in_formulas(self, old, new):
        for s in self.sheet_names():
            ws = self.ws(s)
            for tag in ("{%s}f" % NS, "{%s}formula1" % NS, "{%s}formula2" % NS, "{%s}f" % XM):
                for e in ws.iter(tag):
                    if e.text and old in e.text:
                        e.text = e.text.replace(old, new)

    # ---------- 시트 ----------
    def rename(self, old, new):
        if old == new:
            return
        self.sheet_el(old).set("name", new)
        pat = re.compile(r"(?:'%s'|(?<![\w'])%s)!" % (re.escape(old.replace("'", "''")), re.escape(old)))
        for s in self.sheet_names():
            ws = self.ws(s)
            for tag in ("{%s}f" % NS, "{%s}formula1" % NS, "{%s}formula2" % NS, "{%s}f" % XM):
                for e in ws.iter(tag):
                    if e.text:
                        e.text = pat.sub(quote(new) + "!", e.text)
        dn = self.wb.find("m:definedNames", N)
        if dn is not None:
            for d in dn:
                d.text = pat.sub(quote(new) + "!", d.text or "")

    def reorder(self, names):
        sheets = self.sheets()
        old = self.sheet_names()
        assert sorted(old) == sorted(names), (set(old) ^ set(names))
        els = {s.get("name"): s for s in sheets}
        for s in list(sheets):
            sheets.remove(s)
        for n in names:
            sheets.append(els[n])
        dn = self.wb.find("m:definedNames", N)
        if dn is not None:
            for d in dn:
                if d.get("localSheetId") is not None:
                    d.set("localSheetId", str(names.index(old[int(d.get("localSheetId"))])))
            # localSheetId 순서대로 정렬(엑셀 권장)
            items = sorted(list(dn), key=lambda d: (d.get("name"), int(d.get("localSheetId") or -1)))
            for d in list(dn):
                dn.remove(d)
            for d in items:
                dn.append(d)

    def hide(self, name, hidden=True):
        el = self.sheet_el(name)
        if hidden:
            el.set("state", "hidden")
        elif "state" in el.attrib:
            del el.attrib["state"]

    def select_first(self):
        names = self.sheet_names()
        for i, n in enumerate(names):
            ws = self.ws(n)
            for sv in ws.iter("{%s}sheetView" % NS):
                if i == 0:
                    sv.set("tabSelected", "1")
                elif "tabSelected" in sv.attrib:
                    del sv.attrib["tabSelected"]
        bv = self.wb.find("m:bookViews/m:workbookView", N)
        bv.set("activeTab", "0")
        if "firstSheet" in bv.attrib:
            del bv.attrib["firstSheet"]

    def set_print_area(self, sheet, ref):
        idx = self.sheet_names().index(sheet)
        dn = self.wb.find("m:definedNames", N)
        if dn is None:
            dn = etree.Element("{%s}definedNames" % NS)
            self.wb.find("m:sheets", N).addnext(dn)
        for d in dn:
            if d.get("name") == "_xlnm.Print_Area" and d.get("localSheetId") == str(idx):
                d.text = f"{quote(sheet)}!{ref}"
                return
        d = etree.SubElement(dn, "{%s}definedName" % NS, name="_xlnm.Print_Area", localSheetId=str(idx))
        d.text = f"{quote(sheet)}!{ref}"

    def print_area(self, sheet):
        idx = str(self.sheet_names().index(sheet))
        dn = self.wb.find("m:definedNames", N)
        if dn is None:
            return None
        for d in dn:
            if d.get("name") == "_xlnm.Print_Area" and d.get("localSheetId") == idx:
                return d.text.split("!")[-1].replace("$", "")
        return None

    def last_value_row(self, sheet, max_col=None):
        ws = self.ws(sheet)
        last = 0
        for row in ws.find("m:sheetData", N).findall("m:row", N):
            for c in row.findall("m:c", N):
                col, r = split_ref(c.get("r"))
                if max_col and col2n(col) > max_col:
                    continue
                if c.find("m:v", N) is not None or c.find("m:is", N) is not None or c.find("m:f", N) is not None:
                    v = c.find("m:v", N)
                    if c.get("t") == "s" or (v is not None and (v.text or "").strip()) or c.find("m:is", N) is not None or c.find("m:f", N) is not None:
                        last = max(last, r)
        return last

    # ---------- 저장 ----------
    def save(self, path):
        for name, tr in self.trees.items():
            self.files[name] = etree.tostring(tr, xml_declaration=True, encoding="UTF-8", standalone=True)
        if "docProps/app.xml" in self.files:
            app = etree.fromstring(self.files["docProps/app.xml"])
            for e in list(app):
                if etree.QName(e).localname in ("HeadingPairs", "TitlesOfParts"):
                    app.remove(e)
            self.files["docProps/app.xml"] = etree.tostring(app, xml_declaration=True, encoding="UTF-8", standalone=True)
        if self.drop_calcchain and "xl/calcChain.xml" in self.files:
            del self.files["xl/calcChain.xml"]
            self.order.remove("xl/calcChain.xml")
            rels = self.files["xl/_rels/workbook.xml.rels"].decode("utf8")
            rels = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "", rels)
            self.files["xl/_rels/workbook.xml.rels"] = rels.encode("utf8")
            ct = self.files["[Content_Types].xml"].decode("utf8")
            ct = re.sub(r'<Override [^>]*calcChain[^>]*/>', "", ct)
            self.files["[Content_Types].xml"] = ct.encode("utf8")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                z.writestr(n, self.files[n])
