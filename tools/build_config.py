"""양식 엑셀에서 입력칸 매핑(config/templates.json)을 생성한다.
양식 파일은 읽기만 하며 수정하지 않는다.
사용: python3 tools/build_config.py  (regdoc 폴더에서 실행)
"""
import json, re, zipfile, warnings, os
import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter

warnings.filterwarnings("ignore")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
T = os.path.join(ROOT, "templates")

FILES = {
    "phys": "phys.xlsx",    # 시험요약서00-00-0_물화2026_ver.1.xlsx
    "human": "human.xlsx",  # 시험요약서00-00-0_인체2026_ver.1.xlsx
    "env": "env.xlsx",      # 시험요약서00-00-0_환경2026_ver.1.xlsx
}
ORIG = os.path.join(ROOT, "orig")  # ver.0 원본(작성예시 추출용, 사이트에는 포함하지 않음)

# (분야, 시험항목, 유형, 파일, 시트, 관련 키워드, ver.0 원본 시트명)
ITEMS = [
    ("물리화학", "물질의 상태", "제출", "phys", "1.물질의상태", "physical state appearance colour odour 성상 색", "1.물질의상태"),
    ("물리화학", "물용해도", "제출", "phys", "2.물용해도", "water solubility 용해도 mg/L", "2.물용해도"),
    ("물리화학", "물용해도", "면제", "phys", "2.물용해도_면제", "water solubility hydrolysis half-life 가수분해 반감기", "면제_물용해도"),
    ("물리화학", "녹는점/어는점", "제출", "phys", "3.녹는점어는점", "melting point freezing point 녹는점 어는점", "3.녹는점어는점"),
    ("물리화학", "녹는점/어는점", "면제", "phys", "3.녹는점어는점_면제", "melting point freezing point pour point 녹는점", "면제_녹는점"),
    ("물리화학", "끓는점", "제출", "phys", "4.끓는점", "boiling point 끓는점 decomposition", "4.끓는점"),
    ("물리화학", "증기압", "제출", "phys", "5.증기압", "vapour pressure vapor pressure 증기압 Pa hPa", "5.증기압"),
    ("물리화학", "옥탄올/물 분배계수", "제출", "phys", "6.옥탄올물분배계수", "partition coefficient log Kow log Pow octanol 분배계수", "6.옥탄올"),
    ("물리화학", "밀도", "제출", "phys", "7.밀도", "density relative density 밀도 g/cm3", "7.밀도"),
    ("물리화학", "입도분석", "면제", "phys", "8.입도분석_면제", "particle size granulometry 입도 physical state", "8.입도분석(면제)"),
    ("물리화학", "인화성", "제출", "phys", "9.인화성", "flash point flammability 인화점 인화성", "9.인화성 (인화점)"),
    ("물리화학", "폭발성", "면제", "phys", "10.폭발성_면제", "explosive explosivity oxygen balance 폭발성 structure", "10.폭발성(면제)"),
    ("물리화학", "산화성", "면제", "phys", "11.산화성_면제", "oxidising oxidizing 산화성", "11.산화성(면제)"),
    ("인체", "급성경구독성", "제출", "human", "14.급성경구", "acute oral toxicity LD50 oral rat 급성경구", "14.급성경구"),
    ("인체", "급성경구독성", "생략", "human", "14.급성경구_생략", "acute oral toxicity LD50 oral rat 급성경구", "14.경구(생략)"),
    ("인체", "복귀돌연변이", "제출", "human", "15.복귀돌연변이", "Ames bacterial reverse mutation Salmonella TA98 TA100 복귀돌연변이", "15.복귀돌연변이"),
    ("인체", "피부 자극성/부식성", "제출", "human", "16.피부자극성", "skin irritation corrosion dermal erythema oedema 피부자극", "16.피부자극성"),
    ("인체", "피부 과민성", "제출", "human", "17.피부과민성", "skin sensitisation sensitization LLNA guinea pig maximisation 과민성", "17.과민성"),
    ("인체", "급성경피독성", "제출", "human", "18.급성경피", "acute dermal toxicity LD50 dermal 급성경피", "18. 급성경피"),
    ("인체", "급성경피독성", "생략", "human", "18.급성경피_생략", "acute dermal toxicity LD50 dermal 급성경피", "생략_급성경피"),
    ("인체", "급성흡입독성", "생략", "human", "19.급성흡입_생략", "acute inhalation toxicity LC50 inhalation 급성흡입", "생략_급성흡입"),
    ("인체", "눈 자극성/부식성", "제출", "human", "20.눈자극", "eye irritation cornea iris conjunctiva 눈자극", "20.눈자극"),
    ("인체", "눈 자극성/부식성", "생략", "human", "20.눈자극_생략", "eye irritation cornea iris conjunctiva 눈자극", "20.눈자극(생략)"),
    ("인체", "염색체이상(In vitro)", "제출", "human", "21.염색체이상(In vitro)", "chromosome aberration in vitro micronucleus CHL CHO 염색체이상", "21.In vitro"),
    ("인체", "유전독성(In vivo)", "면제", "human", "22.유전독성(In vivo)_면제", "in vivo micronucleus genotoxicity 유전독성", "22. In vivo_면제"),
    ("인체", "반복투여독성(28일)", "면제", "human", "23.반복투여28일_면제", "repeated dose 28-day 90-day subchronic NOAEL 반복투여", "23.반복(28일)_면제"),
    ("인체", "생식·발달독성 스크리닝", "면제", "human", "24.생식발달스크리닝_면제", "reproductive developmental screening OECD 421 422 생식 발달", "24.스크리닝_면제"),
    ("인체", "반복투여독성(90일)", "생략", "human", "26.반복투여90일_생략", "repeated dose 90-day subchronic NOAEL 반복투여", "26.90일 반복(생략)"),
    ("인체", "최기형성", "제출", "human", "27.최기형성", "prenatal developmental toxicity teratogenicity OECD 414 최기형성", "27.최기형성"),
    ("환경", "어류급성독성", "제출", "env", "30.어류급성", "fish acute toxicity LC50 96 h 어류", "30.어류급성"),
    ("환경", "이분해성", "제출", "env", "31.이분해성", "ready biodegradability OECD 301 BOD ThOD 분해도 이분해성", "31.이분해성"),
    ("환경", "이분해성", "생략", "env", "31.이분해성_생략", "ready biodegradability OECD 301 BOD ThOD 분해도 이분해성", "생략_이분해성"),
    ("환경", "물벼룩급성독성", "제출", "env", "32.물벼룩급성", "Daphnia acute immobilisation EC50 48 h 물벼룩", "32.물벼룩급성"),
    ("환경", "담수조류 성장저해", "제출", "env", "33.담수조류", "algae growth inhibition ErC50 72 h Pseudokirchneriella 조류", "33. 담수조류"),
    ("환경", "가수분해", "제출", "env", "34.가수분해", "hydrolysis half-life pH 4 7 9 가수분해", "34.가수분해(제출)"),
    ("환경", "가수분해", "생략", "env", "34.가수분해_생략", "hydrolysis half-life pH 가수분해", "생략_가수분해"),
    ("환경", "가수분해", "면제", "env", "34.가수분해_면제", "hydrolysis half-life pH 가수분해", "면제_가수분해"),
]

SKIP_TOKENS = {":", "-", "：", "·"}


def sheet_xml_paths(path):
    z = zipfile.ZipFile(path)
    wb = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rmap = {}
    for m in re.finditer(r"<Relationship [^>]*>", rels):
        tag = m.group(0)
        i = re.search(r'Id="([^"]+)"', tag).group(1)
        t = re.search(r'Target="([^"]+)"', tag).group(1)
        rmap[i] = "xl/" + t.lstrip("/").replace("xl/", "", 1)
    out = {}
    for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        name = m.group(1).replace("&amp;", "&")
        out[name] = rmap[m.group(2)]
    return z, out


def in_sqref(cell, sqref):
    col, row = openpyxl.utils.cell.coordinate_from_string(cell)
    c = openpyxl.utils.column_index_from_string(col)
    for part in sqref.split():
        try:
            a, b, cc, d = range_boundaries(part)
        except Exception:
            continue
        if b is None or b > 300:
            continue
        if a <= c <= cc and b <= row <= d:
            return True
    return False


def list_options(formula, wb):
    formula = formula.strip()
    if formula.startswith('"'):
        return [x.strip() for x in formula.strip('"').split(",") if x.strip()]
    m = re.match(r"'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", formula)
    if m and m.group(1) in wb.sheetnames:
        ws = wb[m.group(1)]
        vals = []
        for row in ws[f"{m.group(2)}{m.group(3)}:{m.group(4)}{m.group(5)}"]:
            for c in row:
                if c.value not in (None, ""):
                    vals.append(str(c.value).strip())
        return vals
    return None


def get_validations(xml):
    dvs = []
    for m in re.finditer(r'<dataValidation\b([^>]*)>(.*?)</dataValidation>', xml, re.S):
        attrs, body = m.group(1), m.group(2)
        if 'type="list"' not in attrs:
            continue
        sq = re.search(r'sqref="([^"]+)"', attrs)
        f = re.search(r"<formula1>(.*?)</formula1>", body, re.S)
        if sq and f:
            dvs.append((sq.group(1), f.group(1)))
    for m in re.finditer(r'<x14:dataValidation\b([^>]*)>(.*?)</x14:dataValidation>', xml, re.S):
        attrs, body = m.group(1), m.group(2)
        if 'type="list"' not in attrs:
            continue
        f = re.search(r"<xm:f>(.*?)</xm:f>", body, re.S)
        sq = re.search(r"<xm:sqref>(.*?)</xm:sqref>", body, re.S)
        if sq and f:
            dvs.append((sq.group(1), f.group(1)))
    import html
    return [(s, html.unescape(f)) for s, f in dvs]


def print_last_col(path, sheet):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    pa = ws.print_area
    if pa:
        pa = pa if isinstance(pa, str) else pa[0]
        pa = pa.split("!")[-1].replace("$", "")
        a, b, c, d = range_boundaries(pa)
        return c, d
    return ws.max_column, ws.max_row


def build_sheet(path, sheet, wb, xml):
    ws = wb[sheet]
    last_col, last_row = print_last_col(path, sheet)
    last_row = max(last_row, ws.max_row if ws.max_row < 200 else last_row)
    merged = {}
    for r in ws.merged_cells.ranges:
        if r.min_row > 200:
            continue
        for row in range(r.min_row, r.max_row + 1):
            for col in range(r.min_col, r.max_col + 1):
                merged[(row, col)] = (r.min_row, r.min_col, r.max_row, r.max_col)
    dvs = get_validations(xml)

    def block(row, col):
        return merged.get((row, col), (row, col, row, col))

    def val(row, col):
        v = ws.cell(row, col).value
        return None if v is None else str(v)

    fields = []
    section = ""
    sub = ""
    used = set()
    body_cands, body_done = {}, set()
    for row in range(3, min(last_row, 199) + 1):
        a = block(row, 1)
        if a[0] != row:  # 위 행의 병합에 포함된 경우(요약표 세로병합 등)
            # 요약표의 C열 라벨 행 처리는 아래 공통 로직에서
            pass
        cells = []  # (top-left col, maxcol, text) for blocks starting in this row
        col = 1
        while col <= last_col:
            b = block(row, col)
            if b[0] == row and b[1] == col:
                cells.append((col, b[3], val(row, col), b))
            col = b[3] + 1 if b[1] == col else col + 1
        texts = [(c, t) for c, _, t, _ in cells if t and t.strip()]
        if not texts and not cells:
            continue
        first = cells[0] if cells else None
        # 섹션 제목: 첫 블록이 인쇄폭 전체를 덮는 경우
        if first and first[1] >= last_col and first[0] == 1:
            t = (first[2] or "").strip()
            if not t and "결론" in section and section not in body_done:
                h = ws.row_dimensions[row].height or 15
                body_cands.setdefault(section, []).append((h, row))
                continue
            if t.startswith("출처"):
                fields.append({"cell": f"A{row}", "label": "출처 표기(요약표 상단, '출처 n: 저자(연도)., 제목' 형식)",
                               "section": section, "sub": sub, "options": None, "default": t, "tall": False})
                used.add(f"A{row}")
                continue
            if t:
                if re.match(r"^\d+\.", t):
                    section, sub = t, ""
                elif t.startswith("•") or t.startswith("["):
                    sub = t
            continue
        # 라벨/값 결정
        label_idx = None
        for i, (c, mc, t, b) in enumerate(cells):
            if t and t.strip() and t.strip() not in SKIP_TOKENS and not re.fullmatch(r"\d+", t.strip()):
                label_idx = i
                break
        if label_idx is None:
            continue
        lc, lmc, lt, lb = cells[label_idx]
        t0 = lt.strip()
        if t0.startswith("="):
            continue
        if re.match(r"^\d+\.\s", t0) or t0.startswith(("[", "<", "No.")):
            if re.match(r"^\d+\.\s", t0):
                section, sub = t0, ""
            elif t0.startswith("["):
                sub = t0
            continue
        if t0.startswith("•"):
            sub = t0
            continue
        if lt.strip().startswith("•"):
            sub = lt.strip()
        rest = [x for x in cells[label_idx + 1:] if not (x[2] and x[2].strip() in SKIP_TOKENS)]
        if not rest:
            # 라벨만 있고 값칸 없음 → 소제목
            t = lt.strip()
            if t.startswith("•") or t.startswith("[") or re.match(r"^\d+\.", t):
                if re.match(r"^\d+\.", t):
                    section, sub = t, ""
                else:
                    sub = t
            continue
        is_table = lc == 3 and any(x[2] and x[2].strip() == ":" for x in cells)
        if not is_table:
            rest = rest[-1:]
        # 값칸: 라벨 다음 첫 블록. 요약표(라벨 C, ':' D, 값 E, 결과 F 세로병합)도 처리
        for j, (vc, vmc, vt, vb) in enumerate(rest):
            addr = f"{get_column_letter(vc)}{row}"
            if addr in used:
                continue
            label = lt.strip()
            if j > 0:
                # 두 번째 값칸(예: 요약표 F열 시험결과)
                hdr = None
                for rr in range(row - 1, max(row - 4, 1), -1):
                    h = ws.cell(rr, vc).value
                    if h:
                        hdr = str(h).strip(); break
                label = hdr or (label + " (추가)")
            opts = None
            for sq, f in dvs:
                if in_sqref(addr, sq):
                    o = list_options(f, wb)
                    if o:
                        opts = o
                        break
            default = vt if vt not in (None, "") else ""
            if isinstance(default, str) and default.startswith("="):
                continue  # 수식칸은 건드리지 않음
            fields.append({
                "cell": addr,
                "label": label,
                "section": section,
                "sub": sub,
                "options": opts,
                "default": default.strip() if isinstance(default, str) else default,
                "tall": (vb[2] - vb[0]) >= 2,
            })
            used.add(addr)
            if j >= 1:
                break
    for sec, cands in body_cands.items():
        if any(f["section"] == sec for f in fields if f["label"] in ("결론", "요약", "결론 ")):
            continue
        h, r = max(cands)
        fields.append({"cell": f"A{r}", "label": "결론 및 요약 본문(서술형 문단)", "section": sec, "sub": "",
                       "options": None, "default": "", "tall": True})
    fields.sort(key=lambda f: int(re.sub(r"[A-Z]", "", f["cell"])))
    return {"title": str(ws["A2"].value or ""), "fields": fields}


def main():
    out = {"files": {k: "templates/" + v for k, v in FILES.items()}, "items": [], "sheets": {}}
    cache = {}
    origwb = {}
    for grp, name, typ, fk, sheet, kw, osheet in ITEMS:
        path = os.path.join(T, FILES[fk])
        if fk not in cache:
            z, sp = sheet_xml_paths(path)
            cache[fk] = (openpyxl.load_workbook(path), z, sp)
        wb, z, sp = cache[fk]
        if sheet not in wb.sheetnames:
            raise SystemExit(f"시트 없음: {fk} {sheet}")
        xml = z.read(sp[sheet]).decode("utf8")
        key = f"{fk}|{sheet}"
        info = build_sheet(path, sheet, wb, xml)
        op = os.path.join(ORIG, FILES[fk])
        if os.path.exists(op):
            if fk not in origwb:
                origwb[fk] = openpyxl.load_workbook(op)
            ows = origwb[fk][osheet]
            for f in info["fields"]:
                ov = ows[f["cell"]].value
                if isinstance(ov, str) and ov.strip() and not ov.startswith("=") and ov.strip() != str(f["default"]).strip():
                    f["example"] = ov.strip()
        out["sheets"][key] = info
        out["items"].append({"group": grp, "name": name, "type": typ, "file": fk, "sheet": sheet,
                             "key": key, "keywords": kw.split()})
    os.makedirs(os.path.join(ROOT, "config"), exist_ok=True)
    with open(os.path.join(ROOT, "config", "templates.json"), "w", encoding="utf8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    for k, v in out["sheets"].items():
        print(k, len(v["fields"]))


if __name__ == "__main__":
    main()
