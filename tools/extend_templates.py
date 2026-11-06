"""양식 ver.1 → ver.2: 시험항목별 국문요약서/면제사유서/생략사유서 시트 추가.
- 2~34번: 국문요약서·면제사유서·생략사유서 3종 모두
- 35~48번: 면제사유서
새 시트는 같은 파일 안의 가장 비슷한 기존 시트를 복제해 서식을 그대로 쓰고, 제목·항목명·드롭다운만 바꾼다.
사용: python3 tools/extend_templates.py templates out_v2
"""
import os, re, sys, warnings
import openpyxl
sys.path.insert(0, os.path.dirname(__file__))
from xlsxedit import Book, quote

warnings.filterwarnings("ignore")
SRC, DST = sys.argv[1], sys.argv[2]
os.makedirs(DST, exist_ok=True)
LOG = []
log = LOG.append
LIST = "목록_선택값"
EX = "목록_면제조건"

SHORT = {
    2: "물용해도", 3: "녹는점어는점", 4: "끓는점", 5: "증기압", 6: "옥탄올물분배계수", 7: "밀도", 8: "입도분석",
    9: "인화성", 10: "폭발성", 11: "산화성", 12: "점도", 13: "해리상수",
    14: "급성경구", 15: "복귀돌연변이", 16: "피부자극성", 17: "피부과민성", 18: "급성경피", 19: "급성흡입",
    20: "눈자극", 21: "염색체이상(In vitro)", 22: "유전독성(In vivo)", 23: "반복투여28일", 24: "생식발달스크리닝",
    25: "추가유전독성", 26: "반복투여90일", 27: "최기형성", 28: "2세대생식독성", 29: "발암성",
    30: "어류급성", 31: "이분해성", 32: "물벼룩급성", 33: "담수조류", 34: "가수분해", 35: "본질적분해성",
    36: "분해산물확인", 37: "어류만성", 38: "물벼룩만성", 39: "육생식물급성", 40: "육생무척추동물급성",
    41: "활성슬러지호흡저해", 42: "흡착탈착", 43: "환경거동추가정보", 44: "육생식물만성", 45: "육생무척추동물만성",
    46: "흡착탈착추가정보", 47: "저서생물만성", 48: "생물농축성",
}
TITLE_FIX = {"반복투여독성_28일": "반복투여독성(28일)", "반복투여독성_90일": "반복투여독성(90일)"}
EXLIST_NAME = {
    2: "물용해도", 3: "녹는점/어는점", 4: "끓는점", 5: "증기압", 6: "옥탄올/물 분배계수", 7: "밀도", 8: "입도분석",
    9: "인화성", 10: "폭발성", 11: "산화성", 12: "점도", 13: "해리상수",
    14: "급성경구독성", 16: "피부 자극성/부식성", 17: "피부 과민성", 18: "급성경피독성", 19: "급성흡입독성",
    20: "눈 자극성/부식성", 21: "포유류 배양세포를 이용한 염색체이상(유전독성)", 22: "시험동물을 이용한 유전독성(유전독성)",
    23: "반복투여독성(28일)", 24: "생식 및 발달(발생)독성 스크리닝", 25: "추가 유전독성(생식세포 유전 독성 등)",
    26: "반복투여독성(90일)", 27: "최기형성", 28: "2세대 생식독성", 29: "발암성",
    30: "어류급성독성", 31: "이분해성", 32: "물벼룩급성독성", 33: "담수조류 성장저해", 34: "pH에 따른 가수분해",
}
RISK_SOIL = "위해성에 관한 자료의 노출시나리오에 따라 토양 노출이 무시할 만한 수준인 물질"
RISK_DEG = ("위해성에 관한 자료 작성 결과 화학물질의 분해성을 추가적으로 조사할 필요가 없다는 증거가 있는 물질"
            "(예, 환경 노출이 매우 낮아 위해성이 충분히 통제되고, 영 별표 1의2 제2호의 가목 또는 나목에 모두 해당하지 않는 물질)")
LOW_ADS = "물리화학적 성질에 근거하여, 흡착성 가능성이 낮은 물질(예, 비이온성 물질로 logKow <3)"
# 국립환경과학원고시 제2022-62호 [별표 2] 시험항목별 시험면제조건 중 기존 목록에 없던 항목
EXTRA_EX = {
    35: ("본질적분해성", ["무기물", "이분해성 물질", "모의 생분해성 시험자료(지표수, 토양, 수중 퇴적물)가 있는 물질",
                       "고분자화합물", RISK_DEG]),
    36: ("분해산물의 확인", ["이분해성 또는 본질적분해성 물질", "고분자화합물",
                         "위해성에 관한 자료 작성 결과 화학물질의 분해산물을 확인할 필요가 없다는 증거가 있는 물질"
                         "(예, 환경 노출이 매우 낮아 위해성이 충분히 통제되고, 영 별표 1의2 제2호의 가목 또는 나목에 모두 해당하지 않는 물질)"]),
    37: ("어류만성독성", ["위해성에 관한 자료 작성 결과 수생 생물에 대한 화학물질의 영향을 추가적으로 조사할 필요가 없다는 증거가 있는 물질(예, 수생환경에 대한 위해도가 1미만)"]),
    38: ("물벼룩만성독성", ["위해성에 관한 자료 작성 결과 수생 생물에 대한 화학물질의 영향을 추가적으로 조사할 필요가 없다는 증거가 있는 물질(예, 수생환경에 대한 위해도가 1미만)"]),
    39: ("육생식물 급성독성", [RISK_SOIL, "평형분배방법에 의한 토양생물 유해성평가가 가능한 물질"]),
    40: ("육생 무척추동물 급성독성", [RISK_SOIL, "평형분배방법에 의한 토양생물 유해성평가가 가능한 물질"]),
    41: ("활성슬러지 호흡저해", ["하수처리시설로 배출되지 않는 물질",
                           "이분해성물질로 시험에 사용된 농도가 하수처리 시설로의 예상 유입 농도의 범위 내인 물질",
                           "물에 극히 불용성 등 미생물 독성이 발생하지 않을 것 같은 완화요건이 있는 물질"]),
    42: ("흡착 및 탈착", [LOW_ADS, "물질 및 그 분해산물이 빠르게 분해되는 물질",
                       "신뢰성 있는 스크리닝 방법 등을 통해 흡착 및 탈착을 추정할 수 있는 물질", "고분자화합물"]),
    43: ("환경 거동 및 동태에 대한 추가정보", ["무기물", "이분해성 또는 본질적분해성 물질", "고분자화합물",
                                     "물에 극히 불용이고, 위해성에 관한 자료의 노출시나리오에 따라 토양과 침전물 노출이 무시할 만한 수준인 물질",
                                     "위해성에 관한 자료 작성 결과 화학물질의 환경거동 및 동태를 추가적으로 조사할 필요가 없다는 증거가 있는 물질"]),
    44: ("육생식물 만성독성", [RISK_SOIL, "위해성에 관한 자료 작성 결과 육상 생물에 대한 화학물질의 영향을 추가적으로 조사할 필요가 없다는 증거가 있는 물질(예, 토양환경에 대한 위해도가 1미만)"]),
    45: ("육생 무척추동물 만성독성", [RISK_SOIL, "위해성에 관한 자료 작성 결과 육상 생물에 대한 화학물질의 영향을 추가적으로 조사할 필요가 없다는 증거가 있는 물질(예, 토양환경에 대한 위해도가 1미만)"]),
    46: ("흡착 및 탈착에 대한 추가 정보", [LOW_ADS, "물질 및 그 분해산물이 빠르게 분해되는 물질",
                                   "흡착 및 탈착에 관한 측정 시험자료가 있는 물질", "고분자화합물"]),
    47: ("저서생물만성독성", ["위해성에 관한 자료 작성 결과 저서생물에 대한 화학물질의 영향을 조사할 필요가 없다는 증거가 있는 물질(예, 평형분배방법에 의한 저생생물 유해성평가결과 위해도가 1미만)"]),
    48: ("생물농축성", ["생물농축 가능성이 낮거나(예, logKow <3), 생체막을 통과할 수 없다고 예상되는 물질",
                     "이분해성 또는 본질적분해성 물질", "무기물", "고분자화합물",
                     "위해성에 관한 자료의 노출시나리오에 따라 수생환경에 대한 노출이 무시할 만한 수준인 물질"]),
}

OMIT_REASON = ("국가·국제기구 평가보고서를 통해 해당 시험항목의 유해성을 판단할 수 있으므로 화학물질의 등록 및 평가 등에 관한 "
               "법률 시행령 제13조 제6호의2에 따라 자료제출을 생략하고자 한다.")


def sheet_no(name):
    m = re.match(r"^(\d+)\.", name)
    return int(m.group(1)) if m else None


def kind_of(name):
    return "생략" if name.endswith("_생략") else "면제" if name.endswith("_면제") else "제출"


class Ctx:
    def __init__(self, key):
        self.key = key
        self.path = os.path.join(SRC, f"{key}.xlsx")
        self.ro = openpyxl.load_workbook(self.path)
        self.b = Book(self.path)
        self.titles = {}
        toc = self.ro["목차"]
        for r in range(3, 40):
            no, nm = toc.cell(r, 1).value, toc.cell(r, 2).value
            if isinstance(no, (int, float)) and nm:
                nm = str(nm).strip()
                self.titles[int(no)] = (TITLE_FIX.get(nm, nm), r)
        self.ex = {}
        ws = self.ro[EX]
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 2).value
            if v:
                self.ex.setdefault(str(v).strip(), []).append(r)
        self.next_helper_col = 32  # AF

    def has(self, name):
        return name in self.b.sheet_names()

    def name(self, no, kind):
        return f"{no}.{SHORT[no]}" + ("" if kind == "제출" else "_" + kind)

    def title(self, no):
        return self.titles[no][0]

    def ex_formula(self, no):
        nm = EXLIST_NAME.get(no)
        rows = self.ex.get(nm) if nm else None
        if not rows:
            return None
        return f"{quote(EX)}!$C${rows[0]}:$C${rows[-1]}"

    def block(self, col, row, skip_header=True):
        """목록 시트에서 row를 포함하는 연속 구간 → 수식"""
        ws = self.ro[LIST]
        c = openpyxl.utils.column_index_from_string(col)
        s = row
        while s > 1 and ws.cell(s - 1, c).value not in (None, ""):
            s -= 1
        e = row
        while ws.cell(e + 1, c).value not in (None, ""):
            e += 1
        if skip_header and str(ws.cell(s, c).value).startswith(("#", "<")):
            s += 1
        return f"{quote(LIST)}!${col}${s}:${col}${e}"

    def find_list(self, col, text):
        ws = self.ro[LIST]
        c = openpyxl.utils.column_index_from_string(col)
        for r in range(1, ws.max_row + 1):
            if str(ws.cell(r, c).value or "").strip() == text:
                return r
        raise KeyError(text)

    def helper(self, values):
        """목록 시트 빈 열에 새 선택지 목록을 쓰고 수식 반환"""
        col = openpyxl.utils.get_column_letter(self.next_helper_col)
        self.next_helper_col += 1
        for i, v in enumerate(values, start=1):
            self.b.set(LIST, f"{col}{i}", v)
        return f"{quote(LIST)}!${col}$1:${col}${len(values)}"

    def relabel(self, sheet, ref, text):
        """항목명 셀의 앞 들여쓰기는 유지하고 글자만 변경"""
        old = self.b.get(sheet, ref) or ""
        pre = old[: len(old) - len(old.lstrip())]
        self.b.set(sheet, ref, pre + text)


# ---------------- 공통: 면제·생략 시트 ----------------
def add_exempt(cx, no, base):
    new = cx.name(no, "면제")
    if cx.has(new):
        return
    cx.b.copy_sheet(base, new)
    b = cx.b
    b.set(new, "A2", cx.title(no))
    for ref in ("D6", "D7", "D14", "D15", "D19", "D20", "D24"):
        b.set(new, ref, "")
    b.set(new, "D8", "면제")
    f = cx.ex_formula(no)
    if f:
        b.add_list_dv(new, "D14", f)
    else:
        b.remove_dv(new, "D14")
    log(f"  + {new}" + ("" if f else "  (면제조건 목록 없음 → 직접 입력)"))


def add_omit_F(cx, no, base, skeleton, relabels=None, title_cells=("A2",), item_cell=None):
    new = cx.name(no, "생략")
    if cx.has(new):
        return
    cx.b.copy_sheet(base, new)
    b = cx.b
    for c in title_cells:
        b.set(new, c, cx.title(no))
    if item_cell:
        b.set(new, item_cell, cx.title(no))
    for c, v in (relabels or {}).items():
        b.set(new, c, v)
    return new


def omit_human(cx, no, skeleton, relabels=None):
    new = add_omit_F(cx, no, "14.급성경구_생략", skeleton, relabels, title_cells=("A2", "F8"))
    if not new:
        return
    b = cx.b
    b.set(new, "F9", "주요자료"); b.set(new, "F10", "국가·국제기구 평가보고서")
    b.set(new, "F15", OMIT_REASON); b.set(new, "F16", ""); b.set(new, "A24", "출처 1: ")
    b.set(new, "F27", skeleton)
    b.remove_dv(new, "F8")   # 시험항목명은 직접 표기(기존 목록에 없는 항목 포함)
    log(f"  + {new}")


def omit_env(cx, no, skeleton, relabels=None):
    new = add_omit_F(cx, no, "31.이분해성_생략", skeleton, relabels, title_cells=("A2", "F8"))
    if not new:
        return
    cx.b.set(new, "F24", skeleton)
    log(f"  + {new}")


def omit_phys(cx, no, six2):
    new = cx.name(no, "생략")
    if cx.has(new):
        return
    b = cx.b
    b.copy_sheet("2.물용해도_면제", new)
    b.set(new, "A2", cx.title(no)); b.set(new, "F8", cx.title(no))
    b.set_row_height(new, 13, 62); b.set_row_height(new, 14, 62)
    b.set(new, "F9", "주요자료"); b.add_list_dv(new, "F9", '"주요자료, 참고자료"')
    b.set(new, "A11", "2. 등록제출자료 생략의 사유")
    b.set(new, "A13", "생략사유"); b.set(new, "F13", OMIT_REASON)
    b.set(new, "A14", "시험생략조건(영제13조)"); b.set(new, "F14", six2)
    b.add_list_dv(new, "F14", f"{quote(LIST)}!$D$1:$D$15")
    b.set(new, "A15", "자료출처")
    for ref in ("F15", "F19", "F20"):
        b.set(new, ref, "")
    b.set(new, "A24", "출처 1: "); b.set(new, "A33", "출처 2: ")
    b.set(new, "F27", "- 결과값 : \n- 시험조건 : ")
    log(f"  + {new}")


def submit_copy(cx, no, base, guideline_formula, relabels=(), dvs=(), cond_cell=None, clear=()):
    new = cx.name(no, "제출")
    if cx.has(new):
        return None
    b = cx.b
    b.copy_sheet(base, new)
    b.set(new, "A2", cx.title(no))
    vc = "D" if cx.key == "phys" else "C"
    b.add_list_dv(new, f"{vc}27", guideline_formula)
    for ref, text in relabels:
        cx.relabel(new, ref, text)
    for ref, f in dvs:
        if f is None:
            b.remove_dv(new, ref)
        else:
            b.add_list_dv(new, ref, f)
    for ref in clear:
        b.set(new, ref, "")
    if cond_cell:
        f = cx.ex_formula(no)
        b.set(new, cond_cell, "해당없음")
        if f:
            b.add_list_dv(new, cond_cell, f)
        else:
            b.remove_dv(new, cond_cell)
    log(f"  + {new}  (기준: {base})")
    return new


def finish(cx, desc_note=None):
    b = cx.b
    names = b.sheet_names()
    items = [n for n in names if sheet_no(n)]
    kord = {"제출": 0, "생략": 1, "면제": 2}
    items.sort(key=lambda n: (sheet_no(n), kord[kind_of(n)]))
    lists = [n for n in names if n.startswith("목록_")]
    order = ["목차"] + items + lists
    b.reorder(order)
    for n in lists:
        b.hide(n)
    b.select_first()
    # 목차 시트명 갱신
    for no, (_, r) in cx.titles.items():
        sh = [n for n in items if sheet_no(n) == no]
        b.set("목차", f"D{r}", ", ".join(sh) if sh else "(양식 없음)")
        if desc_note and no in desc_note:
            b.set("목차", f"E{r}", desc_note[no])
    out = os.path.join(DST, f"{cx.key}.xlsx")
    b.save(out)
    return out


# ======================= 물화 =======================
def do_phys():
    cx = Ctx("phys")
    log("■ 물화")
    six2 = openpyxl.load_workbook(os.path.join(SRC, "env.xlsx"))["31.이분해성_생략"]["F11"].value
    # 새 시험방법 목록 (목록_선택값 빈 열에 추가)
    g_visc = cx.helper(["OECD Guideline 114 (Viscosity of Liquids)", "other guideline:"])
    g_pka = cx.helper(["OECD Guideline 112 (Dissociation Constants in Water)", "other guideline:"])
    g_expl = cx.helper(["EU Method A.14 (Explosive properties)", "UN Manual of Tests and Criteria (Test Series 1~3)",
                        "ASTM E537 (DSC 열안정성)", "other guideline:"])
    g_oxid = cx.helper(["EU Method A.17 (Oxidising Properties (Solids))", "EU Method A.21 (Oxidising Properties (Liquids))",
                        "UN Test O.1 (Test for oxidizing solids)", "UN Test O.2 (Test for oxidizing liquids)",
                        "UN Test O.3 (Gravimetric test for oxidizing solids)", "other guideline:"])
    g_psd = cx.block("D", 76)

    # 제출(국문요약서)
    submit_copy(cx, 8, "7.밀도", g_psd,
                relabels=[("A28", "측정방법"), ("A42", "평가항목"), ("A45", "측정 조건"), ("A46", "분포/비율(%)")],
                dvs=[("D28", '"sieve analysis, laser diffraction, sedimentation, microscopy, cascade impactor, other:"'),
                     ("D42", '"D10, D50 (median), D90, mean diameter, MMAD, particle size distribution, other:"'),
                     ("D44", '"µm, nm, mm, %, other:"'), ("D46", None)],
                clear=["D37"], cond_cell="D60")
    for no, g in ((10, g_expl), (11, g_oxid)):
        lab = {
            10: ["열 민감도 시험 결과", "비고", "기계적 민감도(충격) 시험 결과", "비고", "기계적 민감도(마찰) 시험 결과",
                 "비고", "폭발성 판정", "비고", "발열 분해 에너지(J/g)", "추가정보"],
            11: ["연소시간/연소속도(시험물질 혼합물)", "비고", "기준물질 대비 결과", "비고", "압력상승 시간(액체)",
                 "비고", "산화성 판정", "비고", "구분(GHS)", "추가정보"],
        }[no]
        submit_copy(cx, no, "9.인화성", g, relabels=[(f"A{44 + i}", t) for i, t in enumerate(lab)], cond_cell="D62")
    submit_copy(cx, 12, "7.밀도", g_visc,
                relabels=[("A28", "측정방법")],
                dvs=[("D28", '"capillary viscometer, rotational viscometer, falling ball viscometer, other:"'),
                     ("D42", '"dynamic viscosity, kinematic viscosity, other:"'),
                     ("D44", '"mPa·s, Pa·s, mm²/s, cSt, other:"')],
                clear=["D37"], cond_cell="D60")
    submit_copy(cx, 13, "7.밀도", g_pka,
                relabels=[("A28", "측정방법"), ("A44", "해리 작용기/형태")],
                dvs=[("D28", '"titration method, spectrophotometric method, conductometric method, calculation, other:"'),
                     ("D42", '"pKa, pKa1, pKa2, pKa3, pKb, other:"'), ("D44", None)],
                clear=["D37"], cond_cell="D60")
    # 면제·생략
    for no in range(2, 14):
        add_exempt(cx, no, "8.입도분석_면제")
        omit_phys(cx, no, six2)
    return finish(cx)


# ======================= 인체 =======================
def do_human():
    cx = Ctx("human")
    log("■ 인체")
    g_rep = cx.helper(sorted(set(
        [cx.ro[LIST].cell(r, c).value for c in (20, 22, 24) for r in range(52, 90)
         if cx.ro[LIST].cell(r, c).value and not str(cx.ro[LIST].cell(r, c).value).startswith(("#", "other"))]),
        key=lambda s: (not s.startswith("OECD"), s)) + ["other guideline:"])
    g_inh = cx.block("H", 53)
    g_vivo = cx.block("R", 53)
    g_repro = cx.block("Z", 53)
    g_carc = cx.block("AD", 53)
    r_gen = cx.find_list("A", "#유전독성 여부")
    r_cyt = cx.find_list("A", "#세포독성")
    l_gen = cx.block("A", r_gen + 1)
    l_cyt = cx.block("A", r_cyt + 1)
    l_route_inh = cx.block("F", cx.find_list("F", "inhalation: gas"))
    l_lc = cx.block("A", cx.find_list("A", "LC50"))
    l_unit_inh = cx.block("E", cx.find_list("E", "mg/L air"))

    # 19 급성흡입 (기준: 14.급성경구)
    submit_copy(cx, 19, "14.급성경구", g_inh,
                relabels=[("A46", "노출농도"), ("A47", "농도별 성별 개체 수")],
                dvs=[("C44", l_route_inh), ("C53", l_lc), ("C55", l_unit_inh), ("C28", None)],
                cond_cell="C69")
    # 반복투여·생식·발암 (기준: 27.최기형성)
    animal = {
        23: (g_rep, ["회복군 여부", "산모 검사", "임상관찰·체중·혈액·요검사", "부검·조직병리검사"]),
        26: (g_rep, ["회복군 여부", "산모 검사", "임상관찰·체중·혈액·요검사", "부검·조직병리검사"]),
        24: (g_repro, ["교배절차 세부사항", "", "모체(F0) 검사", "새끼(F1) 검사"]),
        28: (g_repro, ["교배절차 세부사항", "", "부모세대(P) 검사", "자손세대(F1/F2) 검사"]),
        29: (g_carc, ["투여기간·중간부검 여부", "", "임상관찰·체중·생존율", "종양 병리검사"]),
    }
    for no, (g, (l45, _, l51, l52)) in animal.items():
        submit_copy(cx, no, "27.최기형성", g,
                    relabels=[("A45", l45), ("A51", l51), ("A52", l52)], cond_cell="C73")
    # In vivo 유전독성·추가 유전독성 (기준: 27.최기형성)
    for no in (22, 25):
        submit_copy(cx, no, "27.최기형성", g_vivo,
                    relabels=[("A45", "채취 시점(sampling time)"), ("A51", "검사 세포·조직"), ("A52", "양성대조물질"),
                              ("A58", "유전독성 여부"), ("A59", "결과 상세"), ("A60", "표적세포 독성(세포독성)")],
                    dvs=[("C58", l_gen), ("C60", l_cyt)], cond_cell="C73")

    sk_acute = lambda e: f"- 종말점 및 결과값:\n • {e} = \n • 사망개체: \n • 주요영향: "
    sk_rep = "- 종말점 및 결과값:\n • NOAEL = \n • LOAEL = \n • 주요영향: "
    sk_gen = "- 결과:\n • 유전독성 여부: \n • 대사활성계 유무: \n • 세포독성: "
    invitro = {"C31": "대사활성계", "C34": "시험균주/세포", "C35": "시험농도"}
    omit = {
        15: (sk_gen, invitro), 16: ("- 결과:\n • 피부 자극성 판정: \n • 평가 점수: \n • 주요영향: ", None),
        17: ("- 결과:\n • 피부 과민성 판정: \n • 평가 결과(SI/EC3 등): \n • 주요영향: ", None),
        21: (sk_gen, invitro), 22: (sk_gen, None), 23: (sk_rep, None), 24: (sk_rep, None), 25: (sk_gen, None),
        27: (sk_rep, None), 28: (sk_rep, None),
        29: ("- 종말점 및 결과값:\n • NOAEL = \n • 종양 발생: \n • 주요영향: ", None),
    }
    for no, (sk, rl) in omit.items():
        omit_human(cx, no, sk, rl)
    for no in range(14, 30):
        add_exempt(cx, no, "23.반복투여28일_면제")
    return finish(cx)


# ======================= 환경 =======================
def do_env():
    cx = Ctx("env")
    log("■ 환경")
    eco = {"C27": "시험종", "C28": "노출기간/시험농도"}
    omit_env(cx, 30, "- 종말점 및 결과값:\n • LC50(96 h) = \n • 주요영향: ", eco)
    omit_env(cx, 32, "- 종말점 및 결과값:\n • EC50(48 h) = \n • 주요영향: ", eco)
    omit_env(cx, 33, "- 종말점 및 결과값:\n • ErC50(72 h) = \n • NOErC = \n • 주요영향: ", eco)
    ws = cx.ro[EX]
    r = max(rr for rows in cx.ex.values() for rr in rows) + 2
    for no, (nm, conds) in EXTRA_EX.items():
        start = r
        for c in conds:
            cx.b.set(EX, f"B{r}", nm); cx.b.set(EX, f"C{r}", c); r += 1
        cx.ex[nm] = list(range(start, r))
        EXLIST_NAME[no] = nm
        r += 1
    log("  - 목록_면제조건에 35~48번 면제조건 추가(국립환경과학원고시 제2022-62호 별표 2)")
    for no in list(range(30, 34)) + list(range(35, 49)):
        add_exempt(cx, no, "34.가수분해_면제")
    note = {no: "면제사유서만 제공" for no in range(35, 49)}
    return finish(cx, note)


outs = [do_phys(), do_human(), do_env()]
open(os.path.join(DST, "변경내역_v2.txt"), "w", encoding="utf8").write("\n".join(LOG))
print("\n".join(LOG))
