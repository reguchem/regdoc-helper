"""원본 양식(ver.0) → 정리본(ver.1) 생성.
서식(셀 스타일·병합·열너비·이미지)은 그대로 두고 값/시트명/순서/드롭다운/인쇄영역만 수정한다.
"""
import os, sys, re, warnings, zipfile
import openpyxl
sys.path.insert(0, os.path.dirname(__file__))
from xlsxedit import Book, quote
import build_config as bc

warnings.filterwarnings("ignore")
SRC = sys.argv[1]   # 원본 폴더 (phys.xlsx, human.xlsx, env.xlsx)
DST = sys.argv[2]   # 출력 폴더
os.makedirs(DST, exist_ok=True)
LOG = []


def log(msg):
    LOG.append(msg)


def find_label_cell(wb, sheet, label, value_cols="DF"):
    """라벨 텍스트가 있는 행의 값 셀 주소."""
    ws = wb[sheet]
    for row in ws.iter_rows(max_row=120):
        for c in row:
            if isinstance(c.value, str) and c.value.strip() == label:
                for vc in value_cols:
                    return f"{vc}{c.row}"
    return None


def clear_examples(book, wb, path, sheet, from_section=3, extra=()):
    """3번 절 이후의 입력칸·출처표기·G열 메모를 비운다(물질별 예시 제거)."""
    z, sp = bc.sheet_xml_paths(path)
    xml = z.read(sp[sheet]).decode("utf8")
    info = bc.build_sheet(path, sheet, wb, xml)
    n = 0
    for f in info["fields"]:
        m = re.match(r"^(\d+)\.", f["section"] or "")
        sec = int(m.group(1)) if m else 0
        if sec >= from_section:
            if f["label"].startswith("출처 표기"):
                k = re.match(r"(출처\s*\d*\s*:)", str(f["default"]))
                book.set(sheet, f["cell"], (k.group(1) + " ") if k else "출처: ")
            else:
                book.set(sheet, f["cell"], "")
            n += 1
    ws = wb[sheet]
    cur = 0
    for r in range(1, 121):
        v = ws.cell(r, 1).value
        if not isinstance(v, str) or not v.strip():
            continue
        t = v.strip()
        m = re.match(r"^(\d+)\.\s", t)
        if m:
            cur = int(m.group(1)); continue
        if cur >= from_section and len(t) > 40 and not t.startswith("출처"):
            book.set(sheet, f"A{r}", ""); n += 1
    for row in ws.iter_rows(min_row=1, max_row=120, min_col=7, max_col=13):
        for c in row:
            if c.value not in (None, ""):
                book.set(sheet, c.coordinate, "")
                n += 1
    for ref in extra:
        book.set(sheet, ref, "")
        n += 1
    log(f"  - [{sheet}] 물질별 작성예시 {n}칸 비움")


def common(book):
    c = book.replace_text("Data waving", "Data waiving")
    if c: log(f"  - 'Data waving' → 'Data waiving' ({c}곳)")
    c = book.replace_text("바업", "방법")
    if c: log(f"  - 시행령 2의2호 '바업' → '방법' ({c}곳)")
    c = book.replace_text(r"한함\s*○\s*분자 내에 폭발성과 관련 있는 화학그룹이 없는 물질", "한함", regex=True)
    if c: log(f"  - 면제목록 폭발성 800J/g 항목 끝에 붙은 중복 문장 삭제 ({c}곳)")
    book.replace_in_formulas("4(신뢰성이 부여될 수 없으)", "4(신뢰성이 부여될 수 없음)")
    log("  - 신뢰도 드롭다운 '4(신뢰성이 부여될 수 없으)' → '…없음)'")
    removed = 0
    for s in book.sheet_names():
        removed += book.prune_dv(s)
    log(f"  - 시트 전체로 복제돼 있던 드롭다운 잔여 범위 {removed}개 정리")


def finish(book, renames, order, helpers, description_rows, header):
    # 이름 변경
    for old, new in renames.items():
        book.rename(old, new)
    book.reorder(order)
    for h in helpers:
        book.hide(h)
    book.select_first()
    # 목차(Description)
    book.set("목차", "B2", header)
    for r, val in description_rows.items():
        book.set("목차", f"D{r}", val.replace("\n", ", "))
    # 인쇄영역: 내용 끝 행까지 확장
    for s in order:
        pa = book.print_area(s)
        if not pa or not ("_생략" in s or "_면제" in s):
            continue
        a, b = pa.split(":")
        lastcol = re.match(r"[A-Z]+", b).group(0)
        from xlsxedit import col2n
        lr = book.last_value_row(s, max_col=col2n(lastcol))
        oldr = int(re.sub(r"[A-Z]", "", b))
        if lr > oldr:
            book.set_print_area(s, f"$A$1:${lastcol}${lr + 1}")
            log(f"  - [{s}] 인쇄영역 {oldr}행 → {lr + 1}행으로 확장")


# ======================= 물화 =======================
def do_phys():
    path = os.path.join(SRC, "phys.xlsx")
    wb = openpyxl.load_workbook(path)
    b = Book(path)
    log("■ 물화 (시험요약서00-00-0_물화2026_ver.1.xlsx)")
    common(b)
    b.replace_text("물질명는", "AAA는"); b.replace_text("물질명은", "AAA은")
    log("  - 면제_물용해도 '물질명는' → 'AAA는'")
    # 면제조건 드롭다운 + 예시 제거
    b.set("면제_물용해도", "F14", "")
    b.add_list_dv("면제_물용해도", "F14", "면제!$C$16:$C$18")
    b.add_list_dv("면제_녹는점", "F14", "면제!$C$2:$C$2")
    b.add_list_dv("8.입도분석(면제)", "D14", "면제!$C$20:$C$20")
    b.add_list_dv("10.폭발성(면제)", "D12", "면제!$C$30:$C$35")
    b.add_list_dv("11.산화성(면제)", "D12", "면제!$C$37:$C$41")
    log("  - 면제 시트 5개(물용해도·녹는점·입도분석·폭발성·산화성) 면제조건 칸에 해당 항목 드롭다운 연결")
    clear_examples(b, wb, path, "면제_물용해도")
    renames = {
        "Description": "목차", "6.옥탄올": "6.옥탄올물분배계수", "9.인화성 (인화점)": "9.인화성",
        "면제_물용해도": "2.물용해도_면제", "면제_녹는점": "3.녹는점어는점_면제",
        "8.입도분석(면제)": "8.입도분석_면제", "10.폭발성(면제)": "10.폭발성_면제", "11.산화성(면제)": "11.산화성_면제",
        "Sheet3": "목록_선택값", "면제": "목록_면제조건",
    }
    order = ["목차", "1.물질의상태", "2.물용해도", "2.물용해도_면제", "3.녹는점어는점", "3.녹는점어는점_면제",
             "4.끓는점", "5.증기압", "6.옥탄올물분배계수", "7.밀도", "8.입도분석_면제", "9.인화성",
             "10.폭발성_면제", "11.산화성_면제", "목록_선택값", "목록_면제조건"]
    desc = {3: "1.물질의상태", 4: "2.물용해도\n2.물용해도_면제", 5: "3.녹는점어는점\n3.녹는점어는점_면제",
            6: "4.끓는점", 7: "5.증기압", 8: "6.옥탄올물분배계수", 9: "7.밀도", 10: "8.입도분석_면제",
            11: "9.인화성", 12: "10.폭발성_면제", 13: "11.산화성_면제", 14: "(양식 없음)", 15: "(양식 없음)"}
    finish(b, renames, order, ["목록_선택값", "목록_면제조건"], desc, "물리화학적 특성 시험항목")
    log("  - 시트명 '번호.항목[_생략|_면제]' 규칙으로 통일, 번호순 정렬, 목차 시트명 실제 탭과 일치, 목록 시트 숨김")
    b.save(os.path.join(DST, "phys.xlsx"))


# ======================= 인체 =======================
def do_human():
    path = os.path.join(SRC, "human.xlsx")
    wb = openpyxl.load_workbook(path)
    b = Book(path)
    log("■ 인체 (시험요약서00-00-0_인체2026_ver.1.xlsx)")
    common(b)
    # 생략_급성흡입
    b.set("생략_급성흡입", "D6", ""); b.set("생략_급성흡입", "D7", ""); b.set("생략_급성흡입", "D14", "")
    b.set("생략_급성흡입", "D8", "생략"); b.set("생략_급성흡입", "D23", "급성경피독성")
    log("  - 생략_급성흡입: 자료의 성격 '면제'→'생략', 첨부1 → '급성경피독성', 예시 물질명·근거자료 비움")
    # 면제조건 드롭다운
    for sheet, label, rng in [("22. In vivo_면제", "면제조건", "$C$64:$C$66"),
                              ("23.반복(28일)_면제", "면제조건", "$C$70:$C$72"),
                              ("24.스크리닝_면제", "면제조건", "$C$78:$C$82")]:
        ref = find_label_cell(wb, sheet, label, "D")
        b.add_list_dv(sheet, ref, "면제!" + rng)
    log("  - 면제 시트 3개(In vivo·반복28일·스크리닝) 면제조건 칸 드롭다운을 해당 항목 범위로 연결")
    b.set("23.반복(28일)_면제", "D15", "")
    b.set("22. In vivo_면제", "D16", "")
    log("  - [22·23 면제] 근거자료 예시 비움")
    clear_examples(b, wb, path, "26.90일 반복(생략)", extra=("F16",))
    clear_examples(b, wb, path, "생략_급성경피", extra=("F16",))
    renames = {
        "Description": "목차", "14.경구(생략)": "14.급성경구_생략", "17.과민성": "17.피부과민성",
        "18. 급성경피": "18.급성경피", "생략_급성경피": "18.급성경피_생략", "생략_급성흡입": "19.급성흡입_생략",
        "20.눈자극(생략)": "20.눈자극_생략", "21.In vitro": "21.염색체이상(In vitro)",
        "22. In vivo_면제": "22.유전독성(In vivo)_면제", "23.반복(28일)_면제": "23.반복투여28일_면제",
        "24.스크리닝_면제": "24.생식발달스크리닝_면제", "26.90일 반복(생략)": "26.반복투여90일_생략",
        "Sheet3": "목록_선택값", "면제": "목록_면제조건",
    }
    order = ["목차", "14.급성경구", "14.급성경구_생략", "15.복귀돌연변이", "16.피부자극성", "17.피부과민성",
             "18.급성경피", "18.급성경피_생략", "19.급성흡입_생략", "20.눈자극", "20.눈자극_생략",
             "21.염색체이상(In vitro)", "22.유전독성(In vivo)_면제", "23.반복투여28일_면제",
             "24.생식발달스크리닝_면제", "26.반복투여90일_생략", "27.최기형성", "목록_선택값", "목록_면제조건"]
    desc = {3: "14.급성경구\n14.급성경구_생략", 4: "15.복귀돌연변이", 5: "16.피부자극성", 6: "17.피부과민성",
            7: "18.급성경피\n18.급성경피_생략", 8: "19.급성흡입_생략", 9: "20.눈자극\n20.눈자극_생략",
            10: "21.염색체이상(In vitro)", 11: "22.유전독성(In vivo)_면제", 12: "23.반복투여28일_면제",
            13: "24.생식발달스크리닝_면제", 14: "(양식 없음)", 15: "26.반복투여90일_생략", 16: "27.최기형성",
            17: "(양식 없음)", 18: "(양식 없음)"}
    finish(b, renames, order, ["목록_선택값", "목록_면제조건"], desc, "인체 유해성 시험항목")
    log("  - 시트명 규칙 통일(번호 뒤 공백 제거 포함), 번호순 정렬, 목차 시트명 일치, 목록 시트 숨김")
    b.save(os.path.join(DST, "human.xlsx"))


# ======================= 환경 =======================
def do_env():
    path = os.path.join(SRC, "env.xlsx")
    wb = openpyxl.load_workbook(path)
    b = Book(path)
    log("■ 환경 (시험요약서00-00-0_환경2026_ver.1.xlsx)")
    common(b)
    for s in ("생략_이분해성", "생략_가수분해"):
        b.set(s, "F6", ""); b.set(s, "F7", ""); b.set(s, "F16", "")
        clear_examples(b, wb, path, s)
        b.add_list_dv(s, "F11", "목록!$M$4:$M$18")
    log("  - 생략 시트 2개: 시험생략조건(영제13조) 칸에 조문 드롭다운 연결, 예시 물질명·출처 비움")
    for r in range(3, 8):
        b.set("Description", f"E{r}", "")
    log("  - 목차 비고란의 특정 물질 작업 메모(시험생산·국평/EPA 생략 등) 비움")
    ref = find_label_cell(wb, "면제_가수분해", "면제조건", "D")
    b.add_list_dv("면제_가수분해", ref, "면제!$C$100:$C$103")
    log("  - 면제_가수분해 면제조건 칸 드롭다운 연결")
    renames = {
        "Description": "목차", "생략_이분해성": "31.이분해성_생략", "33. 담수조류": "33.담수조류",
        "34.가수분해(제출)": "34.가수분해", "생략_가수분해": "34.가수분해_생략", "면제_가수분해": "34.가수분해_면제",
        "목록": "목록_선택값", "방법": "목록_시험방법", "면제": "목록_면제조건", "생략": "목록_생략조건",
    }
    order = ["목차", "30.어류급성", "31.이분해성", "31.이분해성_생략", "32.물벼룩급성", "33.담수조류",
             "34.가수분해", "34.가수분해_생략", "34.가수분해_면제",
             "목록_선택값", "목록_시험방법", "목록_면제조건", "목록_생략조건"]
    desc = {3: "30.어류급성", 4: "31.이분해성\n31.이분해성_생략", 5: "32.물벼룩급성", 6: "33.담수조류",
            7: "34.가수분해\n34.가수분해_생략\n34.가수분해_면제"}
    for r in range(8, 22):
        desc[r] = "(양식 없음)"
    finish(b, renames, order, ["목록_선택값", "목록_시험방법", "목록_면제조건", "목록_생략조건"], desc,
           "환경 유해성 시험항목")
    log("  - 시트명 규칙 통일, 번호순 정렬, 목차 시트명('제출'/'생략' 표기 → 실제 탭명), 목록 시트 숨김")
    b.save(os.path.join(DST, "env.xlsx"))


do_phys(); do_human(); do_env()
open(os.path.join(DST, "변경내역.txt"), "w", encoding="utf8").write("\n".join(LOG))
print("\n".join(LOG))
