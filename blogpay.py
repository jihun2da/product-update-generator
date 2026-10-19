# -*- coding: utf-8 -*-
"""
블로그페이 파일 생성 모듈  (블로그페이_변환.py 의 변환 규칙을 그대로 옮긴 것)
=====================================================================
- 입력 : S파일(카페24 상품 양식 .xlsx, 1행이 열 제목) 의 bytes
- 양식 : 블로그페이 파일.xls 의 bytes  (앱 폴더의 '블로그페이_기본.xls')
- 출력 : (블로그페이 .xls bytes, 로그 목록)

원본 변환기(블로그페이_변환.py)와 같은 규칙으로 동작하며, 파일 대신 bytes 로 주고받도록만
바꿨습니다. pipeline.py 에서 한 줄로 불러 쓰며, 이 모듈이 실패해도 다른 결과물에는
영향을 주지 않습니다.
"""
import io
import re
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

import openpyxl
import xlrd
from xlutils.filter import process, XLRDReader, XLWTWriter

# =====================================================================
# 설정 (원본 변환기와 동일한 값)
# =====================================================================
START_ROW = 2            # 양식에서 상품 데이터가 시작되는 엑셀 행 번호

REMOVE_KC = True         # 브랜드명 뒤의 'KC' 를 지움 (디세인트KC -> 디세인트)
USE_ORIGINAL_DETAIL = True   # I열(상품 설명)에 원본 '상품 상세설명' 값을 그대로 넣음
OPTION_LINE_BREAK = True     # L열(옵션값) 옵션명 그룹 사이에 ' /' + 줄바꿈
OPTION_PRICE_SEP = ","       # M,N,O열은 콤마로만 연결

SALE_PRICE_HEADER = "소비자가"    # D열 판매가
SUPPLY_PRICE_HEADER = "공급가"    # F열 공급가
STOCK_QTY = 999                   # G열 재고수량
OPTION_PRICE_RATE = "1.6"         # 옵션가 = 옵션 공급가 x 1.6
OPTION_STOCK_QTY = 100            # 옵션 재고수량
OPTION_LIST_MODE = "VALUE"        # "VALUE"(옵션값 1개당 1개) / "COMBO"(조합 1개당 1개)

DETAIL_HTML = ('<p><img src="{url}" title="천도복숭아-상세.jpg">'
               '<br style="clear:both;">&nbsp;</p>')

FIXED_VALUES = {
    "J": "일반옵션",              # 옵션형태
    "P": "N",                     # 재고옵션 재고노출
    "Y": "Y",                     # 배송여부 (필수)
    "Z": "기본배송비/직접 수령",  # 배송비
    "AC": "Y",                    # 묶음배송
    "AI": "Y",                    # 상품전시
    "AJ": "Y",                    # 쇼핑몰 메인 진열
    "AK": "Y",                    # 결제수단 휴대전화
    "AL": "Y",                    # 결제수단 계좌이체
    "AM": "Y",                    # 결제수단 가상계좌
    "AN": "Y",                    # 결제수단 무통장
    "AO": "Y",                    # 결제수단 대리결제
    "AT": "Y",                    # 적립금 추가
}

REQUIRED_COLS = ["B", "D", "G", "Y"]   # 양식 2행에 '필수' 로 표시된 열

SOURCE_HEADERS = ["상품명", SALE_PRICE_HEADER, SUPPLY_PRICE_HEADER,
                  "옵션입력", "이미지등록(상세)", "상품 상세설명"]

XLS_MAX_CELL_CHARS = 32767   # .xls 한 칸에 넣을 수 있는 최대 글자 수
MAX_WARN_LINES = 30          # 로그에 상세 표시할 경고 최대 개수 (나머지는 건수만)


class BlogpayError(Exception):
    """블로그페이 파일 생성 중 사용자에게 알려야 하는 오류"""


# =====================================================================
# 도우미 함수
# =====================================================================
def col_idx(letter):
    """엑셀 열 이름(A, B, ..., AT) -> 0부터 시작하는 번호"""
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def to_int(value):
    """숫자/숫자 문자열 -> 정수. 비어 있으면 None. 숫자가 아니면 ValueError"""
    if value is None or str(value).strip() == "":
        return None
    try:
        return int(Decimal(str(value).replace(",", "").strip()))
    except InvalidOperation:
        raise ValueError(str(value))


def times_rate(amount):
    """옵션 공급가 x 1.6 -> 옵션가 (소수점 오차 없이 계산)"""
    result = Decimal(amount) * Decimal(OPTION_PRICE_RATE)
    return int(result.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def clean_name(raw):
    """상품명에서 <br> 뒤쪽(색상/사이즈 표기)을 떼고 앞부분만 사용, 브랜드 뒤 KC 제거"""
    if raw is None:
        return ""
    first = re.split(r"<br\s*/?>", str(raw), flags=re.IGNORECASE)[0]
    name = re.sub(r"\s+", " ", first).strip()
    if REMOVE_KC:
        name = re.sub(r"^(\S+?)KC(?=\s|$)", r"\1", name).strip()
    return name


def detail_html(raw):
    """원본 상세설명 HTML 에서 이미지 주소만 뽑아 블로그페이용 HTML 로 감쌈"""
    if raw is None:
        return ""
    m = re.search(r"<img[^>]*?src\s*=\s*[\"']([^\"']+)[\"']", str(raw), flags=re.IGNORECASE)
    if not m:
        return str(raw).strip()
    return DETAIL_HTML.format(url=m.group(1).strip())


def detail_value(raw):
    if raw is None:
        return ""
    if USE_ORIGINAL_DETAIL:
        return str(raw)
    return detail_html(raw)


def parse_options(raw):
    """카페24 옵션입력 문자열 해석
       '색상{노랑|백메란지}//사이즈{5|7|13(+1000)}'
       -> [('색상', [('노랑', 0), ('백메란지', 0)]), ('사이즈', [('5', 0), ('7', 0), ('13', 1000)])]
       맨 끝의 (+숫자) 만 추가금으로 인식합니다."""
    groups = []
    if raw is None or str(raw).strip() == "":
        return groups
    for part in str(raw).split("//"):
        m = re.match(r"^\s*(.*?)\s*\{(.*)\}\s*$", part, flags=re.DOTALL)
        if not m:
            continue
        title = m.group(1).strip()
        values = []
        for v in m.group(2).split("|"):
            v = v.strip()
            if v == "":
                continue
            pm = re.match(r"^(.*?)\s*\(\s*([+-])\s*([\d,]+)\s*\)\s*$", v)
            if pm:
                extra = int(pm.group(3).replace(",", ""))
                if pm.group(2) == "-":
                    extra = -extra
                values.append((pm.group(1).strip(), extra))
            else:
                values.append((v, 0))
        groups.append((title, values))
    return groups


def build_option_cells(groups):
    """옵션 정보 -> K(옵션명), L(옵션값), M(옵션 공급가), N(옵션가), O(옵션 재고수량)"""
    if not groups:
        return "", "", "", "", ""
    k = ", ".join(title for title, _ in groups)
    group_sep = " /\n" if OPTION_LINE_BREAK else " / "
    l = group_sep.join(",".join(name for name, _ in values) for _, values in groups)

    has_extra = any(extra != 0 for _, values in groups for _, extra in values)
    if not has_extra:                      # 추가금 있는 옵션이 하나도 없으면 M,N,O 는 비움
        return k, l, "", "", ""

    if OPTION_LIST_MODE == "COMBO":
        extras = [0]
        for _, values in groups:
            extras = [base + extra for base in extras for _, extra in values]
        m = ",".join(str(x) for x in extras)
        n = ",".join(str(times_rate(x)) for x in extras)
        o = ",".join(str(OPTION_STOCK_QTY) for _ in extras)
        return k, l, m, n, o

    price_sep = OPTION_PRICE_SEP
    m = price_sep.join(",".join(str(extra) for _, extra in values) for _, values in groups)
    n = price_sep.join(",".join(str(times_rate(extra)) for _, extra in values) for _, values in groups)
    o = price_sep.join(",".join(str(OPTION_STOCK_QTY) for _ in values) for _, values in groups)
    return k, l, m, n, o


# =====================================================================
# S파일 읽기
# =====================================================================
def read_source(sfile_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(sfile_bytes), data_only=True)
    ws = wb.worksheets[0]

    header = {}
    for c in range(1, ws.max_column + 1):
        name = ws.cell(1, c).value
        if name is not None and str(name).strip() not in header:
            header[str(name).strip()] = c

    missing = [h for h in SOURCE_HEADERS if h not in header]
    if missing:
        raise BlogpayError("S파일 1행에서 다음 열 제목을 찾지 못했습니다: " + ", ".join(missing))

    products = []
    for r in range(2, ws.max_row + 1):
        def get(name, _r=r):
            return ws.cell(_r, header[name]).value

        raw_name = get("상품명")
        if raw_name is None or str(raw_name).strip() == "":
            continue                                   # 상품명이 없는 빈 줄은 건너뜀

        bad = []                                       # 숫자로 읽지 못한 값 (경고용)

        def num(col_header, _get=get, _bad=bad):
            try:
                return to_int(_get(col_header))
            except ValueError as e:
                _bad.append("{} 값 '{}' 을(를) 숫자로 읽을 수 없습니다".format(col_header, e))
                return None

        groups = parse_options(get("옵션입력"))
        k, l, m, n, o = build_option_cells(groups)
        image = get("이미지등록(상세)")

        row = dict(FIXED_VALUES)
        row.update({
            "B": clean_name(raw_name),
            "D": num(SALE_PRICE_HEADER),
            "F": num(SUPPLY_PRICE_HEADER),
            "G": STOCK_QTY,
            "H": "" if image is None else str(image).strip(),
            "I": detail_value(get("상품 상세설명")),
            "K": k, "L": l, "M": m, "N": n, "O": o,
        })
        if not groups:                                 # 옵션이 아예 없으면 옵션형태도 비움
            row["J"] = ""
        row["_src_row"] = r
        row["_bad"] = bad
        products.append(row)
    return products


# =====================================================================
# 검증 (필수값 누락 등 -> 경고 문구 목록)
# =====================================================================
def check(products):
    warns = []
    for p in products:
        tag = "S파일 {}행 [{}]".format(p["_src_row"], p["B"])
        for msg in p.get("_bad", []):
            warns.append("{} : {}".format(tag, msg))
        for col in REQUIRED_COLS:
            if p.get(col) in (None, ""):
                warns.append("{} : 필수 {}열이 비어 있습니다".format(tag, col))
        if len(p["B"]) > 100:
            warns.append("{} : 상품명이 100자를 넘습니다 ({}자)".format(tag, len(p["B"])))
        if p["D"] is not None and p["D"] % 10 != 0:
            warns.append("{} : 판매가가 10원 단위가 아닙니다 ({})".format(tag, p["D"]))
        for col in ("H", "I", "F", "K", "L"):
            if p.get(col) in (None, ""):
                warns.append("{} : {}열 값이 없습니다".format(tag, col))
        if p.get("M") and OPTION_LIST_MODE != "COMBO":
            value_count = len([v for v in re.split(r"[,/]", p["L"]) if v.strip() != ""])
            for col in ("M", "N", "O"):
                token_count = len(p[col].split(","))
                if token_count != value_count:
                    warns.append("{} : {}열 값 개수({})가 옵션값 개수({})와 다릅니다".format(
                        tag, col, token_count, value_count))
        for col in ("H", "I", "L", "M", "N", "O"):
            if len(str(p.get(col) or "")) > XLS_MAX_CELL_CHARS:
                warns.append("{} : {}열이 .xls 한 칸 최대 글자수({})를 넘어 비워서 저장했습니다".format(
                    tag, col, XLS_MAX_CELL_CHARS))
    return warns


# =====================================================================
# 양식에 채워서 .xls bytes 로 저장
# =====================================================================
def write_output(template_bytes, products):
    rb = xlrd.open_workbook(file_contents=template_bytes, formatting_info=True)
    rs = rb.sheet_by_index(0)

    # 양식을 서식(글꼴/색/테두리/열너비/행높이) 그대로 복사
    writer = XLWTWriter()
    process(XLRDReader(rb, "template.xls"), writer)
    wb = writer.output[0][1]
    styles = writer.style_list
    ws = wb.get_sheet(0)

    first = START_ROW - 1                              # 0부터 세는 시작 행
    base_row = min(first + 1, rs.nrows - 1)            # 서식을 빌려올 '깨끗한' 데이터 행
    base_height = rs.rowinfo_map[base_row].height if base_row in rs.rowinfo_map else None

    def style_of(r, c):
        rr = r if r < rs.nrows else base_row
        return styles[rs.cell_xf_index(rr, c)]

    for i, product in enumerate(products):
        r = first + i
        if r >= rs.nrows and base_height:              # 양식에 없던 행은 행 높이도 맞춰 줌
            ws.row(r).height_mismatch = True
            ws.row(r).height = base_height
        values = {col_idx(k): v for k, v in product.items() if not k.startswith("_")}
        for c in range(rs.ncols):
            v = values.get(c)
            if v is None or v == "":
                ws.write(r, c, None, style_of(r, c))
            elif isinstance(v, str) and len(v) > XLS_MAX_CELL_CHARS:
                ws.write(r, c, None, style_of(r, c))   # 너무 긴 값은 비움 (check()에서 경고)
            else:
                ws.write(r, c, v, style_of(r, c))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =====================================================================
# 외부에서 부르는 함수
# =====================================================================
def build_blogpay_xls(sfile_bytes, template_bytes):
    """S파일 bytes + 블로그페이 양식 bytes -> (xls bytes, 로그 목록)"""
    products = read_source(sfile_bytes)
    if not products:
        raise BlogpayError("S파일에서 상품(상품명이 있는 행)을 한 건도 찾지 못했습니다.")

    xls_bytes = write_output(template_bytes, products)

    logs = []
    with_price = sum(1 for p in products if p["M"])
    logs.append(
        "[블로그페이] 상품 {}건 생성 (옵션가 있는 상품 {}건 / 없는 상품 {}건), 입력 범위 {}행 ~ {}행".format(
            len(products), with_price, len(products) - with_price,
            START_ROW, START_ROW + len(products) - 1)
    )

    warns = check(products)
    if warns:
        logs.append(
            "[블로그페이 검증] 확인이 필요한 항목 {}건 "
            "(검사 기준: 필수열 B/D/G/Y 비어있음, 상품명 100자 초과, 판매가 10원 단위, "
            "H/I/F/K/L 값 없음, 옵션 M·N·O 값 개수 불일치)".format(len(warns))
        )
        for w in warns[:MAX_WARN_LINES]:
            logs.append("[블로그페이 검증] " + w)
        if len(warns) > MAX_WARN_LINES:
            logs.append("[블로그페이 검증] ... 외 {}건 더 있습니다 (위 {}건만 표시)".format(
                len(warns) - MAX_WARN_LINES, MAX_WARN_LINES))
    else:
        logs.append("[블로그페이 검증] 필수값 검사 이상 없음")
    return xls_bytes, logs
