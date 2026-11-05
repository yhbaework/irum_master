# -*- coding: utf-8 -*-
"""대학정보공시 지침서(PDF)에서 공시항목별 공시시기·자료기준일을 뽑아 tcm1004m 적재용 CSV를 만든다.

사용법:
    python build_tcm1004m_당해연도.py [연도폴더]

    연도폴더를 생략하면 프로젝트 안의 4자리 숫자 폴더 중 가장 최근 것을 쓴다.
    연도폴더 아래(하위 폴더 포함)에서 파일명에 '지침서'가 들어간 PDF 1개를 찾는다.

    예) 2027/지침서/2027년 대학정보공시 지침서.pdf
        2027/지침서_2027.pdf

    연도는 파일명 속 연도(20xx)로 정하고, 없으면 연도 폴더명으로 정한다.

결과는 가공/tcm1004m_{연도}.csv (utf-8-sig) 로 쓴다.

대학/대학원 판별
    【공시기관】 네 칸(대학/대학원/전문대학/대학원대학) 중 해당 칸이 파란색으로 표시된다.
    글자 색으로 판별한다. 2024 지침서부터 이 형식이다. 2021~2023 지침서는 형식이 달라
    초기 적재 때 쓴 build_tcm1004m_과거연도.py(보관용)로만 처리했다.

'자료기준일은 작성지침 참고'인 항목은 작성지침에서 날짜를 자동으로 뽑지 않는다.
원문을 그대로 넣고 확인 목록에 띄운다. 값은 결과 CSV에서 직접 고친다.

결과 CSV가 이미 있으면 PDF를 읽기 전에 멈춘다(손으로 고친 CSV를 덮어쓰지 않도록).
다시 만들려면 기존 파일을 지우거나 이름을 바꾼다.

공시시기가 '수시(…)'이고 자료기준일이 없으면 괄호 안 값을 자료기준일로 옮기고 공시시기는 '수시'로 쓴다.

자료기준일의 날짜는 'YYYY. M. D.'로 맞춘다(날짜_정리). 기간 기호는 '~' 하나로, '2월 말일'은 날짜로 바꾸고,
'2025. 1. 1. ~ 12. 31.'처럼 '~'·'및'로 이어진 연도 없는 날짜에는 앞 날짜의 연도를 넣는다.

지침서 형식이 바뀌어 추출 중 멈추면 멈춘 사유와 함께 형식변경_안내 문구가 나온다.
"""
import calendar
import collections
import csv
import os
import re
import sys
from dataclasses import dataclass

import pdfplumber

TABLE = "tcm1004m"
OUT_DIR = "가공"
항목목록_파일 = "tcm1001m.csv"  # 참고용 대조 (변환 테이블명이 기존 테이블명과 몇 개 겹치는지)

# CSV 컬럼 순서 (tcm1004m)
컬럼 = [
    "disc_year", "univ_grad_gbn", "item_no", "item_nm", "table_name",
    "disc_period", "data_base_dt", "page_no", "inpt_dt", "mod_dt",
]

# 공시기관 판별용
기관명 = ("대학", "대학원", "전문대학", "대학원대학")
파란색 = (0.0, 0.0, 1.0)
# 이보다 앞선 지침서는 형식(분리형)이 달라 이 스크립트로 처리하지 않는다
통합형_시작연도 = 2024

# 추출 단계에서 멈추면 사유 뒤에 붙이는 안내
형식변경_안내 = ("지침서 형식이 변경되었다면 기존 스크립트로는 정보 추출이 불가능하며, "
                "기존 작업지침(tcm1004m_작업지침.md)을 참고하여 스크립트 수정이 필요합니다.")

# 페이지 오른쪽 끝에서 이 폭(pt) 안에서 시작하는 글자는 장 구분 탭으로 보고 버린다
오른쪽_여백 = 60

# 항목 제목: '2-가-1. 성적평가 결과', '6-나-(1) 전임교원 …', '8-바-1-(2)-1 교비회계 …'
항목번호_RE = re.compile(r"^(\d+-[가-힣](?:-\d+|-\(\d+\))*)(?:\.\s*|\s+)(.+)$")

# 【공시양식】 줄. 2021 11-가-2처럼 【공시양식 및 방법】으로 적힌 해도 있다
양식줄_RE = re.compile(r"^\s*【공시양식[^】]*】")

# 자료기준일 머리말: '자료기준일', '자료 기준일', '자료추출 기준일자'
자료기준일_RE = re.compile(r"자료\s*(?:추출\s*)?기준일자?")

# 항목번호의 한글 -> 테이블명 영문 (tcm1001m 테이블명 표기를 따른다)
한글_영문 = {
    "가": "GA", "나": "NA", "다": "DA", "라": "RA", "마": "MA", "바": "BA", "사": "SA",
    "아": "A", "자": "JA", "차": "CHA", "카": "KA", "타": "TA", "파": "PA", "하": "HA",
    "거": "GEO", "너": "NEO", "더": "DEO", "러": "REO", "머": "MEO", "버": "BEO",
    "서": "SEO", "어": "EO", "저": "JEO", "처": "CHEO",
}


def 정리(s):
    """연속 공백을 하나로, 앞뒤 공백 제거. 원문은 이 이상 손대지 않는다."""
    return re.sub(r"\s+", " ", s).strip()


def 표기정리(s):
    """CSV에 쓰기 전 표기 정리: ' : ' -> ': ', 연속 공백 -> 하나."""
    # 콜론 앞에 공백이 있을 때만 고친다 ('10:00', 'http://' 같은 값은 건드리지 않는다)
    return re.sub(r"\s+:\s*", ": ", 정리(s)).rstrip()


_기간기호 = str.maketrans({"～": "~", "∼": "~", "〜": "~"})

# 연도 있는 날짜 '2025.4.1.', '2025. 4. 1' (끝 마침표는 있을 수도 없을 수도)
_날짜_RE = re.compile(r"(?<!\d)(\d{4})\s*\.\s*(\d{1,2})\s*\.\s*(\d{1,2})(?!\d)(?:\s*\.)?")
# 표기를 맞춘 날짜 바로 뒤에 '~' 또는 '및'으로 이어진 연도 없는 날짜: '2025. 1. 1. ~ 12. 31.'
_이어진날짜_RE = re.compile(
    r"(?<!\d)(\d{4})\. (\d{1,2})\. (\d{1,2})\.(\s*(?:~|및)\s*)(?<!\d)(\d{1,2})\s*\.\s*(\d{1,2})(?!\d)(?:\s*\.)?"
    r"(?!\s*\.?\s*\d)"  # '26. 2. 28.'처럼 뒤에 숫자가 더 붙으면 날짜가 아니다
)
# 연도 없는 날짜 '4. 1.'
_월일_RE = re.compile(r"(?<!\d)\d{1,2}\s*\.\s*\d{1,2}\s*\.")


def 기간기호_통일(s):
    """기간 기호 '～'(전각), '∼', '〜'를 '~'로 바꾼다."""
    return s.translate(_기간기호)


def _연도_채우기(m):
    y, mo, d, 사이, mo2, d2 = int(m[1]), int(m[2]), int(m[3]), m[4], int(m[5]), int(m[6])
    if not (1 <= mo2 <= 12 and 1 <= d2 <= 31):  # 월·일이 될 수 없으면 날짜로 보지 않는다
        return m[0]
    # 기간의 끝 날짜가 시작보다 앞서면 해가 넘어간 것이다 ('2025. 9. 1. ~ 2. 28.')
    y2 = y + 1 if 사이.strip() == "~" and (mo2, d2) < (mo, d) else y
    return f"{y}. {mo}. {d}.{사이}{y2}. {mo2}. {d2}."


def 날짜_정리(s):
    """자료기준일의 날짜 표기를 'YYYY. M. D.'로 맞춘다 -> (정리한 값, 연도를 채우지 못한 날짜들).

    '2026. 2월 말일'은 그해 2월 마지막 날로 바꾼다. 연도 없는 날짜는 연도 있는 날짜 바로 뒤에
    '~'나 '및'으로 이어질 때만 그 연도를 넣는다. 그 밖의 연도 없는 날짜('1학기(4. 1.)', '매년 4. 1.',
    다른 항목의 날짜 뒤에 온 날짜)는 연도를 지어내지 않고 그대로 두어 돌려준다.
    """
    s = 기간기호_통일(s)
    s = re.sub(r"(\d{4})\.\s*2월\s*말일",
               lambda m: f"{m[1]}. 2. {29 if calendar.isleap(int(m[1])) else 28}.", s)
    s = _날짜_RE.sub(lambda m: f"{m[1]}. {int(m[2])}. {int(m[3])}.", s)
    while True:  # '및'이 여러 번 이어질 수 있다
        t = _이어진날짜_RE.sub(_연도_채우기, s)
        if t == s:
            break
        s = t
    return s, [정리(x) for x in _월일_RE.findall(_날짜_RE.sub("", s))]


def 괄호_이어붙이기(첫줄, 다음줄들, 최대줄수=3):
    """첫줄의 괄호가 닫히지 않았으면 닫힐 때까지 다음 줄을 공백으로 이어 붙인다."""
    text, i = 첫줄, 0
    while text.count("(") > text.count(")"):
        if i >= min(최대줄수, len(다음줄들)):
            raise ValueError(f"괄호가 닫히지 않습니다: {정리(text)[:60]}")
        text += " " + 다음줄들[i]
        i += 1
    return 정리(text)


def _본문(줄, 머리):
    """'【공시시기】: 2026년 8월' -> '2026년 8월'"""
    return re.sub(rf"^\s*【{머리}[^】]*】\s*:?\s*", "", 줄)


def 공시시기_추출(시기줄, 다음줄들):
    return 괄호_이어붙이기(_본문(시기줄, "공시시기"), 다음줄들)


def _짝괄호(text, 여는위치):
    """여는위치의 '('와 짝이 맞는 ')'의 위치."""
    깊이 = 0
    for i in range(여는위치, len(text)):
        if text[i] == "(":
            깊이 += 1
        elif text[i] == ")":
            깊이 -= 1
            if 깊이 == 0:
                return i
    raise ValueError(f"괄호 짝이 맞지 않습니다: {text[:60]}")


def _감싼괄호(text, 위치):
    """위치를 감싸는 가장 가까운 '('의 위치. 없으면 None."""
    깊이 = 0
    for i in range(위치 - 1, -1, -1):
        if text[i] == ")":
            깊이 += 1
        elif text[i] == "(":
            if 깊이 == 0:
                return i
            깊이 -= 1
    return None


def 자료기준일_추출(양식줄, 다음줄들):
    """【공시양식】 줄에서 자료기준일 원문을 뽑는다. 없으면 None.

    1) '(자료기준일 : …)'          -> 콜론 뒤부터 짝 괄호 직전까지
       ('자료 기준일 :', '자료추출 기준일자 :' 표기도 같다 - 2021~2023)
    2) '(자료기준일은 작성지침 참고)' -> 괄호 안 전체
    3) '다음 양식 적용 (입학정원 : …)' -> 괄호 안 전체 (자료기준일이라는 말 없이 괄호만 있는 경우)
    4) 자체 양식, PDF 파일 형태 등  -> None
    """
    text = 괄호_이어붙이기(_본문(양식줄, "공시양식"), 다음줄들)

    m = 자료기준일_RE.search(text)
    if m:
        여는 = _감싼괄호(text, m.start())
        안 = text[m.start():_짝괄호(text, 여는)] if 여는 is not None else text[m.start():]
        머리 = re.match(자료기준일_RE.pattern + r"\s*:\s*", 안)
        return 정리(안[머리.end():] if 머리 else 안)

    m = re.match(r"다음의?\s*양식\s*적용\s*\(", text)
    if m:
        여는 = m.end() - 1
        return 정리(text[여는 + 1:_짝괄호(text, 여는)])
    return None


def 테이블명_변환(구분, 항목번호):
    """'대학원', '2-가-1' -> 'GRAD_2GA1'

    하이픈 뒤 숫자는 붙이고 괄호 숫자는 '_n'. 괄호 다음에 오는 숫자도 '_n'.
    (8-바-1-(2)-1 -> 8BA1_2_1). 학과별 여부(_M)는 지침서 양식만으로는 DB 테이블과
    맞지 않아 붙이지 않는다. tcm1001m과 맞춰볼 때는 그쪽 이름에서 _M을 떼고 비교한다.
    """
    m = re.fullmatch(r"(\d+)-([가-힣])((?:-\d+|-\(\d+\))*)", 항목번호)
    if not m:
        raise SystemExit(f"항목번호 형식을 알 수 없습니다: {항목번호}")
    if m.group(2) not in 한글_영문:
        raise SystemExit(f"'{m.group(2)}'의 영문 표기가 없습니다 ({항목번호})\n  -> 스크립트 상단 한글_영문에 추가하세요.")

    이름 = m.group(1) + 한글_영문[m.group(2)]
    괄호뒤 = False
    for 조각 in re.findall(r"-(\(\d+\)|\d+)", m.group(3)):
        if 조각.startswith("("):
            이름 += "_" + 조각[1:-1]
            괄호뒤 = True
        else:
            이름 += ("_" if 괄호뒤 else "") + 조각

    접두어 = "GRAD_" if 구분 == "대학원" else "UNIV_"
    return 접두어 + 이름


def 항목제목_찾기(앞줄들):
    """'■ 공시지침' 바로 앞 줄들에서 가장 가까운 항목 제목을 찾는다 -> (항목번호, 항목명, 제목 줄 위치).

    제목 위에 앞쪽 바닥글의 쪽번호나 상위 제목이 올 수 있고, 제목이 두 줄로
    넘어가기도 한다(뒤 줄은 항목명에 이어 붙인다). 이어 붙일 줄이 2줄을 넘거나
    번호처럼 시작하면(제목 형식이 바뀌어 하위 제목을 못 읽은 경우) 상위 제목을
    대신 잡지 않도록 멈춘다.
    """
    for i in range(len(앞줄들) - 1, -1, -1):
        m = 항목번호_RE.match(정리(앞줄들[i]))
        if m:
            뒤 = [z for z in 앞줄들[i + 1:] if not re.fullmatch(r"\s*\d+\s*", z)]
            if len(뒤) > 2 or any(re.match(r"\s*\d+-[가-힣]", z) for z in 뒤):
                raise ValueError(f"제목 뒤에 알 수 없는 줄이 있습니다: {뒤}")
            return m.group(1), 정리(" ".join([m.group(2)] + 뒤)), i
    raise ValueError("항목 제목을 찾지 못했습니다")


# ── PDF 읽기와 블록 분리 ────────────────────────────────────────────────


@dataclass
class 줄:
    page_no: int
    text: str
    # 【공시기관】 줄에만 채운다: {낱말: 글자색}. 통합형 지침서의 대학/대학원 판별용
    기관색: dict | None = None


@dataclass
class 블록:
    """'■ 공시지침' 하나 = 공시항목 하나."""
    page_no: int
    item_no: str
    item_nm: str
    기관줄: 줄
    disc_period: str
    data_base_dt: str | None


def _색(c):
    """pdfplumber 글자색 -> 소수 3자리 튜플. 회색은 (0.602,), 파란색은 (0.0, 0.0, 1.0)"""
    if c is None:
        return ()
    if not isinstance(c, (tuple, list)):
        c = (c,)
    return tuple(round(float(v), 3) for v in c)


def 줄_읽기(pdf_path, pages=None):
    """PDF를 줄 단위로 읽는다. pages는 1부터 시작하는 쪽 번호 목록(생략 시 전체)."""
    결과 = []
    with pdfplumber.open(pdf_path) as pdf:
        대상 = pages or range(1, len(pdf.pages) + 1)
        for pno in 대상:
            page = pdf.pages[pno - 1]
            # 오른쪽 여백의 장 구분 탭 글자를 버린다. 2024 지침서는 탭이 x0 536부터, 본문은 x1 535까지다.
            # 탭을 남기면 '■ 공시지침 학교운영에'처럼 본문 줄 끝에 붙어 블록·제목 인식을 깨뜨린다.
            한계 = page.width - 오른쪽_여백
            page = page.filter(lambda o, 한계=한계: o.get("object_type") != "char" or o["x0"] < 한계)
            words = None
            for ln in page.extract_text_lines(strip=True, return_chars=False):
                색 = None
                if ln["text"].startswith("【공시기관】"):
                    if words is None:
                        words = page.extract_words(extra_attrs=["non_stroking_color"])
                    색 = {
                        w["text"]: _색(w["non_stroking_color"])
                        for w in words
                        if abs(w["top"] - ln["top"]) < 3 and w["x0"] >= ln["x0"] - 1
                    }
                결과.append(줄(pno, ln["text"], 색))
    return 결과


def _머리(text, 이름):
    return re.match(rf"^\s*■\s*{이름}\s*$", text) is not None


def 블록_나누기(줄들):
    """줄 목록을 '■ 공시지침' 블록으로 나눈다 -> (블록들, 무시한 구간의 쪽 번호들).

    【공시기관】이 없는 '■ 공시지침' 구간(지침서 앞쪽 설명문 등)은 무시하고 쪽 번호만 남긴다.
    """
    시작들 = [i for i, z in enumerate(줄들) if _머리(z.text, "공시지침")]
    블록들, 무시 = [], []
    for n, s in enumerate(시작들):
        끝 = 시작들[n + 1] if n + 1 < len(시작들) else len(줄들)
        구간 = 줄들[s + 1:끝]
        작성 = next((i for i, z in enumerate(구간) if _머리(z.text, "작성지침")), len(구간))
        지침 = 구간[:작성]
        쪽 = 줄들[s].page_no

        기관줄 = next((z for z in 지침 if z.text.startswith("【공시기관】")), None)
        if 기관줄 is None:
            # 공시시기·공시양식이 있으면 공시기관 줄을 못 읽은 항목이다. 조용히 넘기지 않는다
            if any(z.text.lstrip().startswith("【공시시기】") or 양식줄_RE.match(z.text) for z in 지침):
                raise SystemExit(f"{쪽}쪽: 【공시시기】는 있는데 【공시기관】 줄을 찾지 못했습니다")
            무시.append(쪽)
            continue

        앞 = 줄들[max(0, s - 5):s]
        try:
            item_no, item_nm, 위치 = 항목제목_찾기([z.text for z in 앞])
        except ValueError as e:
            raise SystemExit(f"{쪽}쪽: 항목 제목을 읽지 못했습니다 - {e} (앞줄: {[z.text for z in 앞]})")
        제목쪽 = 앞[위치].page_no
        # 제목과 '■ 공시지침' 사이에 쪽이 넘어가면 머리글·바닥글이 항목명에 붙는다
        if any(z.page_no != 제목쪽 for z in 앞[위치:]) or 쪽 != 제목쪽:
            raise SystemExit(f"{쪽}쪽 {item_no}: 제목과 '■ 공시지침'이 다른 쪽에 있습니다. 항목명을 확인하세요")

        텍스트 = [z.text for z in 지침]
        try:
            시기 = next(i for i, t in enumerate(텍스트) if t.startswith("【공시시기】"))
        except StopIteration:
            raise SystemExit(f"{쪽}쪽 {item_no}: 【공시시기】가 없습니다")
        try:
            disc_period = 공시시기_추출(텍스트[시기], 텍스트[시기 + 1:시기 + 4])
            양식들 = [i for i, t in enumerate(텍스트) if 양식줄_RE.match(t)]
            기준일들 = []
            for i in 양식들:
                v = 자료기준일_추출(텍스트[i], 텍스트[i + 1:i + 4])
                if v is not None and v not in 기준일들:
                    기준일들.append(v)
        except ValueError as e:
            raise SystemExit(f"{쪽}쪽 {item_no}: {e}")

        블록들.append(블록(
            제목쪽, item_no, item_nm, 기관줄, disc_period,
            " | ".join(기준일들) if 기준일들 else None,
        ))
    return 블록들, 무시


# ── 대학/대학원 판별과 행 생성 ──────────────────────────────────────────


def 기관_판별_통합형(기관줄):
    """2024~ 지침서: 【공시기관】 네 칸 중 파란색 칸이 해당 기관이다.

    회색은 연도마다 값이 조금 다르다(2024·2025 0.6, 2026 0.602). 그래서 성분이
    하나인 색(회색)이면 해당 아님으로 본다. 그 밖의 색은 추측하지 않고 멈춘다.
    """
    색 = 기관줄.기관색 or {}
    쪽 = 기관줄.page_no
    없는 = [n for n in 기관명 if n not in 색]
    if 없는:
        raise SystemExit(f"{쪽}쪽: 【공시기관】에서 {없는}을(를) 찾지 못했습니다 ({기관줄.text})")
    해당 = []
    for n in 기관명:
        if 색[n] == 파란색:
            해당.append(n)
        elif len(색[n]) != 1:
            raise SystemExit(f"{쪽}쪽: 【공시기관】 '{n}'의 글자색 {색[n]}을(를) 판별할 수 없습니다")
    if not 해당:
        raise SystemExit(f"{쪽}쪽: 【공시기관】에 파란색 표시가 없습니다")
    return [n for n in ("대학", "대학원") if n in 해당]


def 행_만들기(연도, 블록들):
    """블록 × 대학/대학원마다 행을 만든다 -> (행들, 보고)."""
    구분목록 = [기관_판별_통합형(k.기관줄) for k in 블록들]

    보고 = {"제외": [], "PDF제외": [], "수동확인": [], "자료기준일_빈값": [], "수시_기준일": [], "연도없는_날짜": []}
    행들, 위치 = [], {}
    for k, 구분들 in zip(블록들, 구분목록):
        if not 구분들:
            보고["제외"].append((k.item_no, k.page_no))
            continue
        # PDF 파일로만 공시하는 항목(항목명에 [PDF])은 DB 데이터가 없으므로 넣지 않는다
        if "[PDF]" in k.item_nm:
            보고["PDF제외"].append((k.item_no, k.page_no))
            continue
        for g in 구분들:
            if (g, k.item_no) in 위치:
                raise SystemExit(
                    f"{연도} {g} {k.item_no}가 두 번 나옵니다 (쪽 {위치[(g, k.item_no)]}, {k.page_no})"
                )
            위치[(g, k.item_no)] = k.page_no

            기준일 = k.data_base_dt
            시기 = k.disc_period
            if not 기준일 and 시기.startswith("수시("):
                # 수시 공시 항목은 자료기준일 대신 '수시(…)' 괄호 안의 공시 기준을 넣고, 공시시기는 '수시'만 남긴다
                기준일 = 정리(시기[3:_짝괄호(시기, 2)])
                시기 = "수시"
                보고["수시_기준일"].append((g, k.item_no, k.page_no))
            if 기준일 is None:
                보고["자료기준일_빈값"].append((g, k.item_no, k.page_no))
            elif re.search(r"작성\s*지침", 기준일):  # 2021 12-바는 '작성 지침'으로 띄어 쓴다
                보고["수동확인"].append((g, k.item_no, k.item_nm, k.page_no))
            기준일, 못채움 = 날짜_정리(표기정리(기준일 or ""))
            보고["연도없는_날짜"] += [(g, k.item_no, k.page_no, x) for x in 못채움]

            행들.append({
                "disc_year": str(연도),
                "univ_grad_gbn": g,
                "item_no": k.item_no,
                "item_nm": 표기정리(k.item_nm),
                "table_name": 테이블명_변환(g, k.item_no),
                "disc_period": 기간기호_통일(표기정리(시기)),
                "data_base_dt": 기준일,
                "page_no": str(k.page_no),
                "inpt_dt": "",
                "mod_dt": "",
            })

    return 행들, 보고


def 완료_검사(줄들, 블록들):
    """모든 【공시기관】은 정확히 하나의 블록(항목)이 되어야 한다. 못 읽은 항목을 조용히 넘기지 않는다.

    줄 첫머리가 아닌 곳(옆 칼럼이 같은 줄로 합쳐진 경우 등)에 있는 【공시기관】도 센다.
    """
    기관줄들 = [z for z in 줄들 if "【공시기관】" in z.text]
    if len(블록들) != len(기관줄들):
        잡힌 = {id(k.기관줄) for k in 블록들}
        남은 = [z.page_no for z in 기관줄들 if id(z) not in 잡힌]
        raise SystemExit(f"블록이 되지 못한 【공시기관】이 있습니다 (쪽 {남은})")


def 결과_검사(연도, 행들):
    """대학·대학원 행이 하나도 없으면 지침서 형식이 바뀐 것이다. 빈 CSV를 '완료'로 쓰지 않는다."""
    n = collections.Counter(r["univ_grad_gbn"] for r in 행들)
    if not n["대학"] or not n["대학원"]:
        raise SystemExit(
            f"{연도}: 대학 {n['대학']}행 / 대학원 {n['대학원']}행 - 지침서 형식이 바뀌었는지 확인하세요"
        )


# ── 파일 찾기·CSV 쓰기·실행 ────────────────────────────────────────────


def 지침서_찾기(폴더):
    """연도폴더 아래(하위 폴더 포함)에서 파일명에 '지침서'가 들어간 PDF 1개를 찾는다."""
    fs = sorted(
        os.path.join(루트, f)
        for 루트, _, 파일들 in os.walk(폴더)
        for f in 파일들
        if f.lower().endswith(".pdf") and "지침서" in f and not f.startswith("~$")
    )
    if not fs:
        raise SystemExit(f"'{폴더}'에 '지침서'가 들어간 PDF가 없습니다.")
    if len(fs) > 1:
        raise SystemExit(f"'{폴더}'에 지침서 PDF가 여러 개입니다. 하나만 두세요: {fs}")
    return fs[0]


def 연도_추출(path):
    """파일명 속 연도(20xx), 없으면 경로 중 4자리 연도 폴더 -> 연도.

    '2026년 대학정보공시 지침서.pdf' -> 2026, '2027/지침서/지침서.pdf' -> 2027.
    파일명 연도와 연도 폴더가 다르면 어느 해 지침서인지 알 수 없으므로 멈춘다.
    """
    m = re.search(r"(20\d{2})", os.path.basename(path))
    파일연도 = int(m.group(1)) if m else None
    폴더연도 = next(
        (int(p) for p in reversed(os.path.normpath(os.path.dirname(path)).split(os.sep)) if re.fullmatch(r"\d{4}", p)),
        None,
    )
    if 파일연도 and 폴더연도 and 파일연도 != 폴더연도:
        raise SystemExit(f"파일명 연도({파일연도})와 폴더 연도({폴더연도})가 다릅니다: {path}")
    연도 = 파일연도 or 폴더연도
    if not 연도:
        raise SystemExit(f"연도를 알 수 없습니다. 파일명에 연도(예: 2027)를 넣거나 연도 폴더에 두세요: {path}")
    return 연도


def 결과_경로(연도):
    return os.path.join(OUT_DIR, f"{TABLE}_{연도}.csv")


def csv_쓰기(행들, 경로):
    os.makedirs(os.path.dirname(경로), exist_ok=True)
    with open(경로, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=컬럼)
        w.writeheader()
        w.writerows(행들)


def _목록출력(제목, 항목들):
    print(f"  {제목}: {len(항목들)}건")
    for x in 항목들:
        print(f"    - {x}")


def 처리(pdf):
    연도 = 연도_추출(pdf)
    if 연도 < 통합형_시작연도:
        raise SystemExit(
            f"{연도} 지침서는 형식이 달라 이 스크립트로 처리하지 않습니다. "
            f"과거 연도는 build_tcm1004m_과거연도.py(보관용)를 참고하세요: {pdf}"
        )
    dst = 결과_경로(연도)
    if os.path.exists(dst):
        raise SystemExit(f"{dst}가 이미 있습니다. 다시 만들려면 기존 파일을 지우거나 이름을 바꾸세요.")
    print(f"\n[{연도}] {os.path.basename(pdf)}")

    # 여기서 멈추는 것은 대부분 지침서 형식이 바뀐 경우다. 사유 뒤에 안내를 붙인다
    try:
        줄들 = 줄_읽기(pdf)
        블록들, 무시 = 블록_나누기(줄들)
        print(f"  읽은 쪽 {줄들[-1].page_no if 줄들 else 0} / 블록 {len(블록들)}개")
        if 무시:
            print(f"  공시기관이 없어 건너뛴 '■ 공시지침' 구간: 쪽 {무시}")
        완료_검사(줄들, 블록들)

        행들, 보고 = 행_만들기(연도, 블록들)
        결과_검사(연도, 행들)
    except SystemExit as e:
        raise SystemExit(f"{e}\n\n{형식변경_안내}") from None
    except Exception:
        # 예상 못 한 오류는 traceback을 그대로 보인다
        print(f"\n{형식변경_안내}", file=sys.stderr)
        raise
    n = collections.Counter(r["univ_grad_gbn"] for r in 행들)
    print(f"  행: 대학 {n['대학']} / 대학원 {n['대학원']} / 제외 블록: 전문대학·대학원대학 전용 {len(보고['제외'])}, "
          f"[PDF] 항목 {len(보고['PDF제외'])}")

    # 공시시기에 가장 많이 나오는 연도가 파일명 연도와 같은지 본다 (파일명 오타 대비)
    연도들 = collections.Counter(y for r in 행들 for y in re.findall(r"(\d{4})년", r["disc_period"]))
    if 연도들 and int(연도들.most_common(1)[0][0]) != 연도:
        print(f"  ! 경고: 공시시기에 가장 많은 연도는 {연도들.most_common(1)[0][0]}년인데 파일명은 {연도}년입니다.")

    빈시기 = [(r["univ_grad_gbn"], r["item_no"], r["page_no"]) for r in 행들 if not r["disc_period"]]
    _목록출력("공시시기 빈 값", 빈시기)
    _목록출력("자료기준일 빈 값(자체 양식 등)", 보고["자료기준일_빈값"])
    _목록출력("자료기준일을 '수시(…)' 괄호 안 값으로 채움", 보고["수시_기준일"])
    _목록출력("자료기준일 확인 필요(작성지침 참고) -> 해당 쪽 【작성지침】을 보고 CSV에서 직접 입력", 보고["수동확인"])
    _목록출력("연도를 채우지 못한 날짜(원문 확인 후 필요하면 CSV에서 직접 수정)", 보고["연도없는_날짜"])

    if os.path.exists(항목목록_파일):
        with open(항목목록_파일, encoding="utf-8-sig", newline="") as f:
            기존 = {re.sub(r"_M$", "", r.get("table_name") or "") for r in csv.DictReader(f)}
        겹침 = sum(r["table_name"] in 기존 for r in 행들)
        print(f"  참고: 변환 테이블명 중 {항목목록_파일}에 있는 이름(_M 제외 비교) {겹침}개 / {len(행들)}행")

    csv_쓰기(행들, dst)
    print(f"  -> {dst} ({len(행들)}행)")


def main(args=None):
    args = sys.argv[1:] if args is None else args
    폴더 = args[0] if args else None
    if 폴더 is None:
        후보 = sorted(d for d in os.listdir(".") if re.fullmatch(r"\d{4}", d) and os.path.isdir(d))
        if not 후보:
            raise SystemExit("연도 폴더(예: 2027)를 찾지 못했습니다. 폴더명을 인자로 주세요.")
        폴더 = 후보[-1]
    if not os.path.isdir(폴더):
        raise SystemExit(f"폴더가 없습니다: {폴더}")
    print(f"대상 폴더: {폴더}")

    처리(지침서_찾기(폴더))
    print("\n완료")


if __name__ == "__main__":
    main()
