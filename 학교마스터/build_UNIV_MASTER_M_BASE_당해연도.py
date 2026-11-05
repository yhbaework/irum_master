# -*- coding: utf-8 -*-
"""당해 연도 대교협 파일 2개로 UNIV_MASTER_M_BASE를 생성한다.

사용법:
    python build_UNIV_MASTER_M_BASE_당해연도.py [연도폴더]

    연도폴더를 생략하면 프로젝트 안의 4자리 숫자 폴더 중 가장 최근 것을 쓴다.
    폴더에는 그 해 기준의 '학과정보' 파일과 '학교개황' 파일이 각각 1개 이상 있어야 한다.
    파일명은 '<구분>_yyyyMMdd 기준.<확장자>' 형식이어야 한다 (.xlsx / .xls).

    예) 2026/학과정보_20260219 기준.xlsx
        2026/학교개황_20260722 기준.xlsx

파일이 여러 개인 경우

    학과정보 - 상반기(조사차수 01)와 하반기(02)가 따로 공개되므로 2개가 될 수 있다.
               두 파일을 병합하며, 학교코드 + 학과코드가 같은 행은 차수가 큰 쪽
               (하반기)을 우선한다. 하반기에 없고 상반기에만 있는 학과는 그대로 살린다.
    학교개황 - 한 해에 여러 차례 공개될 수 있다. 행을 모두 합쳐두고, 조회할 때
               키별로 파일명 기준일자가 가장 나중인 값을 우선 적용한다.

대교협 원본은 해마다 컬럼명과 컬럼 구성이 바뀌므로 알려진 모든 표기를 표준명으로
매핑한다(ALIAS). 새로운 표기가 나타나면 매핑되지 않은 컬럼으로 오류를 내고 멈추니,
그때 ALIAS에 표기를 추가하면 된다.
"""
import glob
import os
import re
import sys

import pandas as pd

DEF_FILE = "tcm1002m.csv"
TABLE = "UNIV_MASTER_M_BASE"
OUT_DIR = "가공"

# DB에 적재할 때 자동으로 값이 채워지는 컬럼. 정의서상 NOT NULL이지만
# CSV에는 비워 두고, NOT NULL 점검에서도 제외한다.
자동입력_컬럼 = ("CREATE_DT", "MOD_DT")

# ── 원본 컬럼 표기 매핑 ────────────────────────────────────────────────
# 표준명 -> 연도별로 등장한 모든 표기 (공백·엑셀개행 제거 후 비교)
학과_ALIAS = {
    "조사년도": ["조사년도"],
    "조사차수": ["조사차수"],
    "학교명": ["학교명"],
    "학교코드": ["학교코드"],
    "대학구분": ["대학구분"],
    "학교구분": ["학교구분"],
    "설립구분": ["설립구분"],
    "본분교": ["본분교명", "본분교"],
    "지역": ["지역", "(대학)지역", "학교(지역)"],
    "단과대학명": ["단과대학명"],
    "단과대학코드": ["단과대학코드", "단과대학코드[SCM29]"],
    "학교별학과코드": ["학교별학과코드"],
    "학부·과(전공)명": ["학부·과(전공)명"],
    "주야간구분": ["주야구분", "주야간구분", "주야간구분명"],
    "학과특성": ["학과특성", "학과특성명"],
    "학과상태": ["학과상태", "학과상태명"],
    "대계열분류": ["표준분류대계열", "대계열분류"],
    "중계열분류": ["표준분류중계열", "중계열분류"],
    "소계열분류": ["표준분류소계열", "소계열분류"],
    "대학자체대계열": ["대학자체대계열", "대학자체5대계열"],
    "수업연한": ["수업연한", "수업연한명"],
    "학위과정": ["학위과정", "학위과정명"],
    "(학과)소재지": ["(학과)소재지"],
    "(학과)소재지(상세)": ["(학과)소재지(상세)"],
}

개황_ALIAS = {
    "학교구분": ["학교구분"],
    "학교코드": ["학교코드"],
    "학교명": ["학교명"],
    "본분교": ["본분교"],
    "학제": ["학제"],
    "지역": ["지역"],
    "설립구분": ["설립구분"],
    "관련법령": ["관련법령"],
    "법인명": ["법인명"],
    "학교상태": ["학교상태"],
    "학교명(한자)": ["중문학교명", "학교명(한자)"],
    "학교명(영문)": ["영문학교명", "학교명(영문)"],
    "주소": ["주소"],
    "영문주소": ["영문주소"],
    "중문주소": ["중문주소"],
    "우편번호": ["우편번호"],
    "학교개교일": ["학교개교일"],
    "학교홈페이지": ["학교홈페이지"],
    "총장명": ["총장(대학원장)명", "총장명"],
    "학교대표번호": ["학교대표번호"],
    "학교대표팩스번호": ["학교대표팩스번호"],
}

# 정의서(tcm1002m.csv)의 column_comment -> 표준명
MAPPING = {
    "공시연도": "조사년도",
    "학교코드": "학교코드",
    "학교": "학교명",
    "본분교": "본분교",
    "대학구분": "대학구분",
    "학교종류": "학교구분",
    "설립구분": "설립구분",
    "지역": "지역",
    "학교상태": "학교상태",
    "단과대학코드": "단과대학코드",
    "단과대학": "단과대학명",
    "학과코드": "학교별학과코드",
    "학과_전공": "학부·과(전공)명",
    "대학자체계열": "대학자체대계열",
    "표준분류_대계열": "대계열분류",
    "표준분류_중계열": "중계열분류",
    "표준분류_소계열": "소계열분류",
    "주야간구분": "주야간구분",
    "학과특성": "학과특성",
    "학과상태": "학과상태",
    "수업연한": "수업연한",
    "학위과정": "학위과정",
    "학과_소재지": "(학과)소재지",
    "학과_소재지_상세": "(학과)소재지(상세)",
    "조사차수": "조사차수",
}

# 학교개황으로도 못 채우는 건을 직접 지정한다. (학교코드, 표준명) -> 값
# 지금은 비어 있다. 접두어 보완(아래)으로 해결되지 않는 예외가 생기면 여기에
# 학교코드와 근거를 적어 등록한다.
수기값 = {}

# 학교개황에서 가져와 채우는 컬럼 (학과정보 원본에 없거나 연도별로 빠지는 항목)
보완대상 = ["본분교", "설립구분", "학교상태"]

# 학교명 접두어로 채울 컬럼.
# 같은 대학의 본교·분교·대학원은 설립구분이 같다는 성질을 이용한다.
# 학교상태(폐교 여부)와 본분교는 학교마다 다르므로 이 방식을 쓰면 안 된다.
접두어_보완대상 = {"설립구분"}


def norm(s):
    """컬럼명 정규화: 엑셀 개행문자(_x000D_)와 공백 제거."""
    return re.sub(r"\s+", "", str(s).replace("_x000D_", ""))


def 헤더행_찾기(path, sheet, lookup, 탐색행수=20):
    """머리글(자료기준일·안내문)을 건너뛰고 실제 컬럼명이 있는 행을 찾는다.

    대학알리미 학과정보는 위에 안내문이 3~4줄 붙어 있고, 학교개황은 첫 행이
    바로 컬럼명이다. 알려진 컬럼명이 가장 많이 맞는 행을 헤더로 본다.
    """
    raw = pd.read_excel(path, sheet_name=sheet, dtype=str, header=None, nrows=탐색행수)
    최다행, 최다수 = None, 0
    for i in range(len(raw)):
        n = sum(1 for v in raw.iloc[i] if pd.notna(v) and norm(v) in lookup)
        if n > 최다수:
            최다행, 최다수 = i, n
    if 최다수 < 3:
        raise SystemExit(
            f"컬럼명 행을 찾지 못했습니다: {os.path.basename(path)}\n"
            f"  -> 상위 {탐색행수}행 안에 알려진 컬럼명이 3개 이상 있어야 합니다."
        )
    return 최다행


def 읽기(path, alias):
    """엑셀 첫 시트를 표준 컬럼명으로 읽는다. 원본에 없는 컬럼은 결측으로 채운다."""
    lookup = {norm(v): std for std, vs in alias.items() for v in vs}
    sheet = pd.ExcelFile(path).sheet_names[0]
    헤더행 = 헤더행_찾기(path, sheet, lookup)
    if 헤더행:
        print(f"    (머리글 {헤더행}행 건너뜀)")
    df = pd.read_excel(path, sheet_name=sheet, dtype=str, header=헤더행)
    df = df.dropna(axis=1, how="all")  # 안내문 때문에 생기는 빈 컬럼 제거

    renamed, unknown = {}, []
    for c in df.columns:
        std = lookup.get(norm(c))
        if std is None:
            unknown.append(c)
        else:
            renamed[c] = std
    if unknown:
        raise SystemExit(
            f"매핑되지 않은 컬럼 {unknown} ({os.path.basename(path)})\n"
            f"  -> 스크립트 상단 ALIAS에 이 표기를 추가하세요."
        )
    df = df.rename(columns=renamed)

    없는컬럼 = [c for c in alias if c not in df.columns]
    for c in 없는컬럼:
        df[c] = pd.NA
    if 없는컬럼:
        print(f"    원본에 없는 컬럼: {없는컬럼}")

    df = df[list(alias.keys())]
    # 문자열 정리: 개행·중복공백 정리, 양끝 공백 제거, 빈 문자열은 결측 처리
    for c in df.columns:
        df[c] = df[c].astype("string").str.replace(r"\s+", " ", regex=True).str.strip()
    return df.replace({"": pd.NA, "nan": pd.NA}).reset_index(drop=True)


def 파일들찾기(폴더, 키워드):
    """폴더에서 키워드가 들어간 엑셀을 모두 찾는다. 1개 이상이면 된다.

    학교개황은 구형 .xls로 내려오기도 해서 두 확장자를 모두 본다.
    """
    fs = sorted(
        f
        for 패턴 in ("*.xlsx", "*.xls")
        for f in glob.glob(os.path.join(폴더, 패턴))
        if 키워드 in os.path.basename(f) and not os.path.basename(f).startswith("~$")
    )
    if not fs:
        raise SystemExit(f"'{폴더}'에 '{키워드}' 파일이 없습니다. (.xlsx/.xls)")
    return fs


def 기준일_추출(path):
    """파일명에서 기준일자(YYYYMMDD 8자리)를 뽑는다.

    파일명은 '학과정보_20251121 기준.xlsx' 형식으로 맞춰서 넣어야 한다.
    """
    m = re.search(r"(\d{8})\s*기준", os.path.basename(path))
    return m.group(1) if m else ""


def 학과정보_읽기(paths):
    """학과정보 파일들을 조사차수 순으로 병합한다.

    상반기(차수 01)와 하반기(02)가 따로 공개되므로 한 해에 2개가 될 수 있다.
    학교코드 + 학과코드가 같은 행은 **차수가 큰 쪽(하반기)을 우선**하고,
    하반기에 없고 상반기에만 있는 학과는 그대로 살린다.
    """
    조각 = []
    for p in paths:
        df = 읽기(p, 학과_ALIAS)
        차수 = sorted(df["조사차수"].dropna().unique().tolist())
        print(f"    {os.path.basename(p)}: {len(df):,}행 (조사차수 {','.join(차수)})")
        조각.append(df)

    학과 = pd.concat(조각, ignore_index=True)
    if len(paths) == 1:
        return 학과

    키 = ["학교코드", "학교별학과코드"]
    before = len(학과)
    # 차수 오름차순으로 정렬해 마지막(가장 큰 차수)만 남긴다
    학과 = (
        학과.sort_values("조사차수", kind="stable")
        .drop_duplicates(키, keep="last")
        .sort_index()
        .reset_index(drop=True)
    )
    print(f"    병합: {before:,}행 -> {len(학과):,}행 (중복 {before - len(학과):,}행은 최신 차수 채택)")
    print("    차수별 채택 행수: "
          + ", ".join(f"{k}={v:,}" for k, v in 학과["조사차수"].value_counts().sort_index().items()))
    return 학과


def 학교개황_읽기(paths):
    """학교개황 파일들을 기준일자와 함께 병합한다.

    여러 차례 공개될 수 있으므로 한 해에 여러 개가 될 수 있다. 행을 합쳐두고,
    조회 시점에 **키별로 가장 나중 기준일자의 값을 우선** 적용한다(`최신값_추리기`).
    """
    조각 = []
    for p in paths:
        df = 읽기(p, 개황_ALIAS)
        기준일 = 기준일_추출(p)
        if not 기준일:
            raise SystemExit(
                f"학교개황 파일명에서 기준일자를 찾지 못했습니다: {os.path.basename(p)}\n"
                f"  -> '..._260722 기준.xlsx' 형태여야 합니다."
            )
        df["_기준일자"] = 기준일
        print(f"    {os.path.basename(p)}: {len(df):,}행 (기준일자 {기준일})")
        조각.append(df)

    개황 = pd.concat(조각, ignore_index=True)
    if len(paths) > 1:
        print(f"    병합: {len(개황):,}행 (기준일자 "
              + ", ".join(sorted(개황["_기준일자"].unique())) + " — 나중 날짜 우선)")
    return 개황


def 최신값_추리기(개황, keys, col):
    """keys -> col 조회 테이블을 만든다. 키별로 가장 나중 기준일자의 값만 남긴다.

    같은 기준일자 안에서 값이 갈리는 키는 모호하므로 제외한다.
    (예: 동덕여자대학교 스타트업대학원은 폐교/기존 2행이지만 설립구분은 둘 다 사립 →
     값이 같으므로 채울 수 있고, 학교상태는 값이 달라 제외된다)
    """
    lut = 개황[keys + [col, "_기준일자"]].dropna(subset=keys + [col])
    if lut.empty:
        return lut.drop(columns="_기준일자"), []

    # 키별 최신 기준일자만 남긴다
    최신 = lut.groupby(keys, dropna=False)["_기준일자"].transform("max")
    lut = lut[lut["_기준일자"] == 최신].drop(columns="_기준일자").drop_duplicates()

    모호 = lut.duplicated(keys, keep=False)
    제외 = lut.loc[모호, "학교명"].unique().tolist() if "학교명" in lut.columns else []
    return lut[~모호], 제외


def 보완(학과, 개황, col):
    """학교개황 값으로 학과정보의 결측을 채운다.

    1) 학교코드      : 학교개황에 학교코드 컬럼이 있는 연도
    2) 학교명+본분교  : 학교코드 컬럼이 없는 연도 (2021·2026년 원본이 그랬다)
    당해 연도 파일만 쓰므로 다른 연도 값으로는 대체하지 않는다.
    """
    if 개황[col].isna().all():
        print(f"    학교개황에 {col} 값이 없어 보완할 수 없습니다.")
        return 학과

    for keys, label in (
        (["학교코드"], "학교개황(학교코드)"),
        (["학교명", "본분교"], "학교개황(학교명+본분교)"),
    ):
        if 학과[col].isna().sum() == 0:
            break
        if col in keys:  # 본분교는 자기 자신을 조인 키로 쓸 수 없다
            continue
        if 개황[keys].isna().any(axis=None) and 개황[keys].isna().all(axis=0).any():
            continue  # 그 해 학교개황에 해당 키 컬럼이 아예 없다

        lut, 제외 = 최신값_추리기(개황, keys, col)
        if 제외:
            print(f"    값이 상충해 제외: {제외}")

        값 = 학과[keys].merge(lut, on=keys, how="left")[col]
        # 이미 값이 있는 행은 학교개황 값과 일치해야 한다 (조인 키 검증)
        겹침 = (학과[col].notna() & 값.notna()).values
        불일치 = int((학과.loc[겹침, col].values != 값[겹침].values).sum())
        if 불일치:
            raise SystemExit(f"{col}: 기존 값과 학교개황 값이 {불일치}행 불일치 ({'+'.join(keys)})")

        대상 = (학과[col].isna() & 값.notna()).values
        학과.loc[대상, col] = 값[대상].values
        학과.loc[대상, col + "_보완"] = label
        print(f"    {col} <- {label}: {int(대상.sum()):,}행")
    return 학과


def 접두어(s):
    """학교명의 공백 앞 토큰. '강원대학교 산업대학원' -> '강원대학교'"""
    return s.str.split(" ").str[0]


def 접두어_보완(학과, 개황, col):
    """학교코드·학교명+본분교로 못 채운 행을 학교명 접두어로 채운다.

    같은 대학의 본교·분교·대학원은 설립구분이 같다. 2021~2026년 전 연도에서
    접두어별 설립구분이 예외 없이 유일함을 확인했다(접두어 418개, 상충 0건).
    본교/분교는 접두어 단계에서 이미 갈린다 — '고려대학교'와 '고려대학교(세종)'.

    다만 미래 연도에 접두어 안에서 값이 갈릴 수 있으므로, 그런 접두어는
    통째로 제외하고 결측으로 남긴다(잘못 채우는 것보다 비워두는 편이 낫다).
    """
    # 접두어별로 가장 나중 기준일자의 값을 쓴다 (학교개황이 여러 개일 수 있다)
    표 = pd.DataFrame(
        {
            "_접두어": 접두어(개황["학교명"]),
            col: 개황[col],
            "_기준일자": 개황["_기준일자"],
        }
    ).dropna()
    최신 = 표.groupby("_접두어")["_기준일자"].transform("max")
    lut = 표[표["_기준일자"] == 최신].drop(columns="_기준일자").drop_duplicates()

    상충 = lut.duplicated("_접두어", keep=False)
    if 상충.any():
        print(f"    접두어 안에서 값이 갈려 제외: {sorted(lut.loc[상충, '_접두어'].unique())}")
        lut = lut[~상충]

    p = 접두어(학과["학교명"])
    값 = pd.DataFrame({"_접두어": p}).merge(lut, on="_접두어", how="left")[col]

    # 이미 채워진 값과 어긋나는 접두어는 신뢰할 수 없으니 제외한다
    겹침 = (학과[col].notna() & 값.notna()).values
    if 겹침.any():
        불일치 = 학과.loc[겹침, col].values != 값[겹침].values
        if 불일치.any():
            나쁜 = sorted(set(p[겹침][불일치]))
            print(f"    기존 값과 어긋나 제외: {나쁜}")
            값 = 값.mask(p.isin(나쁜).values)

    대상 = (학과[col].isna() & 값.notna()).values
    학과.loc[대상, col] = 값[대상].values
    학과.loc[대상, col + "_보완"] = "학교개황(학교명 접두어)"
    print(f"    {col} <- 학교개황(학교명 접두어): {int(대상.sum()):,}행")
    return 학과


def main():
    폴더 = sys.argv[1] if len(sys.argv) > 1 else None
    if 폴더 is None:
        후보 = sorted(
            d for d in os.listdir(".") if re.fullmatch(r"\d{4}", d) and os.path.isdir(d)
        )
        if not 후보:
            raise SystemExit("연도 폴더(예: 2026)를 찾지 못했습니다. 폴더명을 인자로 주세요.")
        폴더 = 후보[-1]
    if not os.path.isdir(폴더):
        raise SystemExit(f"폴더가 없습니다: {폴더}")
    print(f"대상 폴더: {폴더}")

    fs학과 = 파일들찾기(폴더, "학과정보")
    fs개황 = 파일들찾기(폴더, "학교개황")

    print("\n[원본 읽기]")
    print(f"  학과정보 {len(fs학과)}개")
    학과 = 학과정보_읽기(fs학과)
    print(f"  학교개황 {len(fs개황)}개")
    개황 = 학교개황_읽기(fs개황)
    print(f"  -> 학과정보 {len(학과):,}행 / 학교개황 {len(개황):,}행")

    # 학과정보의 조사년도와 학교개황 파일명의 기준연도가 모두 같은 해인지 확인한다
    연도 = 학과["조사년도"].dropna().unique().tolist()
    if len(연도) != 1:
        raise SystemExit(f"학과정보에 조사년도가 여러 개입니다: {연도}")
    연도 = 연도[0]
    개황연도 = sorted({기준일_추출(p)[:4] for p in fs개황})
    if 개황연도 != [연도]:
        raise SystemExit(
            f"학과정보 조사년도({연도})와 학교개황 기준연도({', '.join(개황연도)})가 다릅니다."
        )
    print(f"  공시연도: {연도}")

    print("\n[학교개황으로 결측 보완]")
    학과["학교상태"] = pd.NA  # 학과정보에 없는 항목이라 전부 학교개황에서 가져온다
    for col in 보완대상:
        학과[col + "_보완"] = pd.NA
        n = int(학과[col].isna().sum())
        print(f"  {col}: 결측 {n:,}행")
        if n:
            학과 = 보완(학과, 개황, col)

        # 위 두 키로 못 채운 행은 학교명 접두어로 채운다
        if col in 접두어_보완대상 and 학과[col].isna().any():
            학과 = 접두어_보완(학과, 개황, col)

        for (코드, c), 값 in 수기값.items():
            if c != col:
                continue
            대상 = (학과["학교코드"] == 코드) & 학과[col].isna()
            if 대상.any():
                학과.loc[대상, col] = 값
                학과.loc[대상, col + "_보완"] = "학교개황(학교명+본분교)"
                print(f"    {col} <- 수기 지정값({코드}={값}): {int(대상.sum()):,}행")

        남음 = int(학과[col].isna().sum())
        if 남음:
            학교 = 학과.loc[학과[col].isna(), "학교명"].dropna().unique().tolist()
            print(f"    ! 잔여 결측 {남음:,}행 / 학교 {len(학교)}개: {학교[:5]}")

    print("\n[테이블 형태로 변환]")
    정의 = pd.read_csv(DEF_FILE, dtype=str, encoding="utf-8-sig")
    정의 = 정의[정의["table_name"] == TABLE].copy()
    if 정의.empty:
        raise SystemExit(f"{DEF_FILE}에 {TABLE} 정의가 없습니다.")
    정의["column_no"] = 정의["column_no"].astype(int)
    정의 = 정의.sort_values("column_no")
    print(f"  정의서 컬럼 수: {len(정의)}")

    out = pd.DataFrame(index=range(len(학과)))
    for _, r in 정의.iterrows():
        col, comment = r["column_name"], r["column_comment"]
        if col in 자동입력_컬럼:
            out[col] = pd.NA  # DB 적재 시 자동으로 채워진다
        elif comment in MAPPING:
            s = 학과[MAPPING[comment]]
            if r["data_type"] == "FLOAT":
                # 앞자리 0을 채운 코드값을 숫자로 정규화 (0000579 -> 579)
                s = pd.to_numeric(s, errors="coerce").astype("Int64")
            out[col] = s.values
        else:
            out[col] = pd.NA  # 예비컬럼1~30: 채울 원천이 없다

    if list(out.columns) != 정의["column_name"].tolist():
        raise SystemExit("컬럼 순서가 정의서와 다릅니다.")

    os.makedirs(OUT_DIR, exist_ok=True)
    dst = os.path.join(OUT_DIR, f"{TABLE}_{연도}.csv")
    out.to_csv(dst, index=False, encoding="utf-8-sig")
    print(f"  총 {len(out):,}행 x {out.shape[1]}열 -> {dst}")

    print("\n[NOT NULL 컬럼 점검]")
    위반 = 0
    for _, r in 정의[정의["notnull_yn"] == "Y"].iterrows():
        col, 설명 = r["column_name"], str(r["column_comment"])
        if col in 자동입력_컬럼:
            print(f"  {col:<10} {설명:<9} 건너뜀 (DB 적재 시 자동 입력)")
            continue
        n = int(out[col].isna().sum())
        위반 += n > 0
        print(f"  {col:<10} {설명:<9} {'OK' if n == 0 else f'결측 {n:,}행'}")

    print("\n완료" if not 위반 else "\n! NOT NULL 위반이 있으니 적재 전에 확인하세요.")


if __name__ == "__main__":
    main()
