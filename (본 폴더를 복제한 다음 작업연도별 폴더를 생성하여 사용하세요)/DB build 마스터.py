# -*- coding: utf-8 -*-
"""대학공시 마스터 연간 작업: 학교마스터(Univ_master_m_base)와 지침서(Tcm1004m) 적재용 CSV를 만든다.

사용법 (연도 폴더 안에서):
    ..\\irum_master\\Scripts\\python.exe "DB build 마스터.py"

    인자는 받지 않는다. 이 스크립트 파일이 있는 폴더를 연도 폴더로 보고, 폴더 이름(4자리)을 연도로 쓴다.
    기본폴더를 복제해 폴더 이름을 연도(예: 2027)로 바꾼 뒤 그 안의 스크립트를 실행한다.

    실행하면 작업 번호를 묻는다.
        0. 전체 실행 (1~2)
        1. 마스터 베이스 생성   -> Univ_master_m_base
        2. 정보공시 지침서 정제 -> Tcm1004m
    필요한 파일이 없는 작업은 번호 옆에 '(… 누락 - 작업불가)'로 표시되고, 골라도 실행하지 않는다.

연도 폴더 구성:
    마스터_컬럼정의.csv          UNIV_MASTER_M_BASE 컬럼 정의 (예전 tcm1002m.csv)
    학과정보/학과정보_yyyyMMdd.xlsx|xls   1개 이상 (상반기·하반기)
    학교개황/학교개황_yyyyMMdd.xlsx|xls   1개 이상 (여러 차례 공개분)
    정보공시지침서/<…지침서…>.pdf        1개
    EXPORT/                       결과물 (없으면 만든다)

결과물 (<yyMMdd>는 실행한 날):
    EXPORT/Univ_master_m_base_<yyMMdd>_<연도>.csv
    EXPORT/Tcm1004m_<yyMMdd>_<연도>.csv

다시 실행할 때:
    학교마스터 - 실행할 때마다 그날 날짜로 새로 만든다(같은 날이면 덮어쓴다). 이전 날짜 파일은 남는다.
    지침서     - EXPORT에 그해 Tcm1004m 파일이 하나라도 있으면 건너뛴다(손으로 고친 결과를 지키기 위해).
                 다시 만들려면 기존 파일을 지우거나 EXPORT 밖으로 옮긴다.

0번은 학교마스터 -> 지침서 순서로 한다. 학교마스터에서 멈추면 지침서는 하지 않는다.
지침서에서 멈춰도 학교마스터 결과는 남고, 마지막 요약에 지침서 실패가 표시된다.

처리 규칙은 가이드_학교마스터.md, 가이드_지침서.md를 본다.
"""
import calendar
import collections
import csv
import datetime
import glob
import os
import re
import sys
import traceback
from dataclasses import dataclass

import pandas as pd
import pdfplumber

# ══ 공통 ══════════════════════════════════════════════════════════════

EXPORT_DIR = "EXPORT"
컬럼정의_파일 = "마스터_컬럼정의.csv"
학과정보_폴더 = "학과정보"
학교개황_폴더 = "학교개황"
지침서_폴더 = "정보공시지침서"


def 연도폴더_정하기():
    """이 스크립트 파일이 있는 폴더의 절대경로와 연도(폴더 이름, 4자리)를 돌려준다."""
    폴더 = os.path.dirname(os.path.abspath(__file__))
    이름 = os.path.basename(폴더)
    if not re.fullmatch(r"\d{4}", 이름):
        raise SystemExit(
            "연도 폴더(예: 2027) 안의 스크립트를 실행하세요. "
            "기본폴더를 복제해 폴더 이름을 연도로 바꾼 뒤 그 안에서 실행합니다.\n"
            f"  지금 폴더: {폴더}"
        )
    return 폴더, int(이름)


def 만든날짜():
    return datetime.date.today().strftime("%y%m%d")


# ══ 학교마스터 (UNIV_MASTER_M_BASE) ════════════════════════════════════

마스터_테이블 = "UNIV_MASTER_M_BASE"  # 마스터_컬럼정의.csv의 table_name
마스터_결과이름 = "Univ_master_m_base"

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

# 정의서(마스터_컬럼정의.csv)의 column_comment -> 표준명
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


def 엑셀_목록(폴더, 키워드):
    """폴더에서 키워드가 들어간 엑셀(.xlsx/.xls)을 모두 찾는다. 엑셀 임시파일(~$)은 뺀다.

    학교개황은 구형 .xls로 내려오기도 해서 두 확장자를 모두 본다.
    """
    return sorted(
        f
        for 패턴 in ("*.xlsx", "*.xls")
        for f in glob.glob(os.path.join(폴더, 패턴))
        if 키워드 in os.path.basename(f) and not os.path.basename(f).startswith("~$")
    )


def 파일들찾기(폴더, 키워드):
    """엑셀_목록과 같되, 1개도 없으면 멈춘다."""
    fs = 엑셀_목록(폴더, 키워드)
    if not fs:
        raise SystemExit(f"'{폴더}'에 '{키워드}' 파일이 없습니다. (.xlsx/.xls)")
    return fs


def 기준일_추출(path):
    """파일명에서 기준일자(YYYYMMDD 8자리)를 뽑는다.

    파일명은 '학교개황_20251121.xlsx' 형식으로 맞춰서 넣어야 한다 (밑줄 뒤 8자리).
    """
    m = re.search(r"_(\d{8})(?!\d)", os.path.basename(path))
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
                f"  -> '학교개황_20260722.xlsx' 형태여야 합니다."
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


def 학교마스터_만들기(연도, 날짜):
    """학과정보·학교개황으로 UNIV_MASTER_M_BASE CSV를 만들고 (경로, NOT NULL 위반 컬럼 수)를 돌려준다."""
    if not os.path.isfile(컬럼정의_파일):
        raise SystemExit(
            f"'{컬럼정의_파일}'이 연도 폴더에 없습니다. 기본폴더에서 복사해 넣으세요."
        )
    fs학과 = 파일들찾기(학과정보_폴더, "학과정보")
    fs개황 = 파일들찾기(학교개황_폴더, "학교개황")

    print("\n[원본 읽기]")
    print(f"  학과정보 {len(fs학과)}개")
    학과 = 학과정보_읽기(fs학과)
    print(f"  학교개황 {len(fs개황)}개")
    개황 = 학교개황_읽기(fs개황)
    print(f"  -> 학과정보 {len(학과):,}행 / 학교개황 {len(개황):,}행")

    # 학과정보의 조사년도, 학교개황 파일명의 기준연도, 폴더 연도가 모두 같은 해인지 확인한다
    조사년도 = 학과["조사년도"].dropna().unique().tolist()
    if len(조사년도) != 1:
        raise SystemExit(f"학과정보에 조사년도가 여러 개입니다: {조사년도}")
    조사년도 = 조사년도[0]
    if 조사년도 != str(연도):
        raise SystemExit(f"학과정보 조사년도({조사년도})와 폴더 연도({연도})가 다릅니다.")
    개황연도 = sorted({기준일_추출(p)[:4] for p in fs개황})
    if 개황연도 != [조사년도]:
        raise SystemExit(
            f"학과정보 조사년도({조사년도})와 학교개황 기준연도({', '.join(개황연도)})가 다릅니다."
        )
    print(f"  공시연도: {조사년도}")

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
    정의 = pd.read_csv(컬럼정의_파일, dtype=str, encoding="utf-8-sig")
    정의 = 정의[정의["table_name"] == 마스터_테이블].copy()
    if 정의.empty:
        raise SystemExit(f"{컬럼정의_파일}에 {마스터_테이블} 정의가 없습니다.")
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

    os.makedirs(EXPORT_DIR, exist_ok=True)
    dst = os.path.join(EXPORT_DIR, f"{마스터_결과이름}_{날짜}_{연도}.csv")
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

    print("\n  학교마스터 완료" if not 위반 else "\n  ! NOT NULL 위반이 있으니 적재 전에 확인하세요.")
    return dst, 위반


# ══ 지침서 (tcm1004m) ═════════════════════════════════════════════════

지침서_결과이름 = "Tcm1004m"

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
                "가이드_지침서.md를 참고하여 스크립트 수정이 필요합니다.")

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


def 지침서_목록(폴더):
    """폴더(정보공시지침서/, 하위 폴더 포함)에서 파일명에 '지침서'가 들어간 PDF를 모두 찾는다."""
    return sorted(
        os.path.join(루트, f)
        for 루트, _, 파일들 in os.walk(폴더)
        for f in 파일들
        if f.lower().endswith(".pdf") and "지침서" in f and not f.startswith("~$")
    )


def 지침서_찾기(폴더):
    """지침서_목록에서 PDF 1개를 고른다. 없거나 여러 개면 멈춘다."""
    fs = 지침서_목록(폴더)
    if not fs:
        raise SystemExit(f"'{폴더}'에 '지침서'가 들어간 PDF가 없습니다.")
    if len(fs) > 1:
        raise SystemExit(f"'{폴더}'에 지침서 PDF가 여러 개입니다. 하나만 두세요: {fs}")
    return fs[0]


def 연도_확인(path, 폴더연도):
    """파일명 속 연도(20xx)가 있으면 폴더 연도와 같은지 확인하고, 폴더 연도를 돌려준다.

    '2026년 대학정보공시 지침서.pdf' -> 2026. 파일명 연도와 폴더 연도가 다르면
    어느 해 지침서인지 알 수 없으므로 멈춘다.
    """
    m = re.search(r"(20\d{2})", os.path.basename(path))
    파일연도 = int(m.group(1)) if m else None
    if 파일연도 and 파일연도 != 폴더연도:
        raise SystemExit(f"파일명 연도({파일연도})와 폴더 연도({폴더연도})가 다릅니다: {path}")
    return 폴더연도


def 기존_지침서_결과(연도):
    """EXPORT에 이미 있는 그해 지침서 결과(날짜 무관)."""
    return sorted(glob.glob(os.path.join(EXPORT_DIR, f"{지침서_결과이름}_*_{연도}.csv")))


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


def 지침서_만들기(연도, 날짜):
    """지침서 PDF로 tcm1004m CSV를 만들고 그 경로를 돌려준다. 이미 있어 건너뛰면 None."""
    # 손으로 고친 결과를 지키기 위해, 그해 결과가 하나라도 있으면 PDF를 읽기 전에 건너뛴다
    기존 = 기존_지침서_결과(연도)
    if 기존:
        print(f"  지침서 결과가 이미 있어 건너뜁니다: {', '.join(os.path.basename(f) for f in 기존)}")
        print("  다시 만들려면 기존 파일을 지우거나 EXPORT 밖으로 옮기세요.")
        return None

    pdf = 지침서_찾기(지침서_폴더)
    연도 = 연도_확인(pdf, 연도)
    if 연도 < 통합형_시작연도:
        raise SystemExit(f"{연도} 지침서는 형식이 달라 이 스크립트로 처리하지 않습니다: {pdf}")
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
        print(f"\n{형식변경_안내}", flush=True)
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

    dst = os.path.join(EXPORT_DIR, f"{지침서_결과이름}_{날짜}_{연도}.csv")
    csv_쓰기(행들, dst)
    print(f"  -> {dst} ({len(행들)}행)")
    print("\n  지침서 완료")
    return dst


# ══ 실행 ══════════════════════════════════════════════════════════════


작업목록 = {
    "0": "전체 실행 (1~2)",
    "1": "마스터 베이스 생성",
    "2": "정보공시 지침서 정제",
}


def 작업_상태(연도):
    """작업 번호별로 메뉴 옆에 붙일 표시와 실행 가능 여부를 돌려준다. {번호: (표시, 가능)}

    실제 작업이 파일을 찾는 규칙(엑셀_목록, 지침서_목록)과 같은 규칙으로 미리 확인한다.
    """
    마스터_누락 = []
    if not os.path.isfile(컬럼정의_파일):
        마스터_누락.append(컬럼정의_파일)
    if not 엑셀_목록(학과정보_폴더, "학과정보"):
        마스터_누락.append("학과정보 파일")
    if not 엑셀_목록(학교개황_폴더, "학교개황"):
        마스터_누락.append("학교개황 파일")
    마스터_문제 = f"{', '.join(마스터_누락)} 누락" if 마스터_누락 else ""

    # 지침서 결과가 이미 있으면 PDF가 없어도 건너뛰기만 하므로 막지 않는다
    지침서_있음 = bool(기존_지침서_결과(연도))
    pdf수 = len(지침서_목록(지침서_폴더))
    if 지침서_있음:
        지침서_문제 = ""
    elif pdf수 == 0:
        지침서_문제 = "지침서 PDF 누락"
    elif pdf수 > 1:
        지침서_문제 = f"지침서 PDF {pdf수}개(1개만 필요)"
    else:
        지침서_문제 = ""

    def 표시(문제들, 참고=""):
        문제들 = [m for m in 문제들 if m]
        if 문제들:
            return f" ({' / '.join(문제들)} - 작업불가)", False
        return (f" ({참고})" if 참고 else ""), True

    건너뜀 = "결과 있음 - 건너뜀" if 지침서_있음 else ""
    return {
        "0": 표시([마스터_문제, 지침서_문제]),
        "1": 표시([마스터_문제]),
        "2": 표시([지침서_문제], 건너뜀),
    }


def main():
    # 콘솔 설정과 상관없이 한글이 깨지지 않게, 출력을 파일로 받아도 오류가 나지 않게 한다
    for 스트림 in (sys.stdout, sys.stderr):
        스트림.reconfigure(encoding="utf-8")
    폴더, 연도 = 연도폴더_정하기()  # 기본폴더에서 실행하면 번호를 묻기 전에 멈춘다
    os.chdir(폴더)
    날짜 = 만든날짜()
    print(f"대상 폴더: {폴더} (연도 {연도})\n")

    상태 = 작업_상태(연도)
    for 번호, 이름 in 작업목록.items():
        print(f"{번호}. {이름}{상태[번호][0]}")
    try:
        choice = input("\n작업 번호를 입력하세요: ").strip()
    except EOFError:
        choice = ""
    if choice not in 작업목록:
        raise SystemExit("잘못된 입력입니다.")
    if not 상태[choice][1]:
        raise SystemExit(f"{choice}번은 작업할 수 없습니다{상태[choice][0]}. 필요한 파일을 넣고 다시 실행하세요.")

    마스터 = 위반 = None
    if choice in ("0", "1"):
        print("\n===== 1. 마스터 베이스 생성 =====")
        마스터, 위반 = 학교마스터_만들기(연도, 날짜)  # 여기서 멈추면 지침서는 하지 않는다

    지침서 = 지침서실패 = None
    if choice in ("0", "2"):
        print("\n===== 2. 정보공시 지침서 정제 =====")
        try:
            지침서 = 지침서_만들기(연도, 날짜)
        except SystemExit as e:
            # stdout으로 내보내야 파일로 받아도 출력 순서가 유지된다
            print(f"\n{e}")
            지침서실패 = True
        except Exception:
            sys.stdout.flush()
            traceback.print_exc()
            sys.stderr.flush()
            지침서실패 = True

    print("\n[결과]")
    if choice in ("0", "1"):
        print(f"  학교마스터: {마스터}")
        if 위반:
            # 지침서 출력에 묻히지 않도록 요약에 다시 보여 준다
            print(f"              ! NOT NULL 위반 {위반}개 컬럼 - 적재 전에 위 [NOT NULL 컬럼 점검]을 확인하세요.")
    if choice in ("0", "2"):
        if 지침서실패:
            print("  지침서: 실패 (위 메시지를 확인하세요)")
        elif 지침서 is None:
            print(f"  지침서: 건너뜀 (이미 있음: {', '.join(os.path.basename(f) for f in 기존_지침서_결과(연도))})")
        else:
            print(f"  지침서: {지침서}")

    if 지침서실패:
        sys.exit(1)
    print("\n완료" if not 위반 else "\n완료 (학교마스터 NOT NULL 위반 확인 필요)")


if __name__ == "__main__":
    main()
