# process_data.py
# GitHub Actions에서 매일 실행: xlsx 2개 → processed.json 생성

import pandas as pd
import json
import numpy as np
from datetime import datetime
import re
import os

# ═══════════════════════════════════════
# 설정
# ═══════════════════════════════════════
BASE_FILE  = "data/base.xlsx"
DAILY_FILE = "data/daily.xlsx"
OUT_FILE   = "data/processed.json"

HEADER_ROW = 3  # 4행 헤더 (0-based)

# ═══════════════════════════════════════
# ★ 실제 파일 기준으로 전면 수정
# ═══════════════════════════════════════
COL_IDX = {
    "사이트ID":       0,   # A
    "충전소ID":       1,   # B
    "충전기ID":       2,   # C
    "충전소명":       5,   # F
    "충전기상태":    12,   # M
    "주소1":          7,   # H
    "상세주소":       8,   # I
    "누적사용량":    18,   # S
    "설치년":        19,   # T
    "설치월":        20,   # U
    "제조사":        28,   # AC
    "충전기모델명":  29,   # AD  (모델명)
    "충전기모델ID":  30,   # AE  (모델ID)
    "충전기유형":    33,   # AH  (급속/완속)
    "충전용량":      35,   # AJ  (충전기용량)
    "경도":          38,   # AM  (X좌표)
    "위도":          39,   # AN  (Y좌표)
    "계약상태":      42,   # AQ
    "운영계약시작":  43,   # AR
    "운영계약종료":  44,   # AS
    "사용여부":      45,   # AT
}

# daily 파일용 (누적사용량만 필요)
COL_IDX_DAILY = {
    "충전기ID":      2,    # C
    "누적사용량":   18,    # S
}


# ═══════════════════════════════════════
# 유틸
# ═══════════════════════════════════════
def safe_float(v):
    try:
        return float(str(v).replace(",", "").strip())
    except:
        return 0.0

def read_excel_by_index(path: str, col_map: dict) -> pd.DataFrame:
    """헤더=4행, 인덱스 기반으로 필요한 열만 읽기"""
    df_raw = pd.read_excel(path, header=HEADER_ROW, dtype=str, engine="openpyxl")
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    df_raw = df_raw.dropna(how="all").reset_index(drop=True)
    
    all_cols = list(df_raw.columns)
    result   = pd.DataFrame(index=df_raw.index)
    
    for alias, idx in col_map.items():
        # 1순위: alias명으로 직접 탐색
        if alias in df_raw.columns:
            result[alias] = df_raw[alias]
        # 2순위: 인덱스로 탐색
        elif idx < len(all_cols):
            result[alias] = df_raw.iloc[:, idx]
        else:
            result[alias] = ""
    
    return result

def read_a3_date(path: str) -> str:
    """A3 셀에서 기준일시 읽기"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True, read_only=True)
        val = wb.active["A3"].value
        wb.close()
        if val is None:
            return "정보없음"
        if isinstance(val, datetime):
            return val.strftime("%Y-%m-%d %H:%M:%S")
        return str(val).strip()
    except:
        return "정보없음"

# ═══════════════════════════════════════
# 권역 분류
# ═══════════════════════════════════════
def classify_region_series(addresses: pd.Series) -> pd.Series:
    addr   = addresses.fillna("").astype(str).str.strip()
    result = pd.Series("기타", index=addr.index)

    incheon        = addr.str.contains("인천", na=False)
    incheon_detail = addr.str.contains("계양|남동|동구|미추홀|부평|연수|서구|중구|강화", na=False)
    result[incheon & incheon_detail]  = "수도권남서"
    result[incheon & ~incheon_detail] = "인천기타"

    sg = addr.str.contains("서울|경기", na=False) & ~incheon
    nw = addr.str.contains("고양|부천|김포|파주|은평|마포|서대문|양천|강서|용산|종로|중구", na=False)
    ne = addr.str.contains("도봉|노원|중랑|강북|성북|동대문|성동|광진|의정부|남양주|구리|양주|포천|동두천|가평|연천", na=False)
    se = addr.str.contains("강남|서초|송파|강동|성남|용인|하남|광주|안성|수원|평택|오산|이천|여주|양평", na=False)
    sw = addr.str.contains("구로|금천|영등포|동작|관악|의왕|광명|군포|과천|시흥|안산|안양|화성", na=False)

    result[sg & nw]                   = "수도권북서"
    result[sg & ne & ~nw]             = "수도권북동"
    result[sg & se & ~nw & ~ne]       = "수도권남동"
    result[sg & sw & ~nw & ~ne & ~se] = "수도권남서"
    result[sg & ~nw & ~ne & ~se & ~sw]= "수도권기타"

    not_sg = ~incheon & ~sg
    result[addr.str.contains("강원",                       na=False) & not_sg] = "강원권"
    result[addr.str.contains("충청|충남|충북|세종|대전",   na=False) & not_sg] = "충청권"
    result[addr.str.contains("경상|경남|경북|부산|대구|울산", na=False) & not_sg] = "경상권"
    result[addr.str.contains("전라|전남|전북|광주",        na=False) & not_sg] = "전라권"
    result[addr.str.contains("제주",                       na=False) & not_sg] = "제주권"
    return result

# ═══════════════════════════════════════
# 사이트 자동 그루핑
# ★ 같은 주소(도로명 기준) → 동일 사이트
# ═══════════════════════════════════════
def normalize_addr(addr: str) -> str:
    """주소 정규화: 공백·특수문자 제거, 동/호 이하 삭제"""
    addr = str(addr).strip()
    # 동/호수 이하 제거 (예: "101동 지하1층" → "101")
    addr = re.sub(r"\s*(지하|지상|B|F)?\s*\d+층.*$", "", addr)
    addr = re.sub(r"\s*\d+동.*$",  "", addr)
    addr = re.sub(r"\s{2,}", " ",  addr)
    return addr.strip()

def build_site_groups(df: pd.DataFrame) -> pd.Series:
    """
    주소1(H열) 기준으로 사이트 그루핑.
    같은 정규화 주소 → 같은 site_key.
    site_key = 정규화 주소 (없으면 충전소명 폴백)
    """
    addr_col = "주소1" if "주소1" in df.columns else None
    name_col = next((c for c in ["충전소명", "사이트명", "설치장소"] if c in df.columns), None)
    
    if addr_col:
        normalized = df[addr_col].fillna("").apply(normalize_addr)
    elif name_col:
        normalized = df[name_col].fillna("알수없음")
    else:
        normalized = pd.Series("알수없음", index=df.index)
    
    # 빈 값은 이름 폴백
    if name_col:
        fallback = df[name_col].fillna("알수없음")
        normalized = normalized.where(normalized != "", fallback)
    
    return normalized
# ═══════════════════════════════════════
# 분류 함수 (process_data.py에 추가)
# ═══════════════════════════════════════

# 컬럼 후보 목록
_MODEL_COL_CANDIDATES   = ['충전기모델ID', '모델ID']
_MODELNM_COL_CANDIDATES = ['충전기모델명', '모델명']
_TYPE_COL_CANDIDATES    = ['충전기유형', '급속/완속']
_KW_COL_CANDIDATES      = ['충전용량', '충전기용량']


def _pick_series(df, candidates):
    """후보 컬럼명 중 존재하는 첫 번째 열 반환"""
    for c in candidates:
        if c in df.columns:
            return df[c].fillna('').astype(str).str.strip()
    return pd.Series('', index=df.index)


def classify_model_vectorized(df):
    AD = _pick_series(df, _MODEL_COL_CANDIDATES)   # 모델ID
    AG = _pick_series(df, _MODELNM_COL_CANDIDATES) # 모델명
    AH = _pick_series(df, _TYPE_COL_CANDIDATES)    # 급속/완속
    AJ = _pick_series(df, _KW_COL_CANDIDATES)      # 용량

    is_fast = (AH == '급속')

    # ── 급속 분류 ──────────────────────────────────────────
    fast_conds = [
        # 스필 급속
        is_fast & AG.str.startswith('S0F1'),
        is_fast & AG.str.startswith('S0F5'),
        # PNE 급속
        is_fast & AG.str.startswith('EVQ-') & (AJ == '100'),
        is_fast & AG.str.startswith('EVQ-') & (AJ == '50'),
        is_fast & AG.str.startswith('MAXERO-200'),
        is_fast & AG.str.startswith('DP150'),
        # 애플망고(그린카) 급속
        is_fast & AG.str.startswith('AM-FCD'),
        # SK 급속
        is_fast & AG.str.startswith('FC100'),
        is_fast & AG.str.startswith('FC200'),
        # 코스텔 급속
        is_fast & AG.str.startswith('SFC-S050'),
        # 스필 SVI
        is_fast & AG.str.startswith('SVI-0F'),
        # 중앙제어 급속
        is_fast & AG.str.startswith('JC-6933'),
        # 알박 급속
        is_fast & AG.str.startswith('UK-QC50'),
    ]
    fast_vals = [
        '급속스필_100', '급속스필_50',
        '급속PNE_100',  '급속PNE_50',
        '급속PNE_200',  '급속PNE_150',
        '급속애플망고_200',
        '급속SK_100',   '급속SK_200',
        '급속코스텔_50',
        '급속스필_SVI',
        '급속중앙제어_50',
        '급속알박_50',
    ]

    result = pd.Series(
        np.select(fast_conds, fast_vals, default='__PENDING__'),
        index=df.index
    )
    result[(result == '__PENDING__') & is_fast] = '급속기타'

    # ── 완속 분류 ──────────────────────────────────────────
    slow = ~is_fast
    slow_conds = [
        # 알박 완속 (UK-NC7 시리즈 전체)
        slow & AG.str.startswith('UK-NC7'),
        # PNE 완속 (EVL 시리즈)
        slow & AG.str.startswith('EVL-3J10'),          # 10kW
        slow & (AG == 'EVL-1C07027A01'),               # 신형대 (1C)
        slow & AG.str.startswith('EVL-1107'),          # 신형대 (1107)
        slow & AG.str.startswith('EVL-'),              # 구형대 (나머지 EVL)
        # 스필 완속
        slow & AG.str.startswith('S0L'),               # 스필_7kW (S0L)
        slow & AG.str.startswith('S0W'),               # 스필_7kW (S0W)
        # 이카플러그
        slow & AG.str.startswith('CPT'),
        slow & AG.str.startswith('CPW'),
        # SK 완속
        slow & AG.str.startswith('SC7K'),
        # 중앙제어 완속
        slow & AG.str.startswith('JC-6111'),
        slow & AG.str.startswith('JC-6511'),
        # PNE 완속 EVS
        slow & AG.str.startswith('EVS'),
        # MAXERO 완속
        slow & AG.str.startswith('MAXERO-007'),
    ]
    slow_vals = [
        '알박완속',
        '10kW',
        '신형대',  '신형대',  '구형대',
        '스필_7kW', '스필_7kW',
        '이카플러그', '이카플러그',
        'SK_7kW',
        '중앙제어_7kW', '중앙제어_7kW',
        'PNE_7kW',
        '완속기타',
    ]

    slow_result = pd.Series(
        np.select(slow_conds, slow_vals, default='완속기타'),
        index=df.index
    )
    pending = (result == '__PENDING__')
    result[pending] = slow_result[pending]
    return result



def classify_region_series(addresses):
    addr   = addresses.fillna('').astype(str).str.strip()
    result = pd.Series('기타', index=addr.index)

    incheon        = addr.str.contains('인천', na=False)
    incheon_detail = addr.str.contains(
        '계양|남동|동구|미추홀|부평|연수|서구|중구|강화', na=False
    )
    result[incheon & incheon_detail]  = '수도권남서'
    result[incheon & ~incheon_detail] = '인천기타'

    sg = addr.str.contains('서울|경기', na=False) & ~incheon
    nw = addr.str.contains('고양|부천|김포|파주|은평|마포|서대문|양천|강서|용산|종로|중구', na=False)
    ne = addr.str.contains('도봉|노원|중랑|강북|성북|동대문|성동|광진|의정부|남양주|구리|양주|포천|동두천|가평|연천', na=False)
    se = addr.str.contains('강남|서초|송파|강동|성남|용인|하남|광주|안성|수원|평택|오산|이천|여주|양평', na=False)
    sw = addr.str.contains('구로|금천|영등포|동작|관악|의왕|광명|군포|과천|시흥|안산|안양|화성', na=False)

    result[sg & nw]                    = '수도권북서'
    result[sg & ne & ~nw]              = '수도권북동'
    result[sg & se & ~nw & ~ne]        = '수도권남동'
    result[sg & sw & ~nw & ~ne & ~se]  = '수도권남서'
    result[sg & ~nw & ~ne & ~se & ~sw] = '수도권기타'

    not_sg = ~incheon & ~sg
    result[addr.str.contains('강원',                        na=False) & not_sg] = '강원권'
    result[addr.str.contains('충청|충남|충북|세종|대전',    na=False) & not_sg] = '충청권'
    result[addr.str.contains('경상|경남|경북|부산|대구|울산', na=False) & not_sg] = '경상권'
    result[addr.str.contains('전라|전남|전북|광주',         na=False) & not_sg] = '전라권'
    result[addr.str.contains('제주',                        na=False) & not_sg] = '제주권'
    return result


def build_site_groups(df):
    addr_col = '주소1' if '주소1' in df.columns else None
    name_col = next(
        (c for c in ['충전소명', '사이트명', '설치장소'] if c in df.columns), None
    )
    if addr_col:
        normalized = df[addr_col].fillna('').apply(normalize_addr)
    elif name_col:
        normalized = df[name_col].fillna('알수없음')
    else:
        normalized = pd.Series('알수없음', index=df.index)

    if name_col:
        fallback   = df[name_col].fillna('알수없음')
        normalized = normalized.where(normalized != '', fallback)
    return normalized


def normalize_addr(addr):
    import re
    addr = str(addr).strip()
    addr = re.sub(r'\s*(지하|지상|B|F)?\s*\d+층.*$', '', addr)
    addr = re.sub(r'\s*\d+동.*$', '', addr)
    addr = re.sub(r'\s{2,}', ' ', addr)
    return addr.strip()

# ═══════════════════════════════════════
# 메인 전처리
# ═══════════════════════════════════════
def process():
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 전처리 시작")

    # ── 1) 두 파일 읽기 ───────────────────────────────
    print("  파일 #1 (기초) 읽는 중...")
    base_date = read_a3_date(BASE_FILE)
    df_base   = read_excel_by_index(BASE_FILE, COL_IDX)

    print("  파일 #2 (일일) 읽는 중...")
    daily_date = read_a3_date(DAILY_FILE)
    df_daily   = read_excel_by_index(DAILY_FILE, COL_IDX_DAILY)

    # ── 2) 날짜 파싱 ──────────────────────────────────
    def parse_snap_date(s):
        for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"]:
            try:
                return datetime.strptime(str(s).strip()[:19], fmt)
            except:
                continue
        return None

    dt_base  = parse_snap_date(base_date)  or datetime(2024, 1, 1)
    dt_daily = parse_snap_date(daily_date) or datetime.now()
    diff_days = max(1, (dt_daily - dt_base).days)

    print(f"  기초파일 날짜: {dt_base.date()} | 일일파일 날짜: {dt_daily.date()} | 차이: {diff_days}일")

    # ── 3) 누적사용량 숫자 변환 ───────────────────────
    df_base["누적사용량_base"]   = df_base["누적사용량"].apply(safe_float)
    df_daily["누적사용량_daily"] = df_daily["누적사용량"].apply(safe_float)
    df_daily["충전기ID"]         = df_daily["충전기ID"].astype(str).str.strip()
    df_base["충전기ID"]          = df_base["충전기ID"].astype(str).str.strip()

    # ── 4) 병합 ───────────────────────────────────────
    df = df_base.copy()
    df = df.merge(
        df_daily[["충전기ID", "누적사용량_daily"]],
        on="충전기ID", how="left"
    )
    df["누적사용량_daily"] = df["누적사용량_daily"].fillna(df["누적사용량_base"])

    # ── 5) 운영개월수 계산 ────────────────────────────
    # ★ 운영계약 시작일로 계산 (실제 컬럼명: 운영계약시작)
    df["운영계약시작_dt"] = pd.to_datetime(df["운영계약시작"], errors="coerce")
    df["운영계약종료_dt"] = pd.to_datetime(df["운영계약종료"], errors="coerce")
    now = datetime.now()

    df["운영개월수"] = (
        (now - df["운영계약시작_dt"]).dt.days / 30.44
    ).round(1)

    # ── 6) 사용량 계산 ────────────────────────────────
    # 전체기간 월평균: 운영 1개월 이상인 경우만 유효
    valid = df["운영개월수"] >= 1.0

    df["월사용량_전체"] = np.where(
        valid,
        (df["누적사용량_daily"] / df["운영개월수"]).round(2),
        np.nan
    )

    # 최신 1개월 월충전량
    diff_kwh = (df["누적사용량_daily"] - df["누적사용량_base"]).clip(lower=0)
    df["월사용량_최신"] = (diff_kwh / diff_days * 30).round(2)
    df["일사용량_최신"] = (diff_kwh / diff_days).round(3)

    # 전체기간 계산 불가 → 최신값으로 대체
    df["월사용량_전체"] = df["월사용량_전체"].fillna(df["월사용량_최신"])
    df["운영개월수"]    = df["운영개월수"].fillna(0)

    # ── 7) 모델 분류 ──────────────────────────────────
    # 실제 컬럼명에 맞게 rename
    rename_for_classify = {
        "충전기모델ID": "충전기모델ID",
        "충전기모델명": "충전기모델명",
        "급속/완속":    "충전기유형",   # ★ 실제 컬럼명
        "충전용량":     "충전용량",
    }
    df_classify = df.rename(columns=rename_for_classify)

    if "모델분류" not in df.columns or df["모델분류"].isna().all():
        df["모델분류"] = classify_model_vectorized(df_classify)

    # ── 8) 권역 분류 ──────────────────────────────────
    if "권역" not in df.columns or df["권역"].isna().all():
        df["권역"] = classify_region_series(df["주소1"])

    # ── 9) 사이트 자동 그루핑 ─────────────────────────
    df["사이트키"] = build_site_groups(df)

    # ── 10) 계약 상태 ─────────────────────────────────
    def contract_status(row):
        if pd.isna(row["운영계약종료_dt"]): return "정보없음"
        r = (row["운영계약종료_dt"] - now).days
        if r < 0:    return "계약만료"
        if r <= 90:  return "만료임박"
        if r <= 365: return "만료예정"
        return "정상운영"

    df["계약상태_계산"] = df.apply(contract_status, axis=1)
    df["잔여일수"]      = (df["운영계약종료_dt"] - now).dt.days.fillna(-9999).astype(int)

    # 원본 계약상태 컬럼이 있으면 계산값 우선
    df["계약상태"] = df["계약상태_계산"]

    # ── 9) 사이트별 집계 ──────────────────────────────
    site_agg = (
        df.groupby("사이트키")
        .agg(
            충전기수=("충전기ID",       "count"),
            총누적사용량=("누적사용량_daily", "sum"),
            월사용량_전체합=("월사용량_전체",  "sum"),
            월사용량_최신합=("월사용량_최신",  "sum"),
            월사용량_전체평균=("월사용량_전체", "mean"),
            월사용량_최신평균=("월사용량_최신", "mean"),
            권역=("권역",               "first"),
            충전소명=("충전소명",        "first"),
            주소=("주소1",              "first"),
        )
        .round(2)
        .reset_index()
    )
    
    # ── 10) 모델별 집계 ───────────────────────────────
    model_col = "모델분류" if "모델분류" in df.columns else "충전기모델명"
    if model_col in df.columns:
        model_agg = (
            df.groupby(model_col)
            .agg(
                충전기수=("충전기ID",         "count"),
                총누적사용량=("누적사용량_daily", "sum"),
                월사용량_전체합=("월사용량_전체",  "sum"),
                월사용량_최신합=("월사용량_최신",  "sum"),
                충전기당_전체평균=("월사용량_전체", "mean"),
                충전기당_최신평균=("월사용량_최신", "mean"),
            )
            .round(2)
            .reset_index()
            .rename(columns={model_col: "모델분류"})
        )
    else:
        model_agg = pd.DataFrame()
    
    # ── 11) 권역별 집계 ───────────────────────────────
    region_agg = (
        df.groupby("권역")
        .agg(
            충전기수=("충전기ID",         "count"),
            총누적사용량=("누적사용량_daily", "sum"),
            월사용량_전체합=("월사용량_전체",  "sum"),
            월사용량_최신합=("월사용량_최신",  "sum"),
            충전기당_전체평균=("월사용량_전체", "mean"),
            충전기당_최신평균=("월사용량_최신", "mean"),
        )
        .round(2)
        .reset_index()
    )
    
    # ── 12) 개별 충전기 레코드 준비 ───────────────────
    keep_cols = [
        "충전기ID", "충전소명", "사이트명", "설치장소", "상세주소",
        "주소1", "충전기상태", "운영계약시작", "운영계약종료",
        "충전기모델명", "충전기유형", "충전용량",
        "위도", "경도", "모델분류", "권역", "사이트키", "계약상태", "잔여일수",
        "누적사용량_base", "누적사용량_daily",
        "월사용량_전체", "월사용량_최신", "일사용량_최신",
        "운영개월수",
    ]
    avail_cols = [c for c in keep_cols if c in df.columns]
    df_out     = df[avail_cols].copy()
    
    # NaN → None (JSON 직렬화)
    df_out = df_out.where(pd.notnull(df_out), None)
    
    # ── 13) JSON 직렬화 ───────────────────────────────
    def to_serializable(obj):
        if isinstance(obj, (np.integer,)):  return int(obj)
        if isinstance(obj, (np.floating,)): return None if np.isnan(obj) else float(obj)
        if isinstance(obj, pd.Timestamp):   return obj.strftime("%Y-%m-%d")
        if isinstance(obj, datetime):       return obj.strftime("%Y-%m-%d %H:%M:%S")
        return obj
    
    output = {
        "meta": {
            "generated_at":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "base_date":     base_date,
            "daily_date":    daily_date,
            "diff_days":     diff_days,
            "total_chargers":len(df_out),
            "total_sites":   int(df["사이트키"].nunique()),
        },
        "chargers": json.loads(
            df_out.to_json(orient="records", force_ascii=False, default_handler=to_serializable)
        ),
        "sites":    json.loads(
            site_agg.to_json(orient="records", force_ascii=False, default_handler=to_serializable)
        ),
        "models":   json.loads(
            model_agg.to_json(orient="records", force_ascii=False, default_handler=to_serializable)
        ) if not model_agg.empty else [],
        "regions":  json.loads(
            region_agg.to_json(orient="records", force_ascii=False, default_handler=to_serializable)
        ),
    }
    
    os.makedirs("data", exist_ok=True)
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",", ":"))
    
    sz_kb = os.path.getsize(OUT_FILE) / 1024
    print(f"  ✅ {OUT_FILE} 생성 완료 ({sz_kb:.1f} KB)")
    print(f"  충전기: {len(df_out):,}대 | 사이트: {output['meta']['total_sites']:,}개")
    print(f"  모델: {len(output['models'])}종 | 권역: {len(output['regions'])}개")
    return True

if __name__ == "__main__":
    process()
