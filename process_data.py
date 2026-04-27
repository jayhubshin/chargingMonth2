# process_data.py
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
HEADER_ROW = 3

# COL_IDX에 시리얼번호 추가
COL_IDX = {
    "사이트ID":      0,
    "충전소ID":      1,
    "충전기ID":      2,
    "충전소명":      5,
    "충전기상태":   12,
    "주소1":         7,
    "상세주소":      8,
    "누적사용량":   18,
    "제조사":       28,
    "충전기모델명": 29,
    "충전기모델ID": 30,
    "시리얼번호":   32,   # ★ 추가
    "충전기유형":   33,
    "충전용량":     35,
    "경도":         38,
    "위도":         39,
    "계약상태":     42,
    "운영계약시작": 43,
    "운영계약종료": 44,
    "사용여부":     45,
}




COL_IDX_DAILY = {
    "충전기ID":    2,
    "누적사용량": 18,
}

# 전역 변수 수정
_MODEL_COL_CANDIDATES   = ['충전기모델ID', '모델ID']
_MODELNM_COL_CANDIDATES = ['충전기모델명', '모델명']
_TYPE_COL_CANDIDATES    = ['충전기유형', '급속/완속']
_KW_COL_CANDIDATES      = ['충전용량', '충전기용량']
_SN_COL_CANDIDATES      = ['시리얼번호']   # ★ 추가


REGION_ORDER = [
    '수도권북동','수도권북서','수도권남동','수도권남서',
    '수도권기타','인천기타','강원권','충청권','전라권','경상권','제주권','기타'
]

# ═══════════════════════════════════════
# 유틸 함수
# ═══════════════════════════════════════
def safe_float(v):
    try:
        return float(str(v).replace(",", "").strip())
    except:
        return 0.0


def parse_snap_date(s: str):
    """날짜 문자열 → datetime 변환"""
    if not s or s == "정보없음":
        return None
    match = re.search(r'(\d{4}-\d{2}-\d{2})', str(s))
    if match:
        try:
            return datetime.strptime(match.group(1), "%Y-%m-%d")
        except:
            pass
    return None


def read_a3_date(path: str) -> str:
    """A3 셀에서 날짜 문자열 추출"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True, read_only=True)
        val = wb.active["A3"].value
        wb.close()
        if val is None:
            return "정보없음"
        if isinstance(val, datetime):
            return val.strftime("%Y-%m-%d")
        s = str(val).strip()
        match = re.search(r'(\d{4}-\d{2}-\d{2})', s)
        if match:
            return match.group(1)
        return s
    except:
        return "정보없음"


def read_excel_by_index(path: str, col_map: dict) -> pd.DataFrame:
    df_raw = pd.read_excel(path, header=HEADER_ROW, dtype=str, engine="openpyxl")
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    df_raw = df_raw.dropna(how="all").reset_index(drop=True)
    all_cols = list(df_raw.columns)
    result = pd.DataFrame(index=df_raw.index)
    for alias, idx in col_map.items():
        if alias in df_raw.columns:
            result[alias] = df_raw[alias]
        elif idx < len(all_cols):
            result[alias] = df_raw.iloc[:, idx]
        else:
            result[alias] = ""
    return result


def normalize_addr(addr: str) -> str:
    addr = str(addr).strip()
    addr = re.sub(r'\s*(지하|지상|B|F)?\s*\d+층.*$', '', addr)
    addr = re.sub(r'\s*\d+동.*$', '', addr)
    addr = re.sub(r'\s{2,}', ' ', addr)
    return addr.strip()
    
def extract_site_name(name: str) -> str:
    """
    충전소명에서 대표 사이트명 추출
    예) 래미안아파트1 → 래미안아파트
        힐스테이트 101동 → 힐스테이트
        OO아파트-3 → OO아파트
        OO충전소_02 → OO충전소
    """
    if not name or str(name).strip() == '':
        return '알수없음'
    name = str(name).strip()
    # 뒤쪽 숫자/기호 제거 패턴
    # 예) -01, _02, 1호, 2번, (3), [4] 등
    name = re.sub(r'[\s\-_]+\d+호?기?번?$', '', name)
    name = re.sub(r'[\s\-_]+0\d+$',        '', name)
    name = re.sub(r'\s*\(\d+\)$',          '', name)
    name = re.sub(r'\s*\[\d+\]$',          '', name)
    name = re.sub(r'\s*\d+$',              '', name)
    return name.strip() or str(name).strip()

def build_site_groups(df: pd.DataFrame) -> tuple:
    """
    충전소명 기반으로 사이트 그루핑
    반환: (사이트키 Series, 대표사이트명 Series)
    """
    name_col = next(
        (c for c in ['충전소명', '사이트명', '설치장소'] if c in df.columns), None
    )
    addr_col = '주소1' if '주소1' in df.columns else None

    if name_col:
        # 대표사이트명 추출
        site_name = df[name_col].fillna('알수없음').apply(extract_site_name)
    else:
        site_name = pd.Series('알수없음', index=df.index)

    # 사이트키 = 대표사이트명 + 주소 앞부분 (동일 이름 다른 위치 구분)
    if addr_col:
        addr_short = df[addr_col].fillna('').apply(
            lambda x: ' '.join(str(x).split()[:3])  # 주소 앞 3단어
        )
        site_key = site_name + '|' + addr_short
    else:
        site_key = site_name.copy()

    return site_key, site_name

# ═══════════════════════════════════════
# 분류 함수
# ═══════════════════════════════════════
def _pick_series(df, candidates):
    for c in candidates:
        if c in df.columns:
            return df[c].fillna('').astype(str).str.strip()
    return pd.Series('', index=df.index)


def classify_model_vectorized(df):
    # AD=시리얼번호, AG=모델명, AH=급속/완속, AJ=용량
    AD = _pick_series(df, ['시리얼번호'])
    AG = _pick_series(df, ['충전기모델명', '모델명'])
    AH = _pick_series(df, ['충전기유형', '급속/완속'])
    AJ = _pick_series(df, ['충전용량', '충전기용량'])

    # ★ ad4~ad11 모두 시리얼번호(AD) 기준
    ad4  = AD.str[:4]
    ad3  = AD.str[:3]
    ad6  = AD.str[:6]
    ad11 = AD.str[:11]

    # 모델명 앞자리 (급속 분류용)
    ag4 = AG.str[:4]

    is_fast = (AH == '급속')

    # ── 급속 분류 (모델명 기준) ────────────────────────
    fast_conds = [
        is_fast & (ag4 == 'S0F1'),
        is_fast & (ag4 == 'S0F5'),
        is_fast & AG.str.startswith('EVQ-') & (AJ == '100'),
        is_fast & AG.str.startswith('EVQ-') & (AJ == '50'),
        is_fast & AG.str.startswith('MAXE'),
        is_fast & AG.str.startswith('DP15'),
        is_fast & AG.str.startswith('AM-F'),
        is_fast & AG.str.startswith('FC10'),
        is_fast & AG.str.startswith('FC20'),
        is_fast & AG.str.startswith('SFC-'),
        is_fast & AG.str.startswith('SVI-'),
        is_fast & AG.str.startswith('JC-69'),
        is_fast & AG.str.startswith('UK-Q'),
        is_fast & AG.str.startswith('CEC-'),
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
        '급속기타',
    ]
    result = pd.Series(
        np.select(fast_conds, fast_vals, default='__PENDING__'),
        index=df.index
    )
    result[(result == '__PENDING__') & is_fast] = '급속기타'

    # ── 완속 분류 (시리얼번호 AD 기준, 원본 로직 그대로) ──
    slow = ~is_fast
    slow_conds = [
        # 알박구형: 시리얼 앞4=NC07
        slow & (ad4 == 'NC07'),
        # 알박신형: 시리얼 앞4=23NA/22NA/24NA/25NA
        slow & (ad4.isin(['23NA', '22NA', '24NA', '25NA'])),
        # 10kW: 시리얼에 3J10 포함
        slow & AD.str.contains('3J10', na=False),
        # 신형대: 시리얼 앞11=EVL-1C-22CQ
        slow & (ad11 == 'EVL-1C-22CQ'),
        # 구형대: 시리얼 앞6=EVL-1C (22CQ 제외)
        slow & (ad6 == 'EVL-1C') & (ad11 != 'EVL-1C-22CQ'),
        # 신형대: 시리얼 앞4=EVL- + 시리얼에 1107 포함 (EVL-1C 제외)
        slow & (ad4 == 'EVL-') & AD.str.contains('1107', na=False) & (ad6 != 'EVL-1C'),
        # 구형대: 시리얼 앞4=EVL- + 시리얼에 1107 없음 (EVL-1C 제외)
        slow & (ad4 == 'EVL-') & ~AD.str.contains('1107', na=False) & (ad6 != 'EVL-1C'),
        # 신형대: 시리얼 앞4=SBDA
        slow & (ad4 == 'SBDA'),
        # 신형소: 시리얼 앞4=SBAA
        slow & (ad4 == 'SBAA'),
        # F01: 시리얼 앞4=SBPA + 시리얼에 F01 포함
        slow & (ad4 == 'SBPA') & AD.str.contains('F01', na=False),
        # PC01: 시리얼 앞4=SBPA + F01 없음
        slow & (ad4 == 'SBPA') & ~AD.str.contains('F01', na=False),
        # UC01: 시리얼 앞4=SBUA
        slow & (ad4 == 'SBUA'),
        # 스필_7kW: 시리얼 앞4=SVI0
        slow & (ad4 == 'SVI0'),
        # 이카플러그: 시리얼 앞3=E0C or 시리얼에 CP 포함
        slow & ((ad3 == 'E0C') | AD.str.contains('CP', na=False)),
        # 중앙제어_7kW: 시리얼 앞4=1907/1912
        slow & (ad4.isin(['1907', '1912'])),
        # SK_7kW: 시리얼 앞4=SC-P
        slow & (ad4 == 'SC-P'),
        # 3kW: 시리얼 앞4=SANA
        slow & (ad4 == 'SANA'),
        # PNE_7kW: 시리얼 앞4=EVS-/007S
        slow & (ad4.isin(['EVS-', '007S'])),
        # F01: 시리얼 앞4=SBOA + F01 포함
        slow & (ad4 == 'SBOA') & AD.str.contains('F01', na=False),
        # PC01: 시리얼 앞4=SBOA + F01 없음
        slow & (ad4 == 'SBOA') & ~AD.str.contains('F01', na=False),
    ]
    slow_vals = [
        '알박구형', '알박신형', '10kW',
        '신형대',   '구형대',
        '신형대',   '구형대',
        '신형대',   '신형소',
        'F01',      'PC01',     'UC01',
        '스필_7kW', '이카플러그',
        '중앙제어_7kW', 'SK_7kW', '3kW', 'PNE_7kW',
        'F01',      'PC01',
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
    result[addr.str.contains('강원',                         na=False) & not_sg] = '강원권'
    result[addr.str.contains('충청|충남|충북|세종|대전',     na=False) & not_sg] = '충청권'
    result[addr.str.contains('경상|경남|경북|부산|대구|울산',na=False) & not_sg] = '경상권'
    result[addr.str.contains('전라|전남|전북|광주',          na=False) & not_sg] = '전라권'
    result[addr.str.contains('제주',                         na=False) & not_sg] = '제주권'
    return result

# ═══════════════════════════════════════
# 메인 전처리
# ═══════════════════════════════════════
def process():
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 전처리 시작")

    # ── 1) 파일 읽기 ──────────────────────────────────
    print("  파일 #1 (기초) 읽는 중...")
    base_date = read_a3_date(BASE_FILE)
    df_base   = read_excel_by_index(BASE_FILE, COL_IDX)

    print("  파일 #2 (일일) 읽는 중...")
    daily_date = read_a3_date(DAILY_FILE)
    df_daily   = read_excel_by_index(DAILY_FILE, COL_IDX_DAILY)

    # ── 2) 날짜 파싱 ──────────────────────────────────
    dt_base  = parse_snap_date(base_date)
    dt_daily = parse_snap_date(daily_date)

    print(f"  base_date: '{base_date}' → {dt_base}")
    print(f"  daily_date: '{daily_date}' → {dt_daily}")

    if dt_base is None:
        raise ValueError(f"기초파일 날짜 읽기 실패: '{base_date}'")
    if dt_daily is None:
        raise ValueError(f"일일파일 날짜 읽기 실패: '{daily_date}'")

    diff_days = max(1, (dt_daily - dt_base).days)
    print(f"  기초: {dt_base.date()} | 일일: {dt_daily.date()} | 차이: {diff_days}일")

    # ── 3) 누적사용량 변환 ────────────────────────────
    df_base["누적사용량_base"]   = df_base["누적사용량"].apply(safe_float)
    df_daily["누적사용량_daily"] = df_daily["누적사용량"].apply(safe_float)
    df_daily["충전기ID"] = df_daily["충전기ID"].astype(str).str.strip()
    df_base["충전기ID"]  = df_base["충전기ID"].astype(str).str.strip()

    # ── 4) 병합 ───────────────────────────────────────
    df = df_base.copy()
    df = df.merge(
        df_daily[["충전기ID", "누적사용량_daily"]],
        on="충전기ID", how="left"
    )
    df["누적사용량_daily"] = df["누적사용량_daily"].fillna(df["누적사용량_base"])

    # ── 5) 운영개월수 ─────────────────────────────────
    df["운영계약시작_dt"] = pd.to_datetime(df["운영계약시작"], errors="coerce")
    df["운영계약종료_dt"] = pd.to_datetime(df["운영계약종료"], errors="coerce")
    now = datetime.now()

    df["운영개월수"] = (
        (now - df["운영계약시작_dt"]).dt.days / 30.44
    ).round(1)

    # ── 6) 사용량 계산 ────────────────────────────────
    valid = df["운영개월수"] >= 1.0
    df["월사용량_전체"] = np.where(
        valid,
        (df["누적사용량_daily"] / df["운영개월수"]).round(2),
        np.nan
    )
    diff_kwh = (df["누적사용량_daily"] - df["누적사용량_base"]).clip(lower=0)
    df["월사용량_최신"] = (diff_kwh / diff_days * 30).round(2)
    df["일사용량_최신"] = (diff_kwh / diff_days).round(3)
    df["월사용량_전체"] = df["월사용량_전체"].fillna(df["월사용량_최신"])
    df["운영개월수"]    = df["운영개월수"].fillna(0)

    # ── 7) 모델 분류 ──────────────────────────────────
    df["모델분류"] = classify_model_vectorized(df)

    # ── 8) 권역 분류 ──────────────────────────────────
    df["권역"] = classify_region_series(df["주소1"])

    # ── 9) 사이트 그루핑 ──────────────────────────────
    df["사이트키"], df["대표사이트명"] = build_site_groups(df)

    # ── 10) 계약 상태 ─────────────────────────────────
    def contract_status(row):
        if pd.isna(row["운영계약종료_dt"]): return "정보없음"
        r = (row["운영계약종료_dt"] - now).days
        if r < 0:    return "계약만료"
        if r <= 90:  return "만료임박"
        if r <= 365: return "만료예정"
        return "정상운영"
    
    df["계약상태"] = df.apply(contract_status, axis=1)
    df["잔여일수"] = (df["운영계약종료_dt"] - now).dt.days.fillna(-9999).astype(int)

    # ── 11) 사이트별 집계 ─────────────────────────────
    site_agg = (
        df.groupby("사이트키")
        .agg(
            대표사이트명=("대표사이트명",     "first"),
            충전소명=("충전소명",             "first"),
            충전기수=("충전기ID",             "count"),
            총누적사용량=("누적사용량_daily", "sum"),
            월사용량_전체합=("월사용량_전체", "sum"),
            월사용량_최신합=("월사용량_최신", "sum"),
            월사용량_전체평균=("월사용량_전체","mean"),
            월사용량_최신평균=("월사용량_최신","mean"),
            권역=("권역",                     "first"),
            주소=("주소1",                    "first"),
            상세주소=("상세주소",             "first"),
            계약상태=("계약상태",             "first"),      # ★
            운영계약시작=("운영계약시작",     "first"),      # ★
            운영계약종료=("운영계약종료",     "first"),      # ★
            잔여일수=("잔여일수",             "first"),      # ★
        )
        .round(2)
        .reset_index()
    )

    # ── 12) 모델별 집계 ───────────────────────────────
    model_agg = (
        df.groupby("모델분류")
        .agg(
            충전기수=("충전기ID",           "count"),
            총누적사용량=("누적사용량_daily","sum"),
            월사용량_전체합=("월사용량_전체","sum"),
            월사용량_최신합=("월사용량_최신","sum"),
            충전기당_전체평균=("월사용량_전체","mean"),
            충전기당_최신평균=("월사용량_최신","mean"),
        )
        .round(2)
        .reset_index()
    )

    # ── 13) 권역별 집계 ───────────────────────────────
    region_agg = (
        df.groupby("권역")
        .agg(
            충전기수=("충전기ID",           "count"),
            총누적사용량=("누적사용량_daily","sum"),
            월사용량_전체합=("월사용량_전체","sum"),
            월사용량_최신합=("월사용량_최신","sum"),
            충전기당_전체평균=("월사용량_전체","mean"),
            충전기당_최신평균=("월사용량_최신","mean"),
        )
        .round(2)
        .reset_index()
    )

    # ── 14) 충전기 레코드 정리 ────────────────────────
    keep_cols = [
        "충전기ID", "충전소명", "대표사이트명", "상세주소", "주소1",
        "충전기상태", "운영계약시작", "운영계약종료",
        "충전기모델명", "충전기유형", "충전용량",
        "시리얼번호", "위도", "경도",
        "모델분류", "권역", "사이트키",
        "계약상태", "잔여일수",
        "누적사용량_base", "누적사용량_daily",
        "월사용량_전체", "월사용량_최신", "일사용량_최신",
        "운영개월수",
    ]


    avail_cols = [c for c in keep_cols if c in df.columns]
    df_out = df[avail_cols].copy()
    df_out = df_out.where(pd.notnull(df_out), None)

    # ── 15) JSON 저장 ─────────────────────────────────
    def to_serializable(obj):
        if isinstance(obj, np.integer):  return int(obj)
        if isinstance(obj, np.floating): return None if np.isnan(obj) else float(obj)
        if isinstance(obj, pd.Timestamp):return obj.strftime("%Y-%m-%d")
        if isinstance(obj, datetime):    return obj.strftime("%Y-%m-%d %H:%M:%S")
        return obj

    output = {
        "meta": {
            "generated_at":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "base_date":      base_date,
            "daily_date":     daily_date,
            "diff_days":      diff_days,
            "total_chargers": len(df_out),
            "total_sites":    int(df["사이트키"].nunique()),
        },
        "chargers": json.loads(df_out.to_json(
            orient="records", force_ascii=False, default_handler=to_serializable
        )),
        "sites": json.loads(site_agg.to_json(
            orient="records", force_ascii=False, default_handler=to_serializable
        )),
        "models": json.loads(model_agg.to_json(
            orient="records", force_ascii=False, default_handler=to_serializable
        )),
        "regions": json.loads(region_agg.to_json(
            orient="records", force_ascii=False, default_handler=to_serializable
        )),
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
