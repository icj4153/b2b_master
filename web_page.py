import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import urllib.request
import json
import ast
import ssl
import certifi
import time
import socket
import urllib.error
from datetime import datetime, timedelta
from pathlib import Path
import altair as alt
import re
from env_utils import get_env
from season_analyzer import SEASON_MAP

# [필독] 무조건 1번 줄에 위치
st.set_page_config(page_title="농수산물 소싱 마스터", layout="wide")

# ---------------------------------------------------------
# 네이버 API 출입증은 .env에서 읽습니다.
# ---------------------------------------------------------
NAVER_CLIENT_ID = get_env("NAVER_CLIENT_ID", "")
NAVER_CLIENT_SECRET = get_env("NAVER_CLIENT_SECRET", "")
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = Path(os.getenv("B2B_OUTPUT_DIR", str(BASE_DIR / "output")))
LOG_DIR = Path(os.getenv("B2B_LOG_DIR", str(BASE_DIR / "logs")))
PRICE_HISTORY_CACHE_FILE = OUTPUT_DIR / "price_history.csv"
LAST_CRAWL_TIMESTAMP_FILE = OUTPUT_DIR / "last_crawl_at.txt"
TREND_API_LAST_EVENTS = {}
TREND_API_CACHE_TTL_SECONDS = 3600
DASHBOARD_REFRESH_INTERVAL_SECONDS = int(os.getenv("B2B_DASHBOARD_REFRESH_SECONDS", "300"))
# ---------------------------------------------------------

GOOGLE_EXPORT_RE = re.compile(
    r"https://docs\.google\.com/spreadsheets/d/([^/]+)/export\?gid=([^#&]+)"
)
INTEGRATED_FILE_RE = re.compile(r"전체품목_통합데이터_(\d{8})\.xlsx$")


def normalize_supplier_name(name):
    return str(name).split("_")[0].strip()


def normalize_weight_value(value):
    if pd.isna(value):
        return value
    text = str(value).strip()
    return re.sub(r'(kg|g)$', lambda match: match.group(1).lower(), text, flags=re.IGNORECASE)


def normalize_grade_from_product_name(row):
    product_name = str(row.get('상품명', ''))
    if any(k in product_name for k in ["쥬스용", "주스용"]):
        return "쥬스용"
    if "실속" in product_name:
        return "가정용"
    return row.get('등급')


def normalize_size_from_product_name(row):
    product_name = str(row.get('상품명', ''))
    if any(k in product_name for k in ["혼합", "랜덤"]):
        return "혼합과"
    return row.get('크기')


SEARCH_HINT_KEYWORDS = ["세척", "사과", "부사"]


def normalize_product_search_text(value):
    if pd.isna(value):
        return ""
    text = str(value).casefold()
    text = re.sub(r"[\s\-_./|,()[\]{}]+", " ", text)
    text = re.sub(r"[^0-9a-z가-힣]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def compact_product_search_text(value):
    return normalize_product_search_text(value).replace(" ", "")


def get_product_search_hints(df=None):
    hints = set(SEARCH_HINT_KEYWORDS)

    if "CATEGORY_KEYWORDS" in globals():
        for keyword_list in CATEGORY_KEYWORDS.values():
            hints.update(keyword_list)

    if df is not None and "메인키워드" in df.columns:
        hints.update(df["메인키워드"].dropna().astype(str).unique())

    normalized_hints = {
        compact_product_search_text(hint)
        for hint in hints
        if len(compact_product_search_text(hint)) >= 2
    }
    return sorted(normalized_hints, key=len, reverse=True)


def decompose_compact_search_query(compact_query, hints):
    matched_terms = []
    covered_indexes = set()

    for hint in hints:
        start = compact_query.find(hint)
        if start == -1:
            continue
        matched_terms.append(hint)
        covered_indexes.update(range(start, start + len(hint)))

    if len(matched_terms) < 2 or len(covered_indexes) != len(compact_query):
        return []
    return matched_terms


def product_name_matches_search(product_name, search_query, search_hints=None):
    normalized_product = normalize_product_search_text(product_name)
    compact_product = normalized_product.replace(" ", "")
    normalized_query = normalize_product_search_text(search_query)
    compact_query = normalized_query.replace(" ", "")

    if not compact_query:
        return False

    if compact_query in compact_product:
        return True

    query_terms = normalized_query.split()
    if len(query_terms) > 1:
        return all(term in compact_product for term in query_terms)

    decomposed_terms = decompose_compact_search_query(compact_query, search_hints or [])
    return bool(decomposed_terms) and all(term in compact_product for term in decomposed_terms)


def filter_products_by_name_search(df, search_query):
    if df.empty or "상품명" not in df.columns:
        return df.iloc[0:0].copy()

    search_hints = get_product_search_hints(df)
    mask = df["상품명"].apply(
        lambda product_name: product_name_matches_search(product_name, search_query, search_hints)
    )
    return df[mask].copy()


def normalize_company_link(url):
    match = GOOGLE_EXPORT_RE.match(str(url))
    if match:
        sheet_id, gid = match.groups()
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit?gid={gid}#gid={gid}"
    return url


@st.cache_data
def load_supplier_links():
    base_dir = Path(__file__).resolve().parent
    source_files = {
        "b2b_admin.txt": "ADMIN_PLUS_COMPANIES",
        "b2b_google_sheet.txt": "GOOGLE_SHEET_COMPANIES",
        "b2b_baljuora.txt": "BALJUORA_COMPANIES",
        "b2b_direct.txt": "DIRECT_DOWNLOAD_COMPANIES",
        "b2b_special.txt": "SPECIAL_SITE_COMPANIES",
    }
    links = {}

    for filename, variable_name in source_files.items():
        candidates = [base_dir / filename, base_dir / "b2b_list" / filename]
        source_path = next((path for path in candidates if path.exists()), None)
        if source_path is None:
            continue

        tree = ast.parse(source_path.read_text(encoding="utf-8"), filename=str(source_path))
        for node in tree.body:
            if not isinstance(node, ast.Assign):
                continue
            if not any(isinstance(t, ast.Name) and t.id == variable_name for t in node.targets):
                continue
            try:
                companies = ast.literal_eval(node.value)
            except (ValueError, SyntaxError):
                break

            for company in companies:
                name = company.get("name")
                url = company.get("link_url") or company.get("homepage_url") or company.get("url")
                if name and url:
                    links[normalize_supplier_name(name)] = normalize_company_link(url)
            break

    return links


# 1. 데이터 로드 함수
def find_data_file():
    if OUTPUT_DIR.exists():
        output_files = list(OUTPUT_DIR.glob("전체품목_통합데이터_*.xlsx"))
        if output_files:
            return max(output_files, key=lambda path: path.stat().st_mtime)

    legacy_candidates = [
        BASE_DIR / filename
        for filename in [
            "통합_단가표_제철분석.xlsx",
            "통합_단가표.xlsx",
        ]
    ]

    existing_files = [path for path in legacy_candidates if path.exists()]
    if not existing_files:
        return None
    return max(existing_files, key=lambda path: path.stat().st_mtime)


def get_latest_crawl_label():
    if LAST_CRAWL_TIMESTAMP_FILE.exists():
        timestamp_text = LAST_CRAWL_TIMESTAMP_FILE.read_text(encoding="utf-8").strip()
        crawled_at = pd.to_datetime(timestamp_text, errors="coerce")
        if not pd.isna(crawled_at):
            return f"v{crawled_at.strftime('%Y-%m-%d %H시')}"

    file_path = find_data_file()
    if not file_path:
        return "v데이터 없음"

    crawled_at = datetime.fromtimestamp(file_path.stat().st_mtime)
    return f"v{crawled_at.strftime('%Y-%m-%d %H시')}"


def enable_periodic_dashboard_refresh():
    if DASHBOARD_REFRESH_INTERVAL_SECONDS <= 0:
        return

    components.html(
        f"""
        <script>
        window.setTimeout(function () {{
            window.parent.location.reload();
        }}, {DASHBOARD_REFRESH_INTERVAL_SECONDS * 1000});
        </script>
        """,
        height=0,
        width=0,
    )


@st.cache_data
def load_data(file_path_text, file_mtime_ns, file_size):
    file_path = Path(file_path_text) if file_path_text else None
    if file_path:
        df = pd.read_excel(file_path, dtype=str)
        if {'등급', '상품명'}.issubset(df.columns):
            df['등급'] = df.apply(normalize_grade_from_product_name, axis=1)
        if {'크기', '상품명'}.issubset(df.columns):
            df['크기'] = df.apply(normalize_size_from_product_name, axis=1)
        if '공급가' in df.columns:
            df['공급가'] = pd.to_numeric(df['공급가'], errors='coerce').fillna(0).astype(int)
        return df
    return pd.DataFrame()


def load_latest_data():
    file_path = find_data_file()
    if not file_path:
        return pd.DataFrame()

    file_stat = file_path.stat()
    return load_data(str(file_path), file_stat.st_mtime_ns, file_stat.st_size)


def parse_data_file_date(file_path):
    match = INTEGRATED_FILE_RE.match(file_path.name)
    if not match:
        return None
    return pd.to_datetime(match.group(1), format="%Y%m%d", errors="coerce")


def get_data_file_infos():
    if not OUTPUT_DIR.exists():
        return tuple()

    file_infos = []
    for file_path in sorted(OUTPUT_DIR.glob("전체품목_통합데이터_*.xlsx")):
        file_date = parse_data_file_date(file_path)
        if pd.isna(file_date):
            continue
        file_stat = file_path.stat()
        file_infos.append((str(file_path), file_date.strftime("%Y-%m-%d"), file_stat.st_mtime_ns, file_stat.st_size))
    return tuple(file_infos)


def get_price_history_cache_info():
    if not PRICE_HISTORY_CACHE_FILE.exists():
        return None

    file_stat = PRICE_HISTORY_CACHE_FILE.stat()
    return (str(PRICE_HISTORY_CACHE_FILE), file_stat.st_mtime_ns, file_stat.st_size)


@st.cache_data
def load_price_history_cache(cache_info):
    cache_path_text, _file_mtime_ns, _file_size = cache_info
    history_df = pd.read_csv(cache_path_text, dtype={"상품명": str, "공급사": str})
    required_columns = {"날짜", "상품명", "공급사", "공급가"}
    if not required_columns.issubset(history_df.columns):
        return pd.DataFrame(columns=["날짜", "상품명", "공급사", "공급가"])

    history_df = history_df[["날짜", "상품명", "공급사", "공급가"]].copy()
    history_df["날짜"] = pd.to_datetime(history_df["날짜"], errors="coerce")
    history_df["공급가"] = pd.to_numeric(history_df["공급가"], errors="coerce")
    history_df = history_df.dropna(subset=["날짜", "상품명", "공급가"])
    history_df["공급사"] = history_df["공급사"].fillna("공급사 미표기").astype(str)
    return history_df.sort_values(["날짜", "공급가", "공급사"])


@st.cache_data
def load_price_history_index(file_infos):
    history_frames = []
    price_columns = ["상품명", "공급사", "공급가"]
    required_columns = {"상품명", "공급가"}

    for file_path_text, file_date_text, _file_mtime_ns, _file_size in file_infos:
        try:
            daily_df = pd.read_excel(
                file_path_text,
                dtype=str,
                usecols=lambda column: column in price_columns,
            )
        except Exception as exc:
            print(f"[price-history] read-failed path={file_path_text!r} error={type(exc).__name__}: {exc}")
            continue

        if not required_columns.issubset(daily_df.columns):
            continue

        daily_df = daily_df.dropna(subset=["상품명"]).copy()
        daily_df["날짜"] = pd.to_datetime(file_date_text)
        daily_df["상품명"] = daily_df["상품명"].astype(str)
        daily_df["공급가"] = pd.to_numeric(daily_df["공급가"], errors="coerce")
        daily_df = daily_df.dropna(subset=["공급가"])
        if daily_df.empty:
            continue

        if "공급사" not in daily_df.columns:
            daily_df["공급사"] = "공급사 미표기"
        daily_df["공급사"] = daily_df["공급사"].fillna("공급사 미표기").astype(str)

        history_frames.append(daily_df[["날짜", "상품명", "공급사", "공급가"]])

    if not history_frames:
        return pd.DataFrame(columns=["날짜", "상품명", "공급사", "공급가"])

    return pd.concat(history_frames, ignore_index=True).sort_values(["날짜", "공급가", "공급사"])


def render_product_price_history(product_name):
    file_infos = get_data_file_infos()
    if not file_infos:
        st.info("가격 추적에 사용할 일별 통합 데이터 파일이 없습니다.")
        return

    with st.spinner("가격 이력을 준비하는 중입니다..."):
        cache_info = get_price_history_cache_info()
        if cache_info:
            price_history_index = load_price_history_cache(cache_info)
        else:
            price_history_index = load_price_history_index(file_infos)
    history_df = price_history_index[price_history_index["상품명"] == str(product_name)].copy()
    if history_df.empty:
        st.info("선택한 상품의 과거 가격 데이터가 아직 없습니다.")
        return

    daily_summary = (
        history_df.groupby("날짜", as_index=False)
        .agg(
            최저가=("공급가", "min"),
            평균가=("공급가", "mean"),
            최고가=("공급가", "max"),
            등록수=("공급가", "count"),
        )
        .sort_values("날짜")
    )
    daily_summary["평균가"] = daily_summary["평균가"].round(0)
    chart_df = daily_summary.melt(
        id_vars=["날짜", "등록수"],
        value_vars=["최저가", "평균가", "최고가"],
        var_name="가격구분",
        value_name="공급가",
    )

    st.write(f"### 가격 추적: {product_name}")
    metric_cols = st.columns(4)
    latest_row = daily_summary.iloc[-1]
    first_row = daily_summary.iloc[0]
    price_delta = int(latest_row["최저가"] - first_row["최저가"])
    metric_cols[0].metric("최근 최저가", f"{int(latest_row['최저가']):,} 원", f"{price_delta:,} 원")
    metric_cols[1].metric("최근 평균가", f"{int(latest_row['평균가']):,} 원")
    metric_cols[2].metric("최근 최고가", f"{int(latest_row['최고가']):,} 원")
    metric_cols[3].metric("추적 일수", f"{len(daily_summary):,} 일")

    nearest = alt.selection_point(nearest=True, on="pointerover", fields=["날짜"], empty=False)
    line = alt.Chart(chart_df).mark_line(strokeWidth=3).encode(
        x=alt.X("날짜:T", title="", axis=alt.Axis(format="%m/%d", labelAngle=0)),
        y=alt.Y("공급가:Q", title="공급가", axis=alt.Axis(format=",")),
        color=alt.Color("가격구분:N", title="가격"),
    )
    selectors = alt.Chart(chart_df).mark_point().encode(
        x="날짜:T",
        opacity=alt.value(0),
    ).add_params(nearest)
    points = line.mark_point(size=80, filled=True).encode(
        opacity=alt.condition(nearest, alt.value(1), alt.value(0)),
        tooltip=[
            alt.Tooltip("날짜:T", title="날짜", format="%Y-%m-%d"),
            alt.Tooltip("가격구분:N", title="가격구분"),
            alt.Tooltip("공급가:Q", title="공급가", format=","),
            alt.Tooltip("등록수:Q", title="등록수"),
        ],
    )
    rules = alt.Chart(chart_df).mark_rule(color="#D3D3D3", strokeDash=[4, 4]).encode(
        x="날짜:T"
    ).transform_filter(nearest)
    st.altair_chart(alt.layer(line, selectors, points, rules).properties(height=300), width="stretch")

    if history_df["공급사"].nunique() > 1:
        with st.expander("공급사별 가격 추이"):
            supplier_chart = alt.Chart(history_df).mark_line(point=True).encode(
                x=alt.X("날짜:T", title="", axis=alt.Axis(format="%m/%d", labelAngle=0)),
                y=alt.Y("공급가:Q", title="공급가", axis=alt.Axis(format=",")),
                color=alt.Color("공급사:N", title="공급사"),
                tooltip=[
                    alt.Tooltip("날짜:T", title="날짜", format="%Y-%m-%d"),
                    alt.Tooltip("공급사:N", title="공급사"),
                    alt.Tooltip("공급가:Q", title="공급가", format=","),
                ],
            ).properties(height=260)
            st.altair_chart(supplier_chart, width="stretch")

    with st.expander("일별 가격 데이터"):
        st.dataframe(
            history_df.sort_values(["날짜", "공급가"], ascending=[False, True]),
            width="stretch",
            height=260,
            column_config={
                "날짜": st.column_config.DateColumn(format="YYYY-MM-DD"),
                "공급가": st.column_config.NumberColumn(format="%d 원"),
            },
        )


def parse_season_months(season_str):
    nums = [int(m) for m in re.findall(r'\d+', str(season_str))]
    if len(nums) >= 2 and '~' in str(season_str):
        start, end = nums[0], nums[1]
        if start <= end:
            return list(range(start, end + 1))
        return list(range(start, 13)) + list(range(1, end + 1))
    return nums


def record_trend_api_event(keyword, status, elapsed, **details):
    event = {
        "logged_at": datetime.now().isoformat(timespec="seconds"),
        "logged_at_ts": time.time(),
        "keyword": keyword,
        "status": status,
        "elapsed": round(elapsed, 2),
        **details,
    }
    TREND_API_LAST_EVENTS[keyword] = event

    message_parts = [
        f"[trend] {status}",
        f"keyword={keyword!r}",
        f"elapsed={elapsed:.2f}s",
    ]
    for key, value in details.items():
        message_parts.append(f"{key}={value!r}")
    message = " ".join(message_parts)
    print(message)

    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_path = LOG_DIR / f"trend_api_{datetime.now().strftime('%Y%m%d')}.log"
        with log_path.open("a", encoding="utf-8") as log_file:
            log_file.write(json.dumps(event, ensure_ascii=False) + "\n")
    except Exception as exc:
        print(f"[trend-log] write-failed error={type(exc).__name__}: {exc}")

    return event

def get_trend_api_success_cache():
    return st.session_state.setdefault("trend_api_success_cache", {})


def fetch_trend_data_api(keyword, attempt=1):
    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        record_trend_api_event(keyword, "skip", 0, reason="missing NAVER API credentials")
        return None
    start_time = time.time()
    try:
        url = "https://openapi.naver.com/v1/datalab/search"
        end_date = datetime.now()
        start_date = end_date - timedelta(days=365)
        body = {
            "startDate": start_date.strftime("%Y-%m-%d"),
            "endDate": end_date.strftime("%Y-%m-%d"),
            "timeUnit": "date",
            "keywordGroups": [{"groupName": keyword, "keywords": [keyword]}]
        }
        request = urllib.request.Request(url)
        request.add_header("X-Naver-Client-Id", NAVER_CLIENT_ID)
        request.add_header("X-Naver-Client-Secret", NAVER_CLIENT_SECRET)
        request.add_header("Content-Type", "application/json")
        ssl_context = ssl.create_default_context(cafile=certifi.where())
        response = urllib.request.urlopen(
            request,
            data=json.dumps(body).encode("utf-8"),
            context=ssl_context,
            timeout=5,
        )
        elapsed = time.time() - start_time
        if response.getcode() == 200:
            data = json.loads(response.read())
            rows = len(data['results'][0]['data'])
            if not rows:
                record_trend_api_event(keyword, "no-data", elapsed, attempt=attempt, rows=0)
                return None
            trend_df = pd.DataFrame(data['results'][0]['data'])
            trend_df['날짜'] = pd.to_datetime(trend_df['period'])
            trend_df.set_index('날짜', inplace=True)
            trend_df.rename(columns={'ratio': keyword}, inplace=True)
            trend_df.drop(columns=['period'], inplace=True)
            record_trend_api_event(keyword, "success", elapsed, attempt=attempt, rows=rows)
            return trend_df
        record_trend_api_event(keyword, "http-status", elapsed, attempt=attempt, http_status=response.getcode())
        return None
    except socket.timeout:
        record_trend_api_event(keyword, "timeout", time.time() - start_time, attempt=attempt)
        return None
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:300]
        record_trend_api_event(keyword, "http-error", time.time() - start_time, attempt=attempt, http_status=exc.code, body=details)
        return None
    except urllib.error.URLError as exc:
        record_trend_api_event(keyword, "url-error", time.time() - start_time, attempt=attempt, reason=str(exc.reason))
        return None
    except Exception as exc:
        record_trend_api_event(keyword, "error", time.time() - start_time, attempt=attempt, error=f"{type(exc).__name__}: {exc}")
        return None


def load_trend_data_api(keyword):
    cache = get_trend_api_success_cache()
    cached = cache.get(keyword)
    now = time.time()
    if cached and now - cached["cached_at"] < TREND_API_CACHE_TTL_SECONDS:
        trend_df = cached["data"].copy()
        record_trend_api_event(keyword, "cache-hit", 0, rows=len(trend_df))
        return trend_df
    if cached:
        cache.pop(keyword, None)

    transient_statuses = {"timeout", "url-error", "error"}
    for attempt in range(1, 3):
        trend_df = fetch_trend_data_api(keyword, attempt=attempt)
        event = TREND_API_LAST_EVENTS.get(keyword, {})
        if trend_df is not None:
            cache[keyword] = {"cached_at": now, "data": trend_df.copy()}
            return trend_df
        if event.get("status") not in transient_statuses:
            return None
        time.sleep(0.2)
    return None


# 사장님 비급 알고리즘
def analyze_graph_pattern(trend_df, keyword):
    if trend_df is None or len(trend_df) < 30: return "데이터 부족", "판단 불가"
    data = trend_df[keyword]
    max_val, min_val = data.max(), data.min()
    if max_val == 0: return "검색량 없음", "판단 불가"

    if min_val >= max_val * 0.15: 
        pattern, p_type = "👽 이상한놈 (시즌 미정/꾸준함)", "이상한놈"
    else:
        active_days = len(data[data >= max_val * 0.25])
        if active_days <= 60: pattern, p_type = "🚀 홀쭉이 (1~2달 번쩍!)", "홀쭉이"
        elif active_days <= 100: pattern, p_type = "🐻 퉁퉁이 (2~3달 일반시즌)", "퉁퉁이"
        else: pattern, p_type = "🐘 완전퉁퉁이 (3개월 이상 장기)", "완전퉁퉁이"

    recent_7_avg = data.iloc[-7:].mean()
    prev_7_avg = data.iloc[-14:-7].mean()
    trend_slope = recent_7_avg - prev_7_avg
    is_rising, is_falling = trend_slope > 0, trend_slope < 0
    status = "관망 필요 (타이밍 아님)"

    if p_type == "홀쭉이":
        if is_rising and recent_7_avg >= prev_7_avg * 1.15 and recent_7_avg < max_val * 0.2:
            status = "✨ [골든타임] 상승세 발가락 구간!"
        elif is_falling and recent_7_avg > max_val * 0.2:
            status = "✨ [골든타임] 하락세 전환 (뒷대가리)!"
    elif p_type == "퉁퉁이":
        if is_rising and (max_val * 0.6 <= recent_7_avg <= max_val * 0.85):
            status = "✨ [골든타임] 수요 절정 1~2주 전!"
        elif is_falling and recent_7_avg > max_val * 0.3:
            status = "✨ [골든타임] 하락세 접어듦!"
    elif p_type == "완전퉁퉁이":
        if is_falling: status = "✨ [골든타임] 수요 절정 후 하락세!"
    elif p_type == "이상한놈":
        status = "🔍 쿠팡 직접 분석 필수"

    if "골든타임" not in status and "쿠팡" not in status:
        if recent_7_avg > max_val * 0.85: status = "🔥 수요 절정기"
        elif is_rising: status = "📈 완만한 상승 중"
        else: status = "📉 비시즌 바닥 구간"
    return pattern, status


CATEGORY_ORDER = ["국내과일", "수입과일", "수산물", "야채", "축산", "김치/기타"]
CATEGORY_STYLES = {
    "국내과일": {"bg": "#F7C9C6", "border": "#E88A84", "text": "#7C2520"},
    "수입과일": {"bg": "#F9D8A8", "border": "#E4A449", "text": "#724A12"},
    "수산물": {"bg": "#C8DDF4", "border": "#7FA8D8", "text": "#1E4F82"},
    "야채": {"bg": "#CFE6C8", "border": "#8FBE80", "text": "#2E6532"},
    "축산": {"bg": "#F3CFDA", "border": "#D98BA6", "text": "#7A2943"},
    "김치/기타": {"bg": "#F1D8BC", "border": "#D9A870", "text": "#6E421B"},
}
CATEGORY_SLUGS = {
    "국내과일": "domestic-fruit",
    "수입과일": "imported-fruit",
    "수산물": "seafood",
    "야채": "vegetable",
    "축산": "meat",
    "김치/기타": "etc",
}


def build_category_keywords_from_season_map():
    grouped = {
        "국내과일": [],
        "수입과일": [],
        "수산물": [],
        "야채": [],
    }
    current_category = None
    key_pattern = re.compile(r"^\s*['\"]([^'\"]+)['\"]\s*:")

    for line in (BASE_DIR / "season_analyzer.py").read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if stripped.startswith("#"):
            category_name = stripped.removeprefix("#").strip()
            if category_name in grouped:
                current_category = category_name
            continue

        match = key_pattern.match(line)
        if current_category and match:
            keyword = match.group(1)
            if keyword in SEASON_MAP:
                grouped[current_category].append(keyword)

    return grouped


SEASON_CATEGORY_KEYWORDS = build_category_keywords_from_season_map()
CATEGORY_KEYWORDS = {
    "국내과일": SEASON_CATEGORY_KEYWORDS["국내과일"],
    "수입과일": SEASON_CATEGORY_KEYWORDS["수입과일"],
    "수산물": SEASON_CATEGORY_KEYWORDS["수산물"],
    "야채": SEASON_CATEGORY_KEYWORDS["야채"],
    "축산": [
        "한우", "소고기", "돼지고기", "닭고기", "오리고기", "계란", "달걀", "갈비",
        "삼겹살", "목살", "사골", "우족", "양념육",
    ],
    "김치/기타": ["김치", "기타/상시"],
}
KEYWORD_CATEGORY_MAP = {
    keyword: category
    for category, keyword_list in CATEGORY_KEYWORDS.items()
    for keyword in keyword_list
}


def get_keyword_category(keyword):
    keyword = str(keyword)
    if keyword in KEYWORD_CATEGORY_MAP:
        return KEYWORD_CATEGORY_MAP[keyword]

    for category, keyword_list in CATEGORY_KEYWORDS.items():
        if any(base_keyword in keyword for base_keyword in keyword_list):
            return category
    return "김치/기타"


def group_keywords_by_category(keywords):
    grouped = {category: [] for category in CATEGORY_ORDER}
    for keyword in keywords:
        grouped[get_keyword_category(keyword)].append(keyword)
    return grouped


def render_keyword_category_buttons(keywords):
    st.write("#### 카테고리별 제철 상품")
    grouped_keywords = group_keywords_by_category(keywords)

    if st.session_state.get("selected_keyword_category") not in CATEGORY_ORDER:
        st.session_state.selected_keyword_category = next(
            (category for category in CATEGORY_ORDER if grouped_keywords[category]),
            CATEGORY_ORDER[0],
        )

    selected_category = st.session_state.selected_keyword_category
    base_selectors = []
    selected_rules = []
    for category in CATEGORY_ORDER:
        selector = f".st-key-category-{CATEGORY_SLUGS[category]} button"
        base_selectors.append(selector)
        if selected_category != category:
            continue
        style = CATEGORY_STYLES[category]
        selected_rules.append(
            f"""
            {selector} {{
                background-color: {style["bg"]} !important;
                border-color: {style["border"]} !important;
                color: {style["text"]} !important;
                font-weight: 700 !important;
            }}
            {selector} * {{
                color: {style["text"]} !important;
                font-weight: 700 !important;
            }}
            """
        )

    st.markdown(
        f"""
        <style>
        {", ".join(base_selectors)} {{
            min-height: 38px;
            border-radius: 6px !important;
            transition: background-color 120ms ease, border-color 120ms ease, color 120ms ease;
        }}
        {"".join(selected_rules)}
        </style>
        """,
        unsafe_allow_html=True,
    )

    category_cols = st.columns(len(CATEGORY_ORDER))
    for category, col in zip(CATEGORY_ORDER, category_cols):
        count = len(grouped_keywords[category])
        with col.container(key=f"category-{CATEGORY_SLUGS[category]}"):
            if st.button(f"{category} ({count})", width='stretch', key=f"category_button_{CATEGORY_SLUGS[category]}"):
                st.session_state.selected_keyword_category = category
                st.rerun()

    selected_keywords = grouped_keywords[selected_category]
    st.caption(f"현재 선택: {selected_category} · 상품 {len(selected_keywords)}개")

    if not selected_keywords:
        st.info("선택한 카테고리에 해당하는 제철 상품이 없습니다.")
        return

    item_cols = st.columns(6)
    for i, keyword in enumerate(selected_keywords):
        if item_cols[i % 6].button(keyword, width='stretch', key=f"grid_{selected_category}_{keyword}"):
            st.session_state.clicked_kw = keyword


# ---------------------------------------------------------
# 📊 페이지 1: 실전 소싱 분석 (오리지널 복원)
# ---------------------------------------------------------
def render_analysis_page(df):
    st.title(f"농수산물 AI 소싱 마스터 {get_latest_crawl_label()}")
    st.header("📅 월간 제철 캘린더 & 🤖 맞춤 전략 발굴기")
    supplier_links = load_supplier_links()
    
    selected_month = st.select_slider(
        "소싱 달 선택", options=[f"{i}월" for i in range(1, 13)],
        value=f"{pd.Timestamp.now().month}월"
    )

    month_val = int(selected_month.replace("월", ""))
    seasonal_df = (
        df[df['제철(월)'].apply(lambda value: month_val in parse_season_months(value))]
        if '제철(월)' in df.columns
        else pd.DataFrame()
    )
    keywords = sorted(seasonal_df['메인키워드'].unique()) if not seasonal_df.empty else []
    if "기타/상시" in keywords: keywords.remove("기타/상시")

    if 'clicked_kw' not in st.session_state: st.session_state.clicked_kw = None

    if keywords:
        if st.button(f"✨ '{selected_month}' 사장님 전략 상품 자동 발굴", width='stretch', type="primary"):
            st.session_state.recommendations = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            for i, kw in enumerate(keywords):
                status_text.text(f"분석 중: {kw}...")
                t_data = load_trend_data_api(kw)
                if t_data is not None:
                    pattern, status = analyze_graph_pattern(t_data, kw)
                    if "골든타임" in status or "쿠팡" in status:
                        st.session_state.recommendations.append({"keyword": kw, "pattern": pattern, "status": status})
                progress_bar.progress((i + 1) / len(keywords))
            status_text.empty()
            progress_bar.empty()
            
            if st.session_state.recommendations:
                st.success(f"🎉 전략 상품 {len(st.session_state.recommendations)}개를 발굴했습니다! (클릭 시 아래에 즉시 분석표가 뜹니다)")
                for rec in st.session_state.recommendations:
                    btn_text = f"🎯 {rec['keyword']} - {rec['pattern'].split(' ')[0]} ({rec['status'].replace('✨ [골든타임] ', '')})"
                    # [핵심] 이 버튼을 누르면 세션에 저장되어 즉시 아래 상세 분석이 렌더링 됩니다.
                    if st.button(btn_text, key=f"rec_{rec['keyword']}", width='stretch'):
                        st.session_state.clicked_kw = rec['keyword']
            else: st.info("현재 기준에 들어온 상품이 없습니다.")

    st.divider()
    if keywords:
        render_keyword_category_buttons(keywords)
    st.divider()

    # --- 여기서부터 사장님의 '진짜 원본' 핵심 필터 및 계산기 복원 ---
    st.sidebar.header("🔍 검색 및 정밀 필터")
    search_query = st.sidebar.text_input("🍎 상품명 직접 검색", placeholder="예: 신비복숭아")
    suppliers = sorted(df['공급사'].unique()) if not df.empty else []
    selected_suppliers = st.sidebar.multiselect("🏢 공급사 선택", suppliers)

    target_keyword = search_query if search_query else st.session_state.clicked_kw

    if target_keyword:
        if search_query:
            base_df = filter_products_by_name_search(df, search_query)
            view_title = f"'{search_query}' 검색 결과"
        else:
            base_df = df[df['메인키워드'] == st.session_state.clicked_kw].copy()
            view_title = f"'{st.session_state.clicked_kw}' 제철 소싱 리스트"

        trend_data = load_trend_data_api(target_keyword)
        if trend_data is not None:
            pattern, status = analyze_graph_pattern(trend_data, target_keyword)
            st.write(f"### 📈 '{target_keyword}' 타겟팅 분석")
            m1, m2 = st.columns(2)
            m1.info(f"**체형 판독:** {pattern}")
            if "골든타임" in status: m2.success(f"**진입 타이밍:** {status}")
            elif "쿠팡" in status: m2.warning(f"**진입 타이밍:** {status}")
            else: m2.info(f"**진입 타이밍:** {status}")
            
            # 고급 인터랙티브 그래프 (원본 복원)
            chart_df = trend_data.reset_index()
            x_axis = alt.X('날짜:T', axis=alt.Axis(format='%Y.%-m월', title='', tickCount='month', labelAngle=0))
            nearest = alt.selection_point(nearest=True, on='pointerover', fields=['날짜'], empty=False)
            line = alt.Chart(chart_df).mark_line(color='#FF4B4B', strokeWidth=3).encode(x=x_axis, y=alt.Y(f'{target_keyword}:Q', axis=alt.Axis(title='검색 지수')))
            selectors = alt.Chart(chart_df).mark_point().encode(x='날짜:T', opacity=alt.value(0)).add_params(nearest)
            points = line.mark_point(size=90, color='#FF4B4B', fill='white').encode(
                opacity=alt.condition(nearest, alt.value(1), alt.value(0)),
                tooltip=[alt.Tooltip('날짜:T', title='날짜', format='%Y년 %m월 %d일'), alt.Tooltip(f'{target_keyword}:Q', title='검색량')]
            )
            rules = alt.Chart(chart_df).mark_rule(color='#D3D3D3', strokeDash=[4,4]).encode(x='날짜:T').transform_filter(nearest)
            st.altair_chart(alt.layer(line, selectors, points, rules).properties(height=250), width='stretch')
            st.divider()

        if not base_df.empty:
            if selected_suppliers: base_df = base_df[base_df['공급사'].isin(selected_suppliers)]
            base_df['중량'] = base_df['중량'].apply(normalize_weight_value)
            
            # 등급/크기/중량 등 정밀 필터링 (원본 완벽 복원)
            grades = sorted([x for x in base_df['등급'].unique() if str(x) != 'nan'])
            sel_grades = st.sidebar.multiselect("⭐ 등급 선택", grades, default=grades)
            sizes = sorted([x for x in base_df['크기'].unique() if str(x) != 'nan'])
            sel_sizes = st.sidebar.multiselect("📏 크기 선택", sizes, default=sizes)
            weights = sorted([x for x in base_df['중량'].unique() if str(x) != 'nan'])
            sel_weights = st.sidebar.multiselect("⚖️ 중량 선택", weights, default=weights)

            st.sidebar.divider()
            # 마진율 계산기 (원본 완벽 복원)
            target_price = st.sidebar.number_input("판매가 (원)", min_value=0, value=30000)
            market_fee = st.sidebar.selectbox("마켓", ["쿠팡(11.66%)", "토스(10.56%)", "토스광고(1.76%)"])
            fee_rate = 11.66 if "쿠팡" in market_fee else (10.56 if "일반" in market_fee else 1.76)

            final_df = base_df[
                (base_df['등급'].isin(sel_grades) | base_df['등급'].isna()) &
                (base_df['크기'].isin(sel_sizes) | base_df['크기'].isna()) &
                (base_df['중량'].isin(sel_weights) | base_df['중량'].isna())
            ].copy()

            if target_price > 0:
                fee_amount = target_price * (fee_rate / 100)
                final_df['순수익'] = (target_price - final_df['공급가'] - fee_amount).astype(int)
                final_df['마진율'] = ((final_df['순수익'] / target_price) * 100).round(1)

            final_df['공급사 홈페이지'] = final_df['공급사'].map(
                lambda name: supplier_links.get(normalize_supplier_name(name), "")
            )
            display_columns = ['등급','크기','중량','과수','공급가','공급사','공급사 홈페이지','상품명']
            if '순수익' in final_df.columns:
                display_columns.extend(['순수익','마진율'])
            display_columns = [c for c in display_columns if c in final_df.columns]

            st.write(f"### 📋 {view_title} (총 {len(final_df):,}개)")
            sorted_final_df = final_df.sort_values(by='공급가')
            product_options = sorted(sorted_final_df['상품명'].dropna().astype(str).unique())
            table_df = sorted_final_df[display_columns].reset_index(drop=True)
            selected_product_name = None
            table_event = st.dataframe(table_df, width='stretch', height=400,
                column_config={"공급가": st.column_config.NumberColumn(format="%d 원"),
                               "순수익": st.column_config.NumberColumn(format="%d 원"),
                               "마진율": st.column_config.NumberColumn(format="%.1f %%"),
                               "공급사 홈페이지": st.column_config.LinkColumn(
                                   "공급사 홈페이지",
                                   display_text="열기"
                               )},
                on_select="rerun",
                selection_mode="single-row")
            selected_rows = table_event.selection.rows
            if selected_rows and "상품명" in table_df.columns:
                selected_product_name = str(table_df.iloc[selected_rows[0]]["상품명"])
            elif product_options:
                selected_product_name = st.selectbox(
                    "가격 그래프를 볼 상품명",
                    product_options,
                    index=None,
                    key=f"price_history_product_{target_keyword}",
                    placeholder="상품명을 선택하세요",
                )
            if selected_product_name:
                st.divider()
                render_product_price_history(selected_product_name)
    else: st.info("👆 상단에서 상품을 선택하거나 검색해 주세요.")

# ---------------------------------------------------------
# 📅 페이지 2: 연간 제철 로드맵 (오리지널 깔끔 버전)
# ---------------------------------------------------------
def render_calendar_page(df):
    st.title("🗓️ 연간 농수산물 제철 로드맵")
    st.write("사장님의 제철 사전에 등록된 상품들의 연간 흐름을 한눈에 확인하세요.")

    if df.empty or '메인키워드' not in df.columns or '제철(월)' not in df.columns:
        st.warning("제철 사전 데이터가 부족합니다.")
        return

    calendar_data = []
    unique_items = df.drop_duplicates(subset=['메인키워드'])

    for _, row in unique_items.iterrows():
        kw = row['메인키워드']
        season_str = str(row['제철(월)'])
        months = parse_season_months(season_str)
        
        for m in months:
            calendar_data.append({"상품명": kw, "월": f"{m}월", "status": 1})

    cal_df = pd.DataFrame(calendar_data)
    if not cal_df.empty:
        month_order = [f"{i}월" for i in range(1, 13)]
        chart = alt.Chart(cal_df).mark_rect(cornerRadius=3).encode(
            x=alt.X('월:O', sort=month_order, title=None),
            y=alt.Y('상품명:N', title=None, sort='ascending'),
            color=alt.Color('status:Q', scale=alt.Scale(range=['#FF4B4B', '#FF4B4B']), legend=None),
            tooltip=['상품명', '월']
        ).properties(height=alt.Step(25))

        st.altair_chart(chart, width='stretch')
    else: st.info("로드맵을 그릴 수 있는 월 데이터가 엑셀에 없습니다.")

# ---------------------------------------------------------
# 🚩 메인 실행부
# ---------------------------------------------------------
enable_periodic_dashboard_refresh()
df_main = load_latest_data()
with st.sidebar:
    st.title("👨‍🌾 농수산물 지휘소")
    menu = st.radio("🚩 메뉴 선택", ["실전 소싱 분석", "연간 제철 로드맵"])

if menu == "실전 소싱 분석":
    render_analysis_page(df_main)
else:
    render_calendar_page(df_main)
