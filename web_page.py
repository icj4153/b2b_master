import streamlit as st
import pandas as pd
import os
import urllib.request
import json
import ast
from datetime import datetime, timedelta
from pathlib import Path
import altair as alt
import re
from env_utils import get_env

# [필독] 무조건 1번 줄에 위치
st.set_page_config(page_title="제철한가득 소싱 마스터", layout="wide")

# ---------------------------------------------------------
# 네이버 API 출입증은 .env에서 읽습니다.
# ---------------------------------------------------------
NAVER_CLIENT_ID = get_env("NAVER_CLIENT_ID", "")
NAVER_CLIENT_SECRET = get_env("NAVER_CLIENT_SECRET", "")
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = Path(os.getenv("B2B_OUTPUT_DIR", str(BASE_DIR / "output")))
# ---------------------------------------------------------

GOOGLE_EXPORT_RE = re.compile(
    r"https://docs\.google\.com/spreadsheets/d/([^/]+)/export\?gid=([^#&]+)"
)


def normalize_supplier_name(name):
    return str(name).split("_")[0].strip()


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
@st.cache_data
def find_data_file():
    candidates = []
    if OUTPUT_DIR.exists():
        candidates.extend(OUTPUT_DIR.glob("전체품목_통합데이터_*.xlsx"))

    candidates.extend(
        BASE_DIR / filename
        for filename in [
            "통합_단가표_제철분석.xlsx",
            "통합_단가표.xlsx",
        ]
    )

    existing_files = [path for path in candidates if path.exists()]
    if not existing_files:
        return None
    return max(existing_files, key=lambda path: path.stat().st_mtime)


@st.cache_data
def load_data():
    file_path = find_data_file()
    if file_path:
        df = pd.read_excel(file_path, dtype=str)
        if '공급가' in df.columns:
            df['공급가'] = pd.to_numeric(df['공급가'], errors='coerce').fillna(0).astype(int)
        return df
    return pd.DataFrame()

@st.cache_data(ttl=3600)
def load_trend_data_api(keyword):
    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        return None
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
        response = urllib.request.urlopen(request, data=json.dumps(body).encode("utf-8"))
        if response.getcode() == 200:
            data = json.loads(response.read())
            if not data['results'][0]['data']: return None
            trend_df = pd.DataFrame(data['results'][0]['data'])
            trend_df['날짜'] = pd.to_datetime(trend_df['period'])
            trend_df.set_index('날짜', inplace=True)
            trend_df.rename(columns={'ratio': keyword}, inplace=True)
            trend_df.drop(columns=['period'], inplace=True)
            return trend_df
        return None
    except Exception: return None

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

# ---------------------------------------------------------
# 📊 페이지 1: 실전 소싱 분석 (오리지널 복원)
# ---------------------------------------------------------
def render_analysis_page(df):
    st.title("🍎 제철한가득 AI 소싱 마스터 v3.8")
    st.header("📅 월간 제철 캘린더 & 🤖 맞춤 전략 발굴기")
    supplier_links = load_supplier_links()
    
    selected_month = st.select_slider(
        "소싱 달 선택", options=[f"{i}월" for i in range(1, 13)],
        value=f"{pd.Timestamp.now().month}월"
    )

    month_val = selected_month.replace("월", "")
    seasonal_df = df[df['제철(월)'].str.contains(month_val, na=False)] if '제철(월)' in df.columns else pd.DataFrame()
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
        cols = st.columns(6)
        for i, kw in enumerate(keywords):
            if cols[i % 6].button(kw, width='stretch', key=f"grid_{kw}"):
                st.session_state.clicked_kw = kw
    st.divider()

    # --- 여기서부터 사장님의 '진짜 원본' 핵심 필터 및 계산기 복원 ---
    st.sidebar.header("🔍 검색 및 정밀 필터")
    search_query = st.sidebar.text_input("🍎 상품명 직접 검색", placeholder="예: 신비복숭아")
    suppliers = sorted(df['공급사'].unique()) if not df.empty else []
    selected_suppliers = st.sidebar.multiselect("🏢 공급사 선택", suppliers)

    target_keyword = search_query if search_query else st.session_state.clicked_kw

    if target_keyword:
        if search_query:
            base_df = df[df['상품명'].str.contains(search_query, na=False)].copy()
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
            st.dataframe(final_df.sort_values(by='공급가')[display_columns], width='stretch', height=400,
                column_config={"공급가": st.column_config.NumberColumn(format="%d 원"),
                               "순수익": st.column_config.NumberColumn(format="%d 원"),
                               "마진율": st.column_config.NumberColumn(format="%.1f %%"),
                               "공급사 홈페이지": st.column_config.LinkColumn(
                                   "공급사 홈페이지",
                                   display_text="열기"
                               )})
    else: st.info("👆 상단에서 상품을 선택하거나 검색해 주세요.")

# ---------------------------------------------------------
# 📅 페이지 2: 연간 제철 로드맵 (오리지널 깔끔 버전)
# ---------------------------------------------------------
def render_calendar_page(df):
    st.title("🗓️ 연간 농수산물 제철 로드맵")
    st.write("사장님의 제철 사전에 등록된 상품들의 연간 흐름을 한눈에 확인하세요.")

    if df.empty or '제철(월)' not in df.columns:
        st.warning("제철 사전 데이터가 부족합니다.")
        return

    calendar_data = []
    unique_items = df.drop_duplicates(subset=['메인키워드'])

    for _, row in unique_items.iterrows():
        kw = row['메인키워드']
        season_str = str(row['제철(월)'])
        months = []
        if '~' in season_str:
            nums = re.findall(r'\d+', season_str)
            if len(nums) >= 2:
                start, end = int(nums[0]), int(nums[1])
                months = list(range(start, end + 1))
        else:
            months = [int(m) for m in re.findall(r'\d+', season_str)]
        
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
df_main = load_data()
with st.sidebar:
    st.title("👨‍🌾 제철한가득 지휘소")
    menu = st.radio("🚩 메뉴 선택", ["실전 소싱 분석", "연간 제철 로드맵"])

if menu == "실전 소싱 분석":
    render_analysis_page(df_main)
else:
    render_calendar_page(df_main)
