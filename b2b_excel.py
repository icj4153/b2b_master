import asyncio
import os
import re
import pandas as pd
import requests
from datetime import datetime
from openpyxl import load_workbook
from playwright.async_api import async_playwright
from env_utils import get_required_env

# [설정 1] 사장님 정보
USER_ID = get_required_env("B2B_USER_ID")
USER_PW = get_required_env("B2B_USER_PW")

# [설정 2] 어드민플러스 업체 (19곳)
ADMIN_PLUS_COMPANIES = [
    {"name": "팜허브", "url": "https://priceit.adminplus.co.kr/partner/login.html"},
    {"name": "팜플로우", "url": "https://farmflow.adminplus.co.kr/partner/login.html"},
    {"name": "최고집", "url": "https://zain0401.adminplus.co.kr/partner/login.html"},
    {"name": "덤덤몰", "url": "https://dumdummall.adminplus.co.kr/partner/login.html"},
    {"name": "PBF쥬얼리프룻", "url": "https://pbfcompany.adminplus.co.kr/partner/login.html"},
    {"name": "농부를찾아서", "url": "https://ongreen.adminplus.co.kr/partner/login.html"},
    {"name": "산지이음", "url": "https://orangec.adminplus.co.kr/partner/login.html"},
    {"name": "프리어드", "url": "https://rhdnsl2236.adminplus.co.kr/partner/login.html"},
    {"name": "데일리푸드", "url": "https://dailyfood.adminplus.co.kr/partner/login.html"},
    {"name": "신흥유통", "url": "https://sh0527.adminplus.co.kr/partner/login.html"},
    {"name": "산지도감", "url": "https://dogam.adminplus.co.kr/partner/login.html"},
    {"name": "프루틴", "url": "https://fruitin475.adminplus.co.kr/partner/login.html"},
    {"name": "마니팜", "url": "https://mp3462.adminplus.co.kr/partner/login.html"},
    {"name": "팡이농장", "url": "https://jaehwan0330.adminplus.co.kr/partner/login.html"},
    {"name": "제주다팜", "url": "https://kkangta55.adminplus.co.kr/partner/login.html"},
    {"name": "늘푸른우리", "url": "https://hwanggs3.adminplus.co.kr/partner/login.html"},
    {"name": "제철도매", "url": "https://pjh4732.adminplus.co.kr/partner/login.html"},
    {"name": "과일방앗간", "url": "https://cks0644.adminplus.co.kr/partner/login.html"},
    {"name": "초록청년", "url": "https://chfhrcjdsus.adminplus.co.kr/partner/login.html"}
]

# [설정 3] 구글 시트 업체 (직접 다운로드 링크)
GOOGLE_SHEET_COMPANIES = [
    {
        "name": "성공푸드-선물세트",
        "url": "https://docs.google.com/spreadsheets/d/1aRW86p7PIl9wdFoMhPmMWjfxJ2zA74_9nRbgR9j59Rs/export?gid=509215400#gid=509215400"
    },
    {
        "name": "성공푸드-국내과일",
        "url": "https://docs.google.com/spreadsheets/d/1aRW86p7PIl9wdFoMhPmMWjfxJ2zA74_9nRbgR9j59Rs/export?gid=480414977#gid=480414977"
    },
    {
        "name": "성공푸드-수입과일",
        "url": "https://docs.google.com/spreadsheets/d/1aRW86p7PIl9wdFoMhPmMWjfxJ2zA74_9nRbgR9j59Rs/export?gid=667409228#gid=667409228"
    },
    {
        "name": "성공푸드-농산물",
        "url": "https://docs.google.com/spreadsheets/d/1aRW86p7PIl9wdFoMhPmMWjfxJ2zA74_9nRbgR9j59Rs/export?gid=1651579023#gid=1651579023"
    },
    {
        "name": "성공푸드-수산",
        "url": "https://docs.google.com/spreadsheets/d/1aRW86p7PIl9wdFoMhPmMWjfxJ2zA74_9nRbgR9j59Rs/export?gid=192163975#gid=192163975"
    },
    {
        "name": "성공푸드-축산",
        "url": "https://docs.google.com/spreadsheets/d/1aRW86p7PIl9wdFoMhPmMWjfxJ2zA74_9nRbgR9j59Rs/export?gid=1671972104#gid=1671972104"
    },
    {
        "name": "성공푸드-반찬",
        "url": "https://docs.google.com/spreadsheets/d/1aRW86p7PIl9wdFoMhPmMWjfxJ2zA74_9nRbgR9j59Rs/export?gid=2141484404#gid=2141484404"
    },
    {
        "name": "트리프룻-과일",
        "url": "https://docs.google.com/spreadsheets/d/13S3U2jRdu1nTdTygEM6m9nmA7NyD-7DN/export?gid=113727950#gid=113727950"
    },
    {
        "name": "트리프룻-농산",
        "url": "https://docs.google.com/spreadsheets/d/13S3U2jRdu1nTdTygEM6m9nmA7NyD-7DN/export?gid=1740404040#gid=1740404040"
    },
    {
        "name": "트리프룻-김치",
        "url": "https://docs.google.com/spreadsheets/d/13S3U2jRdu1nTdTygEM6m9nmA7NyD-7DN/export?gid=2105255924#gid=2105255924"
    },
    {
        "name": "트리프룻-수산",
        "url": "https://docs.google.com/spreadsheets/d/13S3U2jRdu1nTdTygEM6m9nmA7NyD-7DN/export?gid=2054210809#gid=2054210809"
    },
    {
        "name": "트리프룻-축산",
        "url": "https://docs.google.com/spreadsheets/d/13S3U2jRdu1nTdTygEM6m9nmA7NyD-7DN/export?gid=1090018950#gid=1090018950"
    },
    {
        "name": "트리프룻-선물세트",
        "url": "https://docs.google.com/spreadsheets/d/13S3U2jRdu1nTdTygEM6m9nmA7NyD-7DN/export?gid=1392357560#gid=1392357560"
    },
        {
        "name": "마이프레셔-농산물",
        "url": "https://docs.google.com/spreadsheets/d/1H1-YTi6ljvn7i2LYVnSuexlxwtx3NjY_cFj9d-CN-mA/export?gid=508993936#gid=508993936"
    },
    {
        "name": "마이프레셔-수산물",
        "url": "https://docs.google.com/spreadsheets/d/1H1-YTi6ljvn7i2LYVnSuexlxwtx3NjY_cFj9d-CN-mA/export?gid=673470564#gid=673470564"
    },
    {
        "name": "신선천재김사장-국산",
        "url": "https://docs.google.com/spreadsheets/d/1-QqD2WfJlcq4wAH42EZGgsM835UzxGEOg8MIWL6n6j4/export?gid=0#gid=0"
    },
    {
        "name": "신선천재김사장-수입",
        "url": "https://docs.google.com/spreadsheets/d/1-QqD2WfJlcq4wAH42EZGgsM835UzxGEOg8MIWL6n6j4/export?gid=1234469799#gid=1234469799"
    },
    {
        "name": "신선천재김사장-수산물",
        "url": "https://docs.google.com/spreadsheets/d/1-QqD2WfJlcq4wAH42EZGgsM835UzxGEOg8MIWL6n6j4/export?gid=343827659#gid=343827659"
    },
    {
        "name": "신선천재김사장-김치가공",
        "url": "https://docs.google.com/spreadsheets/d/1-QqD2WfJlcq4wAH42EZGgsM835UzxGEOg8MIWL6n6j4/export?gid=6062566#gid=6062566"
    },
    {
        "name": "신선천재김사장-축산",
        "url": "https://docs.google.com/spreadsheets/d/1-QqD2WfJlcq4wAH42EZGgsM835UzxGEOg8MIWL6n6j4/export?gid=378135222#gid=378135222"
    },
    {
        "name": "산들리에",
        "url": "https://docs.google.com/spreadsheets/d/1R7REm9w0hRHYuHYOW0wnPlUIe3P0vatLzjmjN-jTolY/export?gid=428486313#gid=428486313"
    },
        {
        "name": "팜링크-과일",
        "url": "https://docs.google.com/spreadsheets/d/1BD8hvjA9NvFeNIMIXD1vT53l4qlowhRR2dEsu-0RxAg/export?gid=1778436418#gid=1778436418"
    },
    {
        "name": "팜링크-농산물",
        "url": "https://docs.google.com/spreadsheets/d/1BD8hvjA9NvFeNIMIXD1vT53l4qlowhRR2dEsu-0RxAg/export?gid=1114939458#gid=1114939458"
    },
    {
        "name": "팜링크-축산",
        "url": "https://docs.google.com/spreadsheets/d/1BD8hvjA9NvFeNIMIXD1vT53l4qlowhRR2dEsu-0RxAg/export?gid=1766711354#gid=1766711354"
    },
    {
        "name": "팜링크-선물세트",
        "url": "https://docs.google.com/spreadsheets/d/1BD8hvjA9NvFeNIMIXD1vT53l4qlowhRR2dEsu-0RxAg/export?gid=1509637737#gid=1509637737"
    },
    {
        "name": "모티스템",
        "url": "https://docs.google.com/spreadsheets/d/1W94T7cHWt2ZDsXYmWLixcEDE3qxGGD0-GFg1j9dA1II/export?gid=0#gid=0"
    },
    {
        "name": "트라이365",
        "url": "https://docs.google.com/spreadsheets/d/1R6iviy3KXjEdK4zknX_2NbVTLiYa0yHZx65XUxSXPeM/export?gid=353676437#gid=353676437"
    }
    
]

# [설정 4] 발주오라 업체 리스트
BALJUORA_COMPANIES = [
    {
        "name": "업푸르트", 
        "url": "https://shop.baljuora.com/bluecoro21/signin", # 로그인하는 첫 화면 주소
        "list_url": "https://shop.baljuora.com/bluecoro21/products" # 사장님이 찾으신 다이렉트 주소!
    },
    {
        "name": "우람씨푸드", 
        "url": "https://shop.baljuora.com/leejaehyeok24/signin", # 로그인하는 첫 화면 주소
        "list_url": "https://shop.baljuora.com/leejaehyeok24/products" # 사장님이 찾으신 다이렉트 주소!
    }
]

# [설정 5 ] 대봉유통 다운로드
DIRECT_DOWNLOAD_COMPANIES = [
    {
        "name": "대봉유통", 
        "url": "https://www.daebong.net/product", 
        "btn_text": "전체 데이터 엑셀 다운로드"
    }
]


DOWNLOAD_DIR = "b2b_downloads"
TARGET_BTN_TEXT = "전체제품 엑셀 다운로드"
OUTPUT_FILE = f"전체품목_통합데이터_{datetime.now().strftime('%Y%m%d')}.xlsx"

# --- 1. 유틸리티 함수 (팝업 닫기, 파싱) ---

async def close_all_popups(page):
    popup_selectors = ["text='하루 동안 이 창을 열지 않습니다'", "text='오늘 하루 열지 않기'", "text='닫기'", ".btn-close", "area[alt*='닫기']"]
    for frame in page.frames:
        for sel in popup_selectors:
            try:
                elements = await frame.locator(sel).all()
                for el in elements:
                    if await el.is_visible(): await el.click(); await asyncio.sleep(0.2)
            except: continue

def parse_product_info(row):
    """상품명과 옵션을 결합하여 정밀 파싱 (보자기, 세트 키워드 강화)"""
    name = str(row.get('상품명', '')).replace('nan', '')
    option = str(row.get('상품 옵션', '')).replace('nan', '')
    full_name = f"{name} {option}".strip()
    
    # 1. 등급 추출 (사장님 추가 요청사항 반영)
    grade = "일반"
    if "가정용" in full_name: 
        grade = "가정용"
    # '선물', '세트', '보자기' 중 하나라도 들어가면 '선물세트'로 분류
    elif any(k in full_name for k in ["선물", "세트", "보자기"]): 
        grade = "선물세트"
    elif "정품" in full_name: 
        grade = "정품"
    
    # 2. 크기 추출
    size = "미표기"
    for k in ["혼합과", "소과", "중소과", "중과", "중대과", "대과"]:
        if k in full_name: 
            size = k
            break
            
    # 3. 과수 추출 (개, 입, 알, 과, 구)
    count = "미표기"
    count_match = re.search(r'(\d+[-~]?\d*)([개입알과구])', full_name)
    if count_match: 
        count = f"{count_match.group(1)}과"
    
    # 4. 중량 추출 (kg, g)
    weight_match = re.search(r'(\d+\.?\d*)(kg|g)', full_name, re.IGNORECASE)
    weight = weight_match.group(0) if weight_match else "미표기"
    
    return pd.Series([grade, size, weight, count, full_name])


# --- 어드민 플러스 전용 다운로드 엔진 (다이렉트 URL + 로케이터 수정) ---
async def download_admin_plus(context, company):
    page = await context.new_page()
    print(f"[{company['name']}] 어드민 플러스 접속 및 로그인 시도 중...")
    try:
        # 1. 로그인 페이지 접속
        await page.goto(company['url'], wait_until="load")
        
        # [수정] 사장님이 알려주신대로 type 속성을 사용하여 아이디/비번 입력
        # 첫 번째 text 타입 입력창에 ID를, 첫 번째 password 타입 입력창에 PW를 넣습니다.
        await page.locator("input[type='text']").first.fill(USER_ID)
        await page.locator("input[type='password']").first.fill(USER_PW)
        await page.keyboard.press("Enter")
        
        # 로그인 승인 대기
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(2)
        await close_all_popups(page)

        # 2. 마법의 주소로 다이렉트 점프
        # URL에서 'login.html'을 잘라내고 파라미터 결합
        if '/login.html' in company['url']:
            base_url = company['url'].split('/login.html')[0]
        else:
            # 혹시나 /login.html이 없는 주소일 경우를 대비한 안전 장치
            base_url = company['url'].rstrip('/')
            
        list_url = f"{base_url}/?mod=product&actpage=prt.list"

        print(f"[{company['name']}] 상품리스트로 다이렉트 점프: {list_url}")
        await page.goto(list_url, wait_until="load")

        # 3. 엑셀 다운로드 실행
        excel_btn = page.get_by_text(TARGET_BTN_TEXT)
        async with page.expect_download() as download_info:
            await page.get_by_text(TARGET_BTN_TEXT).click(force=True)
        download = await download_info.value
        file_path = os.path.join(DOWNLOAD_DIR, f"{company['name']}_{datetime.now().strftime('%Y%m%d')}.xlsx")
        await download.save_as(file_path)
        print(f"[{company['name']}] ★ 성공")
    except Exception as e: print(f"[{company['name']}] 실패: {e}")
    finally: await page.close()


# --- 발주오라 전용 다운로드 엔진 (다이렉트 URL 초스피드 버전) ---
async def download_baljuora(context, company):
    page = await context.new_page()
    print(f"[{company['name']}] 발주오라 접속 및 로그인 시도 중...")
    try:
        # 1. 로그인 과정 (기존과 동일)
        await page.goto(company['url'], wait_until="load")
        await page.locator("input[type='text']").first.fill(USER_ID)
        await page.locator("input[type='password']").first.fill(USER_PW)
        await page.keyboard.press("Enter")
        
        # 로그인 승인 대기
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(2)
        await close_all_popups(page)

        print(f"[{company['name']}] 상품리스트로 다이렉트 점프 중! 🚀")
        
        # ★ 2. 메뉴 클릭 없이 사장님이 찾으신 주소로 다이렉트 이동! ★
        await page.goto(company['list_url'], wait_until="load")
        
        # 리스트 화면이 완전히 뜰 때까지 대기
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(2)
        await close_all_popups(page)

        # 3. '엑셀 다운로드' 버튼 찾기 및 낚아채기
        excel_btn = page.get_by_text("엑셀 다운로드")
        if not await excel_btn.is_visible(timeout=3000):
            excel_btn = page.get_by_text("엑셀다운")

        if await excel_btn.is_visible():
            async with page.expect_download() as download_info:
                await excel_btn.click(force=True)
                
            download = await download_info.value
            file_path = os.path.join(DOWNLOAD_DIR, f"{company['name']}_{datetime.now().strftime('%Y%m%d')}.xlsx")
            await download.save_as(file_path)
            print(f"[{company['name']}] ★ 발주오라 다이렉트 다운로드 성공!")
        else:
            print(f"[{company['name']}] 에러: 상품리스트 이동 후 엑셀 다운로드 버튼이 안 보입니다.")
            
    except Exception as e: 
        print(f"[{company['name']}] 처리 중 오류 발생: {e}")
    finally: 
        await page.close()


def download_google_sheet(company):
    print(f"[{company['name']}] 구글 시트 다운로드 중...")
    try:
        response = requests.get(company['url'])
        if response.status_code == 200:
            file_path = os.path.join(DOWNLOAD_DIR, f"{company['name']}_{datetime.now().strftime('%Y%m%d')}.xlsx")
            with open(file_path, 'wb') as f: f.write(response.content)
            print(f"[{company['name']}] ★ 성공")
        else: print(f"[{company['name']}] 실패 (HTTP {response.status_code})")
    except Exception as e: print(f"[{company['name']}] 에러: {e}")

# --- 대봉유통 다운로드 엔진 ---
async def download_direct(context, company):
    page = await context.new_page()
    print(f"[{company['name']}] 접속 중 (로그인 불필요)...")
    try:
        # 1. URL 바로 접속
        await page.goto(company['url'], wait_until="load")
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(2)
        await close_all_popups(page)

        # 2. 버튼 찾기 및 다운로드
        excel_btn = page.get_by_text(company['btn_text'], exact=True)
        
        if await excel_btn.is_visible():
            async with page.expect_download() as download_info:
                await excel_btn.click(force=True)
                
            download = await download_info.value
            file_path = os.path.join(DOWNLOAD_DIR, f"{company['name']}_{datetime.now().strftime('%Y%m%d')}.xlsx")
            await download.save_as(file_path)
            print(f"[{company['name']}] ★ 엑셀 다운로드 성공!")
        else:
            print(f"[{company['name']}] 에러: '{company['btn_text']}' 버튼을 찾을 수 없습니다.")
            
    except Exception as e: 
        print(f"[{company['name']}] 처리 중 오류: {e}")
    finally: 
        await page.close()

# --- 3. 데이터 통합 엔진 (헤더 자동 찾기 기능 포함) ---
r""" def integrate_data():
    all_rows = []
    if not os.path.exists(DOWNLOAD_DIR): 
        print("다운로드 폴더가 없습니다.")
        return

    print("\n[데이터 표준화 및 통합 작업 시작]")
    for filename in os.listdir(DOWNLOAD_DIR):
        if not filename.endswith(".xlsx") or filename.startswith("~$"): continue
        filepath = os.path.join(DOWNLOAD_DIR, filename)
        supplier = filename.split('_')[0]
        
        try:
            wb = load_workbook(filepath, data_only=True)
            excel_file_dict = pd.read_excel(filepath, sheet_name=None, header=None, dtype=str) # 안전하게 모두 문자로 읽기
            
            for sheet_name, df in excel_file_dict.items():
                if df.empty: continue
                
                ws = wb[sheet_name]
                hidden_rows = [idx for idx, dim in ws.row_dimensions.items() if dim.hidden]
                
                visible_indices = [i for i in range(len(df)) if (i + 1) not in hidden_rows]
                df = df.iloc[visible_indices].reset_index(drop=True)
                
                if df.empty: continue

                # 1. 헤더(제목줄) 찾기
                header_idx = -1
                for idx, row in df.iterrows():
                    row_values = [str(cell).replace(" ", "").replace("\n", "") for cell in row.values]
                    if any(len(v) <= 10 and any(k in v for k in ['품명', '상품명', '옵션', '품목', '현재가', '최종가격', '상품가']) for v in row_values):
                        header_idx = idx
                        break
                
                if header_idx == -1: continue
                
                df.columns = df.iloc[header_idx]
                df = df.iloc[header_idx+1:].reset_index(drop=True)
                df = df.loc[:, df.columns.notna()]
                df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

                # 2. 업체별 상품명/옵션명 열 매핑
                name_col = None
                opt_col = next((c for c in df.columns if str(c).replace(" ","") in ["옵션", "옵션명"]), None)
                
                name_keywords = ['상품명', '품명', '품목명', '품목', '발주서상품명']
                for k in name_keywords:
                    target = next((c for c in df.columns if str(c).replace(" ", "") == k), None)
                    if target:
                        name_col = target
                        break

                if not name_col and opt_col:
                    name_col = opt_col
                    opt_col = None 

                if name_col:
                    if opt_col and "트라이365" in supplier:
                        df[name_col] = df[name_col].ffill()
                        df['상품명'] = df[name_col].astype(str) + " " + df[opt_col].astype(str).replace("nan", "")
                    else:
                        df['상품명'] = df[name_col]

                # 3. 가격 컬럼 찾기
                price_col = None
                price_keywords = ['최종가격', '현재가', '상품가', '공급가', '판매가', '단가', '도매가', '가격', '금액']
                for col in df.columns:
                    col_clean = str(col).replace(" ", "")
                    if any(k in col_clean for k in price_keywords) and not any(k in col_clean for k in ['이전', '변동', '비고']):
                        price_col = col
                        break
                
                if not price_col or '상품명' not in df.columns: continue
                
                df = df.rename(columns={price_col: '공급가'})
                df['공급사'] = supplier
                
                # 4. 가격 숫자 정제
                df['공급가_str'] = df['공급가'].astype(str).str.replace(',','').str.replace('.','').str.replace('원','').str.replace(' ','')
                df = df[df['공급가_str'].str.contains(r'\d', na=False)].copy()
                df['공급가_int'] = pd.to_numeric(df['공급가_str'].str.extract(r'(\d+)')[0], errors='coerce').fillna(0).astype(int)
                df = df[df['공급가_int'] > 0].copy()
                df['공급가'] = df['공급가_int']

                # 5. [핵심 수정] 배송비 합산 로직 (제주/도서산간 비용 차단)
                if "우람씨푸드" in supplier:
                    df['공급가'] = df['공급가'] + 4000
                else:
                    shipping_col = next((c for c in df.columns if any(k in str(c).replace(" ", "") for k in ['배송비', '운임비', '택배비'])), None)
                    if shipping_col:
                        def parse_shipping(val):
                            # 줄바꿈(\n)을 기준으로 첫 번째 줄만 가져옴 (제주/도서산간 할증 텍스트 제거)
                            first_line = str(val).split('\n')[0].replace(" ", "").replace(",", "")
                            
                            # 1. '무배', '무료'가 첫 줄에 있으면 0원
                            if any(k in first_line for k in ['무배', '무료']): 
                                return 0
                                
                            # 2. '0원', '4000원' 처럼 '원' 바로 앞에 붙은 숫자를 정밀하게 찾음
                            match = re.search(r'([0-9]+)원', first_line)
                            if match:
                                return int(match.group(1))
                                
                            # 3. 그 외의 경우 첫 줄의 숫자만 추출
                            nums = "".join(filter(str.isdigit, first_line))
                            return int(nums) if nums else 0
                            
                        df['배송비_숫자'] = df[shipping_col].apply(parse_shipping)
                        df['공급가'] = df['공급가'] + df['배송비_숫자']

                all_rows.append(df)
                print(f"  [성공] {supplier} ({sheet_name}) - ✅ {len(df)}개 상품 파싱 완료")

                # 6. 최종 저장
                if all_rows:
                    combined = pd.concat(all_rows, ignore_index=True)
                    
                    # 사장님의 기존 분석 함수 호출 (parse_product_info가 정의되어 있다고 가정)
                    parsed_data = combined.apply(parse_product_info, axis=1)
                    combined[['등급','크기','중량','과수','통합상품명']] = parsed_data
                    combined['상품명'] = combined['통합상품명']
                    
                    cols = ['등급','크기','중량','과수','공급가','공급사','상품명']
                    combined[[c for c in cols if c in combined.columns]].to_excel(OUTPUT_FILE, index=False)
                    print(f"\n★ 데이터 통합 완료! 파일명: {OUTPUT_FILE}")

                    # ---------------------------------------------------------
                    # 🗑️ [여기에 추가!] 통합 완료 후 원본 엑셀 파일들 싹 다 지우기
                    # ---------------------------------------------------------
                    print("\n[임시 다운로드 파일 정리 시작]")
                    for filename in os.listdir(DOWNLOAD_DIR):
                        filepath = os.path.join(DOWNLOAD_DIR, filename)
                        try:
                            if os.path.isfile(filepath): # 폴더가 아닌 파일만 선택해서
                                os.remove(filepath)      # 영구 삭제!
                        except Exception as e:
                            print(f"  [에러] {filename} 삭제 실패: {e}")
                    print("✅ 다운로드 폴더 내 원본 파일이 모두 깔끔하게 청소되었습니다!")
                    # ---------------------------------------------------------

                else:
                    print("\n❌ 통합할 데이터가 없습니다.")

            wb.close()

        except Exception as e:
            print(f"  [에러] {filename} 처리 중 오류: {e}")

    # 6. 최종 저장
    if all_rows:
        combined = pd.concat(all_rows, ignore_index=True)
        
        # 사장님의 기존 분석 함수 호출 (parse_product_info가 정의되어 있다고 가정)
        parsed_data = combined.apply(parse_product_info, axis=1)
        combined[['등급','크기','중량','과수','통합상품명']] = parsed_data
        combined['상품명'] = combined['통합상품명']
        
        cols = ['등급','크기','중량','과수','공급가','공급사','상품명']
        combined[[c for c in cols if c in combined.columns]].to_excel(OUTPUT_FILE, index=False)
        print(f"\n★ 데이터 통합 완료! 파일명: {OUTPUT_FILE}")
    else:
        print("\n❌ 통합할 데이터가 없습니다.") """

import gc  # 🗑️ 메모리 강제 청소를 위해 상단 추가

def integrate_data():
    all_rows = []
    if not os.path.exists(DOWNLOAD_DIR): 
        print("다운로드 폴더가 없습니다.")
        return

    print("\n[데이터 표준화 및 통합 작업 시작]")
    for filename in os.listdir(DOWNLOAD_DIR):
        if not filename.endswith(".xlsx") or filename.startswith("~$"): continue
        filepath = os.path.join(DOWNLOAD_DIR, filename)
        supplier = filename.split('_')[0]
        
        wb = None # 파일 점유를 풀기 위한 변수 초기화
        try:
            # 1. openpyxl로 열기
            wb = load_workbook(filepath, data_only=True)
            
            # 2. pandas 파일 잠금 방지를 위해 with 구문(컨텍스트 매니저) 사용
            with pd.ExcelFile(filepath) as xls:
                excel_file_dict = pd.read_excel(xls, sheet_name=None, header=None, dtype=str)
            
            for sheet_name, df in excel_file_dict.items():
                if df.empty: continue
                
                ws = wb[sheet_name]
                hidden_rows = [idx for idx, dim in ws.row_dimensions.items() if dim.hidden]
                
                visible_indices = [i for i in range(len(df)) if (i + 1) not in hidden_rows]
                df = df.iloc[visible_indices].reset_index(drop=True)
                
                if df.empty: continue

                header_idx = -1
                for idx, row in df.iterrows():
                    row_values = [str(cell).replace(" ", "").replace("\n", "") for cell in row.values]
                    if any(len(v) <= 10 and any(k in v for k in ['품명', '상품명', '옵션', '품목', '현재가', '최종가격', '상품가']) for v in row_values):
                        header_idx = idx
                        break
                
                if header_idx == -1: continue
                
                df.columns = df.iloc[header_idx]
                df = df.iloc[header_idx+1:].reset_index(drop=True)
                df = df.loc[:, df.columns.notna()]
                df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

                name_col = None
                opt_col = next((c for c in df.columns if str(c).replace(" ","") in ["옵션", "옵션명"]), None)
                
                name_keywords = ['상품명', '품명', '품목명', '품목', '발주서상품명']
                for k in name_keywords:
                    target = next((c for c in df.columns if str(c).replace(" ", "") == k), None)
                    if target:
                        name_col = target
                        break

                if not name_col and opt_col:
                    name_col = opt_col
                    opt_col = None 

                if name_col:
                    if opt_col and "트라이365" in supplier:
                        df[name_col] = df[name_col].ffill()
                        df['상품명'] = df[name_col].astype(str) + " " + df[opt_col].astype(str).replace("nan", "")
                    else:
                        df['상품명'] = df[name_col]

                price_col = None
                price_keywords = ['최종가격', '현재가', '상품가', '공급가', '판매가', '단가', '도매가', '가격', '금액']
                for col in df.columns:
                    col_clean = str(col).replace(" ", "")
                    if any(k in col_clean for k in price_keywords) and not any(k in col_clean for k in ['이전', '변동', '비고']):
                        price_col = col
                        break
                
                if not price_col or '상품명' not in df.columns: continue
                
                df = df.rename(columns={price_col: '공급가'})
                df['공급사'] = supplier
                
                df['공급가_str'] = df['공급가'].astype(str).str.replace(',','').str.replace('.','').str.replace('원','').str.replace(' ','')
                df = df[df['공급가_str'].str.contains(r'\d', na=False)].copy()
                df['공급가_int'] = pd.to_numeric(df['공급가_str'].str.extract(r'(\d+)')[0], errors='coerce').fillna(0).astype(int)
                df = df[df['공급가_int'] > 0].copy()
                df['공급가'] = df['공급가_int']

                if "우람씨푸드" in supplier:
                    df['공급가'] = df['공급가'] + 4000
                else:
                    shipping_col = next((c for c in df.columns if any(k in str(c).replace(" ", "") for k in ['배송비', '운임비', '택배비'])), None)
                    if shipping_col:
                        def parse_shipping(val):
                            first_line = str(val).split('\n')[0].replace(" ", "").replace(",", "")
                            if any(k in first_line for k in ['무배', '무료']): 
                                return 0
                            match = re.search(r'([0-9]+)원', first_line)
                            if match:
                                return int(match.group(1))
                            nums = "".join(filter(str.isdigit, first_line))
                            return int(nums) if nums else 0
                            
                        df['배송비_숫자'] = df[shipping_col].apply(parse_shipping)
                        df['공급가'] = df['공급가'] + df['배송비_숫자']

                all_rows.append(df)
                print(f"  [성공] {supplier} ({sheet_name}) - ✅ {len(df)}개 상품 파싱 완료")

        except Exception as e:
            print(f"  [에러] {filename} 처리 중 오류: {e}")
        finally:
            # ★ 핵심 해결책: 에러가 나든 안 나든 파일의 멱살을 무조건 놓아줌!
            if wb is not None:
                wb.close()

    # 6. 최종 저장 및 삭제
    if all_rows:
        combined = pd.concat(all_rows, ignore_index=True)
        
        parsed_data = combined.apply(parse_product_info, axis=1)
        combined[['등급','크기','중량','과수','통합상품명']] = parsed_data
        combined['상품명'] = combined['통합상품명']
        
        cols = ['등급','크기','중량','과수','공급가','공급사','상품명']
        combined[[c for c in cols if c in combined.columns]].to_excel(OUTPUT_FILE, index=False)
        print(f"\n★ 데이터 통합 완료! 파일명: {OUTPUT_FILE}")

        # ★ 핵심 해결책 2: 파이썬이 잡고 있는 메모리를 강제로 청소해서 잠금을 완벽히 해제
        gc.collect()

        print("\n[임시 다운로드 파일 정리 시작]")
        for filename in os.listdir(DOWNLOAD_DIR):
            filepath = os.path.join(DOWNLOAD_DIR, filename)
            try:
                if os.path.isfile(filepath):
                    os.remove(filepath)
            except Exception as e:
                print(f"  [에러] {filename} 삭제 실패: {e}")
        print("✅ 다운로드 폴더 내 원본 파일이 모두 깔끔하게 청소되었습니다!")

    else:
        print("\n❌ 통합할 데이터가 없습니다.")



# --- 메인 실행 루프 (여기를 덮어씌워주세요) ---
async def main():
    # 다운로드 폴더가 없으면 생성
    if not os.path.exists(DOWNLOAD_DIR): 
        os.makedirs(DOWNLOAD_DIR)
    
    print("🚀 제철한가득 B2B 도매처 자동 수집 및 통합을 시작합니다...")

    # [1~3] 엑셀 자동 다운로드 로직
    async with async_playwright() as p:
        # 브라우저 띄우기 (화면 보려면 False, 숨기려면 True)
        browser = await p.chromium.launch(headless=False) 
        context = await browser.new_context(
            accept_downloads=True,
            viewport={'width': 1920, 'height': 1080}
        )

        # 1. 발주오라 계열 업체 다운로드
        try:
            for company in BALJUORA_COMPANIES:
                await download_baljuora(context, company)
                await asyncio.sleep(1) 
        except NameError:
            pass

        # 2. 어드민 플러스 계열 업체 다운로드 (사장님 코드에 맞춰 추가)
        try:
            for company in ADMIN_PLUS_COMPANIES:
                await download_admin_plus(context, company)
                await asyncio.sleep(1)
        except NameError:
            pass

        # 3. 로그인 불필요 업체 (대봉유통 등) 다운로드
        try:
            for company in DIRECT_DOWNLOAD_COMPANIES:
                await download_direct(context, company)
                await asyncio.sleep(1)
        except NameError:
            pass

        await browser.close()
        print("\n✅ 모든 도매처 엑셀 자동 다운로드가 완료되었습니다!")
    
    # [4] 브라우저 없이 빠르게 다운받는 구글 시트 업체 추가! 🌟
    try:
        # 사장님이 설정해두신 구글시트 리스트 이름(예: GOOGLE_SHEET_COMPANIES)에 맞춰주세요.
        for company in GOOGLE_SHEET_COMPANIES:
            download_google_sheet(company)
    except NameError:
        pass

    # [5] 전체 데이터 통합 (이미 완벽하게 세팅됨!)
    print("\n[전체 데이터 통합 및 파싱 시작]")
    integrate_data()

if __name__ == "__main__":
    asyncio.run(main())
