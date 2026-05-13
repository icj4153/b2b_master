import pandas as pd
import os
import re
from datetime import datetime

# [설정] 파일 경로
DOWNLOAD_DIR = "b2b_downloads"
OUTPUT_FILE = f"integrated_products_{datetime.now().strftime('%Y%m%d')}.xlsx"


def normalize_weight_value(value):
    text = str(value).strip()
    return re.sub(r'(kg|g)$', lambda match: match.group(1).lower(), text, flags=re.IGNORECASE)

def parse_product_info(row):
    """
    상품명에서 등급, 크기, 중량, 과수를 추출하고 단위를 통일함
    """
    name = str(row['상품명'])
    
    # 1. 등급 추출
    grade = "일반"
    if any(k in name for k in ["쥬스용", "주스용"]): grade = "쥬스용"
    elif "실속" in name: grade = "가정용"
    elif "가정용" in name: grade = "가정용"
    elif any(k in name for k in ["선물세트", "선물"]): grade = "선물세트"
    elif "정품" in name: grade = "정품"
    
    # 2. 크기 추출
    size = "미표기"
    if any(k in name for k in ["혼합", "랜덤"]):
        size = "혼합과"
    else:
        size_keywords = ["혼합과", "소과", "중소과", "중과", "중대과", "대과"]
        for k in size_keywords:
            if k in name:
                size = k
                break
            
    # 3. 과수 추출 및 단위 통일 (5개, 5입, 5알 -> 5과)
    # 패턴: 숫자 + (개|입|알|과) 혹은 숫자~숫자 + (개|입|알|과)
    count = "미표기"
    # (\d+~?\d*) : 숫자 또는 숫자~숫자 범위 추출
    # [개입알과] : 뒤에 붙는 단위들
    count_match = re.search(r'(\d+[-~]?\d*)([개입알과])', name)
    
    if count_match:
        number_part = count_match.group(1) # 숫자 부분만 추출 (예: 5 또는 12-15)
        count = f"{number_part}과" # 사장님이 원하시는 '~과'로 강제 통일
    
    # 4. 중량 추출 (숫자 + kg/g)
    weight_match = re.search(r'(\d+\.?\d*)(kg|g)', name, re.IGNORECASE)
    weight = normalize_weight_value(weight_match.group(0)) if weight_match else "미표기"
    
    return pd.Series([grade, size, weight, count])

def integrate_excels():
    all_data = []
    
    if not os.path.exists(DOWNLOAD_DIR):
        print("다운로드 폴더가 없습니다.")
        return

    for filename in os.listdir(DOWNLOAD_DIR):
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            filepath = os.path.join(DOWNLOAD_DIR, filename)
            supplier_name = filename.split('_')[0]
            
            try:
                df = pd.read_excel(filepath)
                df['공급사'] = supplier_name
                
                # 공급가 컬럼 매핑 (판매가, 매입가 등 업체마다 다른 명칭 통일)
                price_cols = ['공급가', '판매가', '매입가', '단가']
                for col in price_cols:
                    if col in df.columns:
                        df = df.rename(columns={col: '공급가'})
                        break
                
                all_data.append(df)
            except Exception as e:
                print(f"[{filename}] 읽기 실패: {e}")

    if not all_data:
        print("합칠 데이터가 없습니다.")
        return

    combined_df = pd.concat(all_data, ignore_index=True)

    if '상품명' in combined_df.columns:
        print("데이터 정밀 파싱 중 (단위 통일 포함)...")
        combined_df[['등급', '크기', '중량', '과수']] = combined_df.apply(parse_product_info, axis=1)
        
        # 사장님 요청 순서대로 정렬
        final_columns = ['등급', '크기', '중량', '과수', '공급가', '공급사', '상품명']
        existing_cols = [c for c in final_columns if c in combined_df.columns]
        
        result_df = combined_df[existing_cols]
        result_df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n★ 데이터 정리 완료! 파일명: {OUTPUT_FILE}")
        print(f"총 {len(result_df)}개의 상품 데이터가 정리되었습니다.")
    else:
        print("에러: '상품명' 컬럼이 엑셀에 없습니다.")

if __name__ == "__main__":
    integrate_excels()
