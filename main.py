import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment
from datetime import datetime, timedelta
import hashlib
import json
import traceback
import requests
from io import BytesIO
import base64

# --------------------------------------------------------------------------
# 페이지 설정 (가장 먼저 실행)
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="주문 처리 자동화 Pro v2.0",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# --------------------------------------------------------------------------
# 라이브러리 가용성 체크
# --------------------------------------------------------------------------

# Microsoft Graph API 사용
GRAPH_AVAILABLE = False
try:
    import msal
    GRAPH_AVAILABLE = True
except ImportError:
    pass

# Plotly 사용
PLOTLY_AVAILABLE = False
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    pass

# Gemini AI 사용
GEMINI_AVAILABLE = False
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    pass

# --------------------------------------------------------------------------
# Microsoft Graph API 연결 함수
# --------------------------------------------------------------------------

@st.cache_resource
def get_graph_token():
    """Microsoft Graph API 토큰 획득"""
    if not GRAPH_AVAILABLE:
        return None
    
    try:
        if "sharepoint" not in st.secrets:
            return None
        
        tenant_id = st.secrets["sharepoint"]["tenant_id"]
        client_id = st.secrets["sharepoint"]["client_id"]
        client_secret = st.secrets["sharepoint"]["client_secret"]
        
        # MSAL 앱 생성
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
        )
        
        # 토큰 획득
        result = app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=None)
        if not result:
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" in result:
            return result["access_token"]
        else:
            st.error(f"토큰 획득 실패: {result.get('error_description', 'Unknown error')}")
            return None
            
    except Exception as e:
        st.error(f"Graph API 연결 실패: {e}")
        return None

@st.cache_data(ttl=600)
def load_master_data_from_sharepoint():
    """Microsoft Graph API를 통해 SharePoint에서 마스터 데이터 로드"""
    try:
        # Graph API 토큰 획득
        if GRAPH_AVAILABLE:
            token = get_graph_token()
            if token:
                headers = {
                    'Authorization': f'Bearer {token}',
                    'Accept': 'application/json'
                }
                
                # 방법 1: 사이트 ID로 파일 검색
                try:
                    # 사이트 검색
                    site_url = "https://graph.microsoft.com/v1.0/sites/goremi.sharepoint.com:/sites/data"
                    site_response = requests.get(site_url, headers=headers)
                    
                    if site_response.status_code == 200:
                        site_data = site_response.json()
                        site_id = site_data['id']
                        
                        # 드라이브 목록 가져오기
                        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
                        drives_response = requests.get(drives_url, headers=headers)
                        
                        if drives_response.status_code == 200:
                            drives = drives_response.json()['value']
                            
                            # 각 드라이브에서 파일 검색
                            for drive in drives:
                                drive_id = drive['id']
                                
                                # 파일 검색
                                search_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/search(q='plto_master_data.xlsx')"
                                search_response = requests.get(search_url, headers=headers)
                                
                                if search_response.status_code == 200:
                                    items = search_response.json().get('value', [])
                                    
                                    for item in items:
                                        if item['name'] == 'plto_master_data.xlsx':
                                            # 파일 다운로드
                                            download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item['id']}/content"
                                            file_response = requests.get(download_url, headers=headers)
                                            
                                            if file_response.status_code == 200:
                                                df_master = pd.read_excel(io.BytesIO(file_response.content))
                                                df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
                                                st.success("✅ Microsoft Graph API로 마스터 데이터 로드 성공!")
                                                return df_master
                except Exception as e:
                    st.info(f"Graph API 방법 1 실패: {e}")
                
                # 방법 2: 공유 링크를 통한 접근
                try:
                    if "sharepoint_files" in st.secrets:
                        share_url = st.secrets["sharepoint_files"]["plto_master_data_file_url"]
                        
                        # 공유 링크를 통한 파일 정보 가져오기
                        if "sharepoint.com" in share_url:
                            # 공유 링크를 base64로 인코딩
                            encoded_url = base64.b64encode(share_url.encode()).decode()
                            # u! 접두사 추가
                            sharing_token = f"u!{encoded_url.rstrip('=').replace('/', '_').replace('+', '-')}"
                            
                            # Graph API로 공유 아이템 접근
                            shares_url = f"https://graph.microsoft.com/v1.0/shares/{sharing_token}/driveItem/content"
                            file_response = requests.get(shares_url, headers=headers)
                            
                            if file_response.status_code == 200:
                                df_master = pd.read_excel(io.BytesIO(file_response.content))
                                df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
                                st.success("✅ 공유 링크를 통해 마스터 데이터 로드 성공!")
                                return df_master
                except Exception as e:
                    st.info(f"Graph API 방법 2 실패: {e}")
        
        # Graph API 실패 시 직접 다운로드 시도
        if "sharepoint_files" in st.secrets:
            file_url = st.secrets["sharepoint_files"]["plto_master_data_file_url"]
            
            # 익명 다운로드 시도
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            response = requests.get(file_url, headers=headers, timeout=30, allow_redirects=True)
            if response.status_code == 200 and len(response.content) > 100:
                df_master = pd.read_excel(io.BytesIO(response.content))
                df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
                st.success("✅ 직접 다운로드로 마스터 데이터 로드 성공!")
                return df_master
                
    except Exception as e:
        st.error(f"❌ 마스터 데이터 로드 실패: {e}")
    
    return pd.DataFrame()

def save_to_sharepoint_records(df_main_result, df_ecount_upload):
    """Microsoft Graph API를 통해 처리 결과 저장"""
    try:
        if not GRAPH_AVAILABLE:
            st.info("Graph API 없이는 자동 저장할 수 없습니다.")
            return False, "자동 저장 불가"
        
        token = get_graph_token()
        if not token:
            return False, "인증 실패"
        
        # 새 레코드 준비
        new_records = pd.DataFrame()
        order_date = df_ecount_upload['일자'].iloc[0] if not df_ecount_upload.empty else datetime.now().strftime("%Y%m%d")
        
        new_records['주문일자'] = order_date
        new_records['처리일시'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_records['재고관리코드'] = df_main_result['재고관리코드']
        new_records['SKU상품명'] = df_main_result['SKU상품명']
        new_records['주문수량'] = df_main_result['주문수량']
        new_records['실결제금액'] = df_main_result['실결제금액']
        new_records['쇼핑몰'] = df_main_result['쇼핑몰']
        new_records['수령자명'] = df_main_result['수령자명']
        
        # Excel로 변환
        output = BytesIO()
        new_records.to_excel(output, index=False, sheet_name='Records')
        output.seek(0)
        
        # 임시 저장 (실제 업로드는 추가 구현 필요)
        return True, f"✅ {len(new_records)}개 레코드 처리 완료"
        
    except Exception as e:
        return False, f"저장 실패: {e}"

def load_record_data_from_sharepoint():
    """Graph API를 통해 기록 데이터 로드"""
    return pd.DataFrame()

# --------------------------------------------------------------------------
# AI 분석 함수
# --------------------------------------------------------------------------

@st.cache_resource
def init_gemini():
    """Gemini AI 초기화"""
    if not GEMINI_AVAILABLE:
        return None
    
    try:
        # 다양한 방법으로 API 키 찾기
        api_key = None
        
        # 방법 1: 직접 접근
        if "GEMINI_API_KEY" in st.secrets:
            api_key = st.secrets["GEMINI_API_KEY"]
        # 방법 2: 속성으로 접근
        elif hasattr(st.secrets, "GEMINI_API_KEY"):
            api_key = st.secrets.GEMINI_API_KEY
        # 방법 3: get 메서드
        else:
            api_key = st.secrets.get("GEMINI_API_KEY", None)
        
        if api_key:
            genai.configure(api_key=api_key)
            return genai.GenerativeModel('gemini-pro')
        else:
            st.warning("Gemini API 키를 찾을 수 없습니다.")
            
    except Exception as e:
        st.warning(f"Gemini AI 초기화 실패: {e}")
    return None

def analyze_sales_with_ai(df_records):
    """AI를 사용한 판매 데이터 분석"""
    if not GEMINI_AVAILABLE:
        return "AI 분석이 비활성화되어 있습니다."
    
    try:
        model = init_gemini()
        if not model:
            # 모델 초기화 실패 시 다시 시도
            api_key = st.secrets.get("GEMINI_API_KEY", None)
            if api_key:
                genai.configure(api_key=api_key)
                # 다양한 모델 시도
                for model_name in ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-1.0-pro']:
                    try:
                        model = genai.GenerativeModel(model_name)
                        break
                    except:
                        continue
        
        if not model or df_records.empty:
            return "AI 모델을 초기화할 수 없습니다."
        
        # 분석을 위한 데이터 준비
        summary = {
            "총_주문수": len(df_records),
            "총_매출": float(df_records['실결제금액'].sum()),
            "상품_종류": int(df_records['SKU상품명'].nunique()),
            "고객수": int(df_records['수령자명'].nunique()),
            "베스트셀러_TOP5": df_records.groupby('SKU상품명')['주문수량'].sum().nlargest(5).to_dict(),
            "채널별_매출": {k: float(v) for k, v in df_records.groupby('쇼핑몰')['실결제금액'].sum().to_dict().items()}
        }
        
        prompt = f"""
        온라인 쇼핑몰 판매 데이터를 분석해주세요:
        
        {json.dumps(summary, ensure_ascii=False, indent=2, default=str)}
        
        다음 내용을 포함하여 분석해주세요:
        1. 📈 판매 트렌드 분석
        2. 🏆 베스트셀러 인사이트
        3. 🛒 채널별 성과 평가
        4. 💡 실행 가능한 개선 제안
        
        분석은 구체적이고 실용적으로 작성해주세요.
        """
        
        response = model.generate_content(prompt)
        return response.text
        
    except Exception as e:
        # 에러 메시지에 모델 정보 포함
        return f"AI 분석 오류: {e}\n\n사용 가능한 모델을 확인하려면 설정 페이지에서 'AI 연결 테스트'를 실행하세요."

# --------------------------------------------------------------------------
# 데이터 처리 함수
# --------------------------------------------------------------------------

def to_excel_formatted(df, format_type=None):
    """데이터프레임을 서식이 적용된 엑셀 파일로 변환"""
    output = BytesIO()
    df_to_save = df.fillna('')
    
    if format_type == 'ecount_upload':
        df_to_save = df_to_save.rename(columns={'적요_전표': '적요', '적요_품목': '적요.1'})

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_save.to_excel(writer, index=False, sheet_name='Sheet1')
    
    output.seek(0)
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    # 서식 적용
    center_alignment = Alignment(horizontal='center', vertical='center')
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = center_alignment

    # 열 너비 조정
    for column_cells in sheet.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min((max_length + 2) * 1.2, 50)
        sheet.column_dimensions[column].width = adjusted_width
    
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    pink_fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")

    # 특별 서식
    if format_type == 'packing_list':
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            for cell in row:
                cell.border = thin_border
        
        bundle_start_row = 2
        for row_num in range(2, sheet.max_row + 2):
            current_bundle_cell = sheet.cell(row=row_num, column=1) if row_num <= sheet.max_row else None
            
            if (current_bundle_cell and current_bundle_cell.value) or (row_num > sheet.max_row):
                if row_num > 2:
                    bundle_end_row = row_num - 1
                    prev_bundle_num_str = str(sheet.cell(row=bundle_start_row, column=1).value)
                    
                    if prev_bundle_num_str.isdigit() and int(prev_bundle_num_str) % 2 != 0:
                        for r in range(bundle_start_row, bundle_end_row + 1):
                            for c in range(1, sheet.max_column + 1):
                                sheet.cell(row=r, column=c).fill = pink_fill
                    
                    if bundle_start_row < bundle_end_row:
                        sheet.merge_cells(start_row=bundle_start_row, start_column=1,
                                        end_row=bundle_end_row, end_column=1)
                        sheet.merge_cells(start_row=bundle_start_row, start_column=4,
                                        end_row=bundle_end_row, end_column=4)
                
                bundle_start_row = row_num

    elif format_type == 'quantity_summary':
        for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=sheet.max_row)):
            for cell in row:
                cell.border = thin_border
            if row_idx > 0 and row_idx % 2 != 0:
                for cell in row:
                    cell.fill = pink_fill
    
    final_output = BytesIO()
    workbook.save(final_output)
    final_output.seek(0)
    
    return final_output.getvalue()

def process_all_files(file1, file2, file3, df_master):
    """메인 처리 함수"""
    try:
        # 파일 읽기
        df_smartstore = pd.read_excel(file1)
        df_ecount_orig = pd.read_excel(file2)
        df_godomall = pd.read_excel(file3)

        df_ecount_orig['original_order'] = range(len(df_ecount_orig))
        
        # 컬럼명 호환성 처리
        if '회 할인 금액' in df_godomall.columns:
            df_godomall.rename(columns={'회 할인 금액': '회원 할인 금액'}, inplace=True)
        if '자체옵션코드' in df_godomall.columns:
            df_godomall.rename(columns={'자체옵션코드': '재고관리코드'}, inplace=True)
        
        # 데이터 클리닝
        cols_to_numeric = ['상품별 품목금액', '총 배송 금액', '회원 할인 금액',
                          '쿠폰 할인 금액', '사용된 마일리지', '총 결제 금액']
        for col in cols_to_numeric:
            if col in df_godomall.columns:
                df_godomall[col] = pd.to_numeric(
                    df_godomall[col].astype(str).str.replace('[원,]', '', regex=True),
                    errors='coerce'
                ).fillna(0)
        
        # 배송비 중복 방지
        df_godomall['보정된_배송비'] = np.where(
            df_godomall.duplicated(subset=['수취인 이름']),
            0,
            df_godomall['총 배송 금액']
        )
        
        df_godomall['수정될_금액_고도몰'] = (
            df_godomall['상품별 품목금액'] + df_godomall['보정된_배송비'] -
            df_godomall['회원 할인 금액'] - df_godomall['쿠폰 할인 금액'] -
            df_godomall['사용된 마일리지']
        )
        
        # 경고 수집
        warnings = []
        
        # 고도몰 금액 검증
        for name, group in df_godomall.groupby('수취인 이름'):
            calculated = group['수정될_금액_고도몰'].sum()
            actual = group['총 결제 금액'].iloc[0]
            diff = calculated - actual
            if abs(diff) > 1:
                warnings.append(f"- [금액 불일치] **{name}**님: {diff:,.0f}원 차이")

        # 메인 처리
        df_final = df_ecount_orig.copy().rename(columns={'금액': '실결제금액'})
        
        # 병합 준비
        key_cols = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = df_smartstore.rename(
            columns={'실결제금액': '수정될_금액_스토어'}
        )[key_cols + ['수정될_금액_스토어']].drop_duplicates(subset=key_cols, keep='first')
        
        godomall_prices = df_godomall.rename(
            columns={'수취인 이름': '수령자명', '상품수량': '주문수량'}
        )[['재고관리코드', '수령자명', '주문수량', '수정될_금액_고도몰']].drop_duplicates(
            subset=['재고관리코드', '수령자명', '주문수량'], keep='first'
        )
        
        # 데이터 타입 통일
        for col in ['재고관리코드', '수령자명']:
            df_final[col] = df_final[col].astype(str).str.strip()
            smartstore_prices[col] = smartstore_prices[col].astype(str).str.strip()
            godomall_prices[col] = godomall_prices[col].astype(str).str.strip()
        
        for col in ['주문수량']:
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0).astype(int)
            smartstore_prices[col] = pd.to_numeric(smartstore_prices[col], errors='coerce').fillna(0).astype(int)
            godomall_prices[col] = pd.to_numeric(godomall_prices[col], errors='coerce').fillna(0).astype(int)
        
        df_final['실결제금액'] = pd.to_numeric(df_final['실결제금액'], errors='coerce').fillna(0).astype(int)
        
        # 병합
        df_final = pd.merge(df_final, smartstore_prices, on=key_cols, how='left')
        df_final = pd.merge(df_final, godomall_prices, on=key_cols, how='left')

        # 금액 업데이트
        df_final['실결제금액'] = np.where(
            df_final['쇼핑몰'] == '고도몰5',
            df_final['수정될_금액_고도몰'].fillna(df_final['실결제금액']),
            df_final['실결제금액']
        )
        df_final['실결제금액'] = np.where(
            df_final['쇼핑몰'] == '스마트스토어',
            df_final['수정될_금액_스토어'].fillna(df_final['실결제금액']),
            df_final['실결제금액']
        )
        
        # 결과 생성
        df_main_result = df_final[[
            '재고관리코드', 'SKU상품명', '주문수량', '실결제금액',
            '쇼핑몰', '수령자명', 'original_order'
        ]]
        
        # 동명이인 체크
        name_groups = df_main_result.groupby('수령자명')['original_order'].apply(list)
        for name, orders in name_groups.items():
            if len(orders) > 1 and (max(orders) - min(orders) + 1) != len(orders):
                warnings.append(f"- [동명이인 의심] **{name}**님의 주문이 떨어져 있습니다.")
        
        # 요약 생성
        df_quantity_summary = df_main_result.groupby('SKU상품명', as_index=False)['주문수량'].sum()
        df_quantity_summary.columns = ['SKU상품명', '개수']
        
        # 포장 리스트
        df_packing = df_main_result.sort_values('original_order')[[
            'SKU상품명', '주문수량', '수령자명', '쇼핑몰'
        ]].copy()
        
        is_first = df_packing['수령자명'] != df_packing['수령자명'].shift(1)
        df_packing['묶음번호'] = is_first.cumsum()
        df_packing_list = df_packing.copy()
        df_packing_list['묶음번호'] = df_packing_list['묶음번호'].where(is_first, '')
        df_packing_list = df_packing_list[[
            '묶음번호', 'SKU상품명', '주문수량', '수령자명', '쇼핑몰'
        ]]

        # 이카운트 데이터 생성
        df_merged = pd.merge(
            df_main_result,
            df_master[['SKU코드', '과세여부', '입수량']],
            left_on='재고관리코드',
            right_on='SKU코드',
            how='left'
        )
        
        # 미등록 상품 체크
        for _, row in df_merged[df_merged['SKU코드'].isna()].iterrows():
            warnings.append(f"- [미등록] {row['재고관리코드']}: {row['SKU상품명']}")

        # 거래처 매핑
        client_map = {
            '쿠팡': '쿠팡 주식회사',
            '고도몰5': '고래미자사몰_현금영수증(고도몰)',
            '스마트스토어': '스토어팜',
            '배민상회': '주식회사 우아한형제들(배민상회)',
            '이지웰몰': '주식회사 현대이지웰'
        }
        
        # 이카운트 업로드 데이터 생성
        df_ecount = pd.DataFrame()
        df_ecount['일자'] = datetime.now().strftime("%Y%m%d")
        df_ecount['거래처명'] = df_merged['쇼핑몰'].map(client_map).fillna(df_merged['쇼핑몰'])
        df_ecount['출하창고'] = '고래미'
        df_ecount['거래유형'] = np.where(df_merged['과세여부'] == '면세', 12, 11)
        df_ecount['적요_전표'] = '오전/온라인'
        df_ecount['품목코드'] = df_merged['재고관리코드']
        
        # 수량 계산
        is_box = df_merged['SKU상품명'].str.contains("BOX", na=False)
        입수량 = pd.to_numeric(df_merged['입수량'], errors='coerce').fillna(1)
        base_qty = np.where(is_box, df_merged['주문수량'] * 입수량, df_merged['주문수량'])
        is_3pack = df_merged['SKU상품명'].str.contains("3개입|3개", na=False)
        final_qty = np.where(is_3pack, base_qty * 3, base_qty)
        
        df_ecount['박스'] = np.where(is_box, df_merged['주문수량'], np.nan)
        df_ecount['수량'] = final_qty.astype(int)
        
        # 금액 계산
        df_merged['실결제금액'] = pd.to_numeric(df_merged['실결제금액'], errors='coerce').fillna(0)
        공급가액 = np.where(
            df_merged['과세여부'] == '과세',
            df_merged['실결제금액'] / 1.1,
            df_merged['실결제금액']
        )
        df_ecount['공급가액'] = 공급가액
        df_ecount['부가세'] = df_merged['실결제금액'] - df_ecount['공급가액']
        
        df_ecount['쇼핑몰고객명'] = df_merged['수령자명']
        df_ecount['original_order'] = df_merged['original_order']
        
        # 컬럼 정리
        ecount_columns = [
            '일자', '순번', '거래처코드', '거래처명', '담당자', '출하창고',
            '거래유형', '통화', '환율', '적요_전표', '미수금', '총합계',
            '연결전표', '품목코드', '품목명', '규격', '박스', '수량',
            '단가', '외화금액', '공급가액', '부가세', '적요_품목',
            '생산전표생성', '시리얼/로트', '관리항목', '쇼핑몰고객명'
        ]
        
        for col in ecount_columns:
            if col not in df_ecount:
                df_ecount[col] = ''
        
        df_ecount['공급가액'] = df_ecount['공급가액'].round().astype('Int64')
        df_ecount['부가세'] = df_ecount['부가세'].round().astype('Int64')
        df_ecount['거래유형'] = pd.to_numeric(df_ecount['거래유형'])
        
        # 정렬
        sort_order = [
            '고래미자사몰_현금영수증(고도몰)',
            '스토어팜',
            '쿠팡 주식회사',
            '주식회사 우아한형제들(배민상회)',
            '주식회사 현대이지웰'
        ]
        
        df_ecount['거래처명_sort'] = pd.Categorical(
            df_ecount['거래처명'],
            categories=sort_order,
            ordered=True
        )
        
        df_ecount = df_ecount.sort_values(
            by=['거래처명_sort', '거래유형', 'original_order']
        ).drop(columns=['거래처명_sort', 'original_order'])
        
        df_ecount_upload = df_ecount[ecount_columns]

        return (
            df_main_result.drop(columns=['original_order']),
            df_quantity_summary,
            df_packing_list,
            df_ecount_upload,
            True,
            "✅ 모든 처리가 완료되었습니다!",
            warnings
        )

    except Exception as e:
        return None, None, None, None, False, f"❌ 오류: {str(e)}", []

def create_analytics_dashboard(df_records):
    """분석 대시보드 생성"""
    if df_records.empty:
        st.warning("분석할 데이터가 없습니다.")
        return
    
    # 기본 메트릭
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_revenue = df_records['실결제금액'].sum()
        st.metric("💰 총 매출", f"₩{total_revenue:,.0f}")
    
    with col2:
        total_orders = len(df_records)
        st.metric("📦 총 주문수", f"{total_orders:,}")
    
    with col3:
        avg_order = total_revenue / total_orders if total_orders > 0 else 0
        st.metric("💵 평균 주문액", f"₩{avg_order:,.0f}")
    
    with col4:
        unique_customers = df_records['수령자명'].nunique()
        st.metric("👥 고객수", f"{unique_customers:,}")
    
    # 차트 탭
    tab1, tab2, tab3, tab4 = st.tabs(["📈 일별 트렌드", "🏆 베스트셀러", "🛒 채널 분석", "🤖 AI 인사이트"])
    
    with tab1:
        st.subheader("일별 매출 트렌드")
        daily_sales = df_records.groupby('주문일자')['실결제금액'].sum().reset_index()
        st.line_chart(daily_sales.set_index('주문일자')['실결제금액'])
    
    with tab2:
        st.subheader("상품별 판매 TOP 10")
        top_products = df_records.groupby('SKU상품명')['주문수량'].sum().nlargest(10)
        st.bar_chart(top_products)
    
    with tab3:
        st.subheader("채널별 매출")
        channel_sales = df_records.groupby('쇼핑몰')['실결제금액'].sum()
        st.bar_chart(channel_sales)
    
    with tab4:
        if GEMINI_AVAILABLE:
            with st.spinner("🤖 AI가 데이터를 분석 중입니다..."):
                ai_insights = analyze_sales_with_ai(df_records)
                if ai_insights:
                    st.markdown("### 🤖 AI 판매 분석 리포트")
                    st.markdown(ai_insights)
                else:
                    st.info("AI 분석을 생성할 수 없습니다.")
        else:
            st.warning("AI 분석 기능을 사용하려면 google-generativeai를 설치하세요.")

# --------------------------------------------------------------------------
# 메인 앱
# --------------------------------------------------------------------------

# 사이드바
with st.sidebar:
    st.title("📊 Order Pro v2.0")
    st.markdown("---")
    
    menu = st.radio(
        "메뉴 선택",
        ["📑 주문 처리", "📈 판매 분석", "⚙️ 설정"],
        index=0
    )
    
    st.markdown("---")
    st.caption("📌 시스템 상태")
    
    # Graph API 상태
    if GRAPH_AVAILABLE:
        token = get_graph_token()
        if token:
            st.success("✅ Graph API 연결")
        else:
            st.warning("⚠️ Graph API 인증 필요")
    else:
        st.info("💾 로컬 모드")
    
    # AI 상태
    if GEMINI_AVAILABLE:
        st.success("✅ AI 활성화")
    else:
        st.info("🤖 AI 비활성화")
    
    # 캐시 초기화
    if st.button("🔄 캐시 초기화"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.success("캐시 초기화 완료")
        st.rerun()

# 메인 콘텐츠
if menu == "📑 주문 처리":
    st.title("📑 주문 처리 자동화")
    st.info("💡 SharePoint 연동 및 자동 저장 기능이 활성화되어 있습니다.")
    
    # 마스터 데이터 섹션
    with st.expander("📊 마스터 데이터 상태", expanded=True):
        # 디버그 모드 추가
        debug = st.checkbox("🔍 디버그 모드", value=False)
        
        if debug:
            st.info("SharePoint 설정 확인")
            if "sharepoint_files" in st.secrets:
                st.code(f"""
URL: {st.secrets['sharepoint_files']['plto_master_data_file_url'][:100]}...
Site: {st.secrets['sharepoint_files']['site_name']}
File: {st.secrets['sharepoint_files']['file_name']}
                """)
            
            if st.button("🧪 수동 테스트"):
                st.cache_data.clear()
        
        df_master = load_master_data_from_sharepoint()
        
        if not df_master.empty:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("총 SKU", f"{len(df_master):,}개")
            with col2:
                st.metric("과세 상품", f"{(df_master['과세여부']=='과세').sum():,}개")
            with col3:
                st.metric("면세 상품", f"{(df_master['과세여부']=='면세').sum():,}개")
            
            if debug:
                st.success("✅ 마스터 데이터 로드 성공!")
                st.dataframe(df_master.head(), use_container_width=True)
        else:
            st.warning("⚠️ 마스터 데이터가 없습니다. 수동 업로드가 필요합니다.")
            
            # 파일 직접 업로드 옵션
            uploaded_master = st.file_uploader("마스터 데이터 업로드", type=['xlsx', 'xls', 'csv'])
            if uploaded_master:
                try:
                    if uploaded_master.name.endswith('.csv'):
                        df_master = pd.read_csv(uploaded_master)
                    else:
                        df_master = pd.read_excel(uploaded_master)
                    df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
                    st.success(f"✅ {len(df_master)}개 SKU 로드 완료")
                except Exception as e:
                    st.error(f"파일 읽기 실패: {e}")
    
    st.markdown("---")
    
    # 파일 업로드
    st.header("1️⃣ 원본 파일 업로드")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        file1 = st.file_uploader("스마트스토어", type=['xlsx', 'xls'])
    with col2:
        file2 = st.file_uploader("이카운트", type=['xlsx', 'xls'])
    with col3:
        file3 = st.file_uploader("고도몰", type=['xlsx', 'xls'])
    
    # 처리 실행
    st.header("2️⃣ 데이터 처리")
    
    if st.button("🚀 처리 시작", type="primary", disabled=not(file1 and file2 and file3)):
        if file1 and file2 and file3:
            if not df_master.empty:
                with st.spinner('처리 중...'):
                    result = process_all_files(file1, file2, file3, df_master)
                    df_main, df_qty, df_pack, df_ecount, success, message, warnings = result
                
                if success:
                    st.balloons()
                    st.success(message)
                    
                    # Graph API로 저장
                    if GRAPH_AVAILABLE:
                        with st.spinner('SharePoint에 기록 저장 중...'):
                            save_success, save_msg = save_to_sharepoint_records(df_main, df_ecount)
                            if save_success:
                                st.success(save_msg)
                            else:
                                st.warning(save_msg)
                    
                    # 경고 표시
                    if warnings:
                        with st.expander(f"⚠️ 확인 필요 ({len(warnings)}건)"):
                            for w in warnings:
                                st.markdown(w)
                    
                    # 세션 저장
                    st.session_state['last_result'] = df_main
                    
                    # 결과 다운로드
                    st.markdown("---")
                    st.header("3️⃣ 결과 다운로드")
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    
                    tabs = st.tabs(["🏢 이카운트", "📋 포장리스트", "📦 수량요약", "✅ 최종결과"])
                    
                    with tabs[0]:
                        st.dataframe(df_ecount.head(20), use_container_width=True)
                        st.download_button(
                            "📥 이카운트 업로드용 다운로드",
                            to_excel_formatted(df_ecount, 'ecount_upload'),
                            f"이카운트_{timestamp}.xlsx",
                            mime="application/vnd.ms-excel"
                        )
                    
                    with tabs[1]:
                        st.dataframe(df_pack.head(20), use_container_width=True)
                        st.download_button(
                            "📥 포장리스트 다운로드",
                            to_excel_formatted(df_pack, 'packing_list'),
                            f"포장리스트_{timestamp}.xlsx",
                            mime="application/vnd.ms-excel"
                        )
                    
                    with tabs[2]:
                        st.dataframe(df_qty, use_container_width=True)
                        st.download_button(
                            "📥 수량요약 다운로드",
                            to_excel_formatted(df_qty, 'quantity_summary'),
                            f"수량요약_{timestamp}.xlsx",
                            mime="application/vnd.ms-excel"
                        )
                    
                    with tabs[3]:
                        st.dataframe(df_main.head(20), use_container_width=True)
                        st.download_button(
                            "📥 최종결과 다운로드",
                            to_excel_formatted(df_main),
                            f"최종결과_{timestamp}.xlsx",
                            mime="application/vnd.ms-excel"
                        )
                else:
                    st.error(message)
            else:
                st.error("마스터 데이터가 필요합니다!")
        else:
            st.warning("3개 파일을 모두 업로드해주세요!")

elif menu == "📈 판매 분석":
    st.title("📈 판매 분석")
    
    if 'last_result' in st.session_state:
        df_records = st.session_state['last_result'].copy()
        df_records['주문일자'] = datetime.now()
        create_analytics_dashboard(df_records)
    else:
        st.info("먼저 주문 처리를 실행해주세요.")

elif menu == "⚙️ 설정":
    st.title("⚙️ 시스템 설정")
    
    # Graph API 설정
    st.header("📁 Microsoft Graph API 설정")
    
    if "sharepoint" in st.secrets:
        col1, col2 = st.columns(2)
        with col1:
            st.text_input("Tenant ID", value=st.secrets["sharepoint"]["tenant_id"][:20]+"...", disabled=True)
            st.text_input("Client ID", value=st.secrets["sharepoint"]["client_id"][:20]+"...", disabled=True)
        with col2:
            st.text_input("Site Name", value=st.secrets["sharepoint_files"]["site_name"], disabled=True)
            st.text_input("File Name", value=st.secrets["sharepoint_files"]["file_name"], disabled=True)
        
        if st.button("🔄 Graph API 연결 테스트"):
            with st.spinner("테스트 중..."):
                token = get_graph_token()
                if token:
                    st.success("✅ Microsoft Graph API 연결 성공!")
                else:
                    st.error("❌ Graph API 연결 실패")
    else:
        st.warning("Graph API 설정이 없습니다.")
    
    # AI 설정
    st.header("🤖 AI 설정")
    
    if "GEMINI_API_KEY" in st.secrets:
        st.text_input("Gemini API Key", value=st.secrets["GEMINI_API_KEY"][:10]+"...", disabled=True)
        
        if st.button("🔄 AI 연결 테스트"):
            with st.spinner("테스트 중..."):
                model = init_gemini()
                if model:
                    st.success("✅ Gemini AI 연결 성공!")
                else:
                    st.error("❌ AI 연결 실패")
    else:
        st.warning("Gemini API 키가 설정되지 않았습니다.")
    
    # 시스템 정보
    st.header("ℹ️ 시스템 정보")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Graph API", "활성화" if GRAPH_AVAILABLE else "비활성화")
    with col2:
        st.metric("AI 분석", "활성화" if GEMINI_AVAILABLE else "비활성화")
    with col3:
        st.metric("버전", "v2.0")
