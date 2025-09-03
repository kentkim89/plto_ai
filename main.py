import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import requests
import tempfile
import os
import plotly.express as px
import plotly.graph_objects as go
import google.generativeai as genai

# --------------------------------------------------------------------------
# Streamlit 페이지 설정
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="주문 처리 자동화 시스템 v2.0", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --------------------------------------------------------------------------
# SharePoint 간단한 연결 함수 (Office365 라이브러리 제거)
# --------------------------------------------------------------------------
@st.cache_data(ttl=600)  # 10분 캐시
def load_master_from_sharepoint():
    """SharePoint에서 마스터 데이터 로드 - 직접 URL 방식"""
    try:
        # SharePoint 직접 다운로드 URL로 변환
        share_url = st.secrets["sharepoint_files"]["plto_master_data_file_url"]
        
        # 공유 링크를 다운로드 링크로 변환
        if "sharepoint.com/:x:" in share_url:
            # Excel 공유 링크 패턴
            file_id = share_url.split("/")[-1].split("?")[0]
            download_url = share_url.replace("/:x:/", "/_layouts/15/download.aspx?UniqueId=").split("?")[0]
        else:
            download_url = share_url
        
        # 직접 다운로드 시도
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(share_url, headers=headers, allow_redirects=True)
        
        if response.status_code == 200:
            df_master = pd.read_excel(io.BytesIO(response.content))
            df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
            st.success(f"✅ 마스터 데이터 로드 완료: {len(df_master)}개 품목")
            return df_master
        else:
            raise Exception(f"다운로드 실패: {response.status_code}")
            
    except Exception as e:
        st.warning(f"⚠️ SharePoint 접근 실패: {e}")
        
        # 로컬 파일 폴백
        try:
            if os.path.exists("master_data.csv"):
                df_master = pd.read_csv("master_data.csv")
                df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
                st.info(f"📁 로컬 백업 파일 사용: {len(df_master)}개 품목")
                return df_master
            else:
                # 샘플 데이터 생성 (테스트용)
                st.warning("⚠️ 마스터 파일이 없어 샘플 데이터를 생성합니다.")
                sample_data = {
                    'SKU코드': ['TEST001', 'TEST002', 'TEST003'],
                    '과세여부': ['과세', '면세', '과세'],
                    '입수량': [1, 1, 1]
                }
                return pd.DataFrame(sample_data)
        except Exception as e2:
            st.error(f"로컬 파일도 실패: {e2}")
            return None

def save_record_locally(df_new_records):
    """로컬에 기록 저장 (SharePoint 대신 임시 사용)"""
    try:
        record_file = "plto_record_data.xlsx"
        
        # 기존 파일 로드 또는 새로 생성
        if os.path.exists(record_file):
            df_existing = pd.read_excel(record_file)
        else:
            df_existing = pd.DataFrame()
        
        # 헤더 정의
        expected_columns = [
            '처리일시', '주문일자', '쇼핑몰', '거래처명', '품목코드', 'SKU상품명', 
            '주문수량', '실결제금액', '공급가액', '부가세', '수령자명', 
            '과세여부', '거래유형', '처리자'
        ]
        
        if df_existing.empty:
            df_existing = pd.DataFrame(columns=expected_columns)
        
        # 새 레코드 준비
        df_new_records['처리일시'] = datetime.now()
        df_new_records['처리자'] = st.session_state.get('user_name', 'Unknown')
        
        # 중복 체크
        if not df_existing.empty and '주문일자' in df_existing.columns:
            merge_keys = ['주문일자', '쇼핑몰', '수령자명', '품목코드']
            
            # 컬럼 존재 확인
            keys_exist = all(key in df_existing.columns and key in df_new_records.columns for key in merge_keys)
            
            if keys_exist:
                df_existing['check_key'] = df_existing[merge_keys].astype(str).agg('_'.join, axis=1)
                df_new_records['check_key'] = df_new_records[merge_keys].astype(str).agg('_'.join, axis=1)
                
                new_keys = set(df_new_records['check_key']) - set(df_existing['check_key'])
                df_new_records = df_new_records[df_new_records['check_key'].isin(new_keys)]
                
                df_existing = df_existing.drop('check_key', axis=1, errors='ignore')
                df_new_records = df_new_records.drop('check_key', axis=1, errors='ignore')
        
        # 데이터 결합
        df_combined = pd.concat([df_existing, df_new_records], ignore_index=True)
        
        # 파일 저장
        df_combined.to_excel(record_file, index=False)
        
        return True, len(df_new_records)
        
    except Exception as e:
        st.error(f"기록 저장 실패: {e}")
        return False, 0

def load_record_data():
    """로컬에서 기록 데이터 로드"""
    try:
        record_file = "plto_record_data.xlsx"
        if os.path.exists(record_file):
            return pd.read_excel(record_file)
        else:
            return pd.DataFrame()
    except:
        return pd.DataFrame()

# --------------------------------------------------------------------------
# Gemini AI 분석 함수 (간소화)
# --------------------------------------------------------------------------
def analyze_with_gemini(df_data, analysis_type="trend"):
    """Gemini AI를 사용한 데이터 분석"""
    try:
        genai.configure(api_key=st.secrets["gemini"]["api_key"])
        model = genai.GenerativeModel('gemini-pro')
        
        if df_data.empty:
            return "분석할 데이터가 없습니다."
        
        # 간단한 요약만 생성
        summary = f"""
        총 주문: {len(df_data)}건
        총 매출: {df_data['실결제금액'].sum():,.0f}원
        평균 주문: {df_data['실결제금액'].mean():,.0f}원
        """
        
        prompt = f"다음 판매 데이터를 간단히 분석해주세요 (한국어로 3-5문장): {summary}"
        
        response = model.generate_content(prompt)
        return response.text
        
    except Exception as e:
        return f"AI 분석 실패: {e}"

# --------------------------------------------------------------------------
# 간단한 대시보드
# --------------------------------------------------------------------------
def create_simple_dashboard(df_record):
    """간단한 대시보드 생성"""
    if df_record.empty:
        st.warning("📊 분석할 데이터가 없습니다.")
        return
    
    # 기본 메트릭
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_revenue = df_record['실결제금액'].sum()
        st.metric("총 매출", f"₩{total_revenue:,.0f}")
    
    with col2:
        total_orders = len(df_record)
        st.metric("총 주문 수", f"{total_orders:,}")
    
    with col3:
        avg_order = total_revenue / total_orders if total_orders > 0 else 0
        st.metric("평균 주문", f"₩{avg_order:,.0f}")
    
    with col4:
        unique_customers = df_record['수령자명'].nunique()
        st.metric("고객 수", f"{unique_customers:,}")
    
    # 간단한 차트
    if '쇼핑몰' in df_record.columns:
        mall_sales = df_record.groupby('쇼핑몰')['실결제금액'].sum().reset_index()
        fig = px.pie(mall_sales, values='실결제금액', names='쇼핑몰', title='쇼핑몰별 매출')
        st.plotly_chart(fig, use_container_width=True)

# --------------------------------------------------------------------------
# 기존 핵심 함수들 (유지)
# --------------------------------------------------------------------------
def to_excel_formatted(df, format_type=None):
    """데이터프레임을 서식이 적용된 엑셀 파일로 변환"""
    output = io.BytesIO()
    df_to_save = df.fillna('')
    
    if format_type == 'ecount_upload':
        df_to_save = df_to_save.rename(columns={'적요_전표': '적요', '적요_품목': '적요.1'})

    df_to_save.to_excel(output, index=False, sheet_name='Sheet1')
    
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    # 공통 서식
    center_alignment = Alignment(horizontal='center', vertical='center')
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = center_alignment

    # 열 너비 자동 조정
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
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    pink_fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")

    if format_type == 'packing_list':
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            for cell in row:
                cell.border = thin_border
        
        # 묶음번호 처리
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

    if format_type == 'quantity_summary':
        for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=sheet.max_row)):
            for cell in row:
                cell.border = thin_border
            if row_idx > 0 and row_idx % 2 != 0:
                for cell in row:
                    cell.fill = pink_fill
    
    final_output = io.BytesIO()
    workbook.save(final_output)
    final_output.seek(0)
    
    return final_output.getvalue()

def process_all_files(file1, file2, file3, df_master):
    """메인 처리 함수"""
    try:
        df_smartstore = pd.read_excel(file1)
        df_ecount_orig = pd.read_excel(file2)
        df_godomall = pd.read_excel(file3)

        # 주문일자 추출
        order_date = datetime.now().date()
        if '일자' in df_ecount_orig.columns:
            try:
                order_date = pd.to_datetime(df_ecount_orig['일자'].iloc[0], format='%Y%m%d', errors='coerce')
                if not pd.isna(order_date):
                    order_date = order_date.date()
            except:
                pass

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
        
        # 배송비 중복 계산 방지
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
        
        # 경고 메시지 수집
        warnings = []
        
        # 기본 처리
        df_final = df_ecount_orig.copy().rename(columns={'금액': '실결제금액'})
        
        # 스마트스토어 병합
        key_cols_smartstore = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = df_smartstore.rename(
            columns={'실결제금액': '수정될_금액_스토어'}
        )[key_cols_smartstore + ['수정될_금액_스토어']].drop_duplicates(
            subset=key_cols_smartstore, keep='first'
        )
        
        # 고도몰 병합
        key_cols_godomall = ['재고관리코드', '수취인 이름', '상품수량']
        godomall_prices = df_godomall[key_cols_godomall + ['수정될_금액_고도몰']].rename(
            columns={'수취인 이름': '수령자명', '상품수량': '주문수량'}
        ).drop_duplicates(subset=['재고관리코드', '수령자명', '주문수량'], keep='first')
        
        # 데이터 타입 통일
        for col in ['재고관리코드', '수령자명']:
            df_final[col] = df_final[col].astype(str).str.strip()
            smartstore_prices[col] = smartstore_prices[col].astype(str).str.strip()
            godomall_prices[col] = godomall_prices[col].astype(str).str.strip()
        
        df_final['주문수량'] = pd.to_numeric(df_final['주문수량'], errors='coerce').fillna(0).astype(int)
        smartstore_prices['주문수량'] = pd.to_numeric(smartstore_prices['주문수량'], errors='coerce').fillna(0).astype(int)
        godomall_prices['주문수량'] = pd.to_numeric(godomall_prices['주문수량'], errors='coerce').fillna(0).astype(int)
        
        df_final['실결제금액'] = pd.to_numeric(df_final['실결제금액'], errors='coerce').fillna(0).astype(int)
        
        # 데이터 병합
        df_final = pd.merge(df_final, smartstore_prices, on=key_cols_smartstore, how='left')
        df_final = pd.merge(df_final, godomall_prices, on=['재고관리코드', '수령자명', '주문수량'], how='left')
        
        # 최종 금액 업데이트
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
        
        df_main_result = df_final[['재고관리코드', 'SKU상품명', '주문수량', '실결제금액', '쇼핑몰', '수령자명', 'original_order']]
        
        # 수량 요약 및 포장 리스트
        df_quantity_summary = df_main_result.groupby('SKU상품명', as_index=False)['주문수량'].sum().rename(columns={'주문수량': '개수'})
        
        df_packing_list = df_main_result.sort_values(by='original_order')[['SKU상품명', '주문수량', '수령자명', '쇼핑몰']].copy()
        is_first_item = df_packing_list['수령자명'] != df_packing_list['수령자명'].shift(1)
        df_packing_list['묶음번호'] = is_first_item.cumsum()
        df_packing_list_final = df_packing_list.copy()
        df_packing_list_final['묶음번호'] = df_packing_list_final['묶음번호'].where(is_first_item, '')
        df_packing_list_final = df_packing_list_final[['묶음번호', 'SKU상품명', '주문수량', '수령자명', '쇼핑몰']]
        
        # 마스터 데이터 병합
        if df_master is not None and not df_master.empty:
            df_merged = pd.merge(df_main_result, df_master[['SKU코드', '과세여부', '입수량']], 
                                left_on='재고관리코드', right_on='SKU코드', how='left')
        else:
            df_merged = df_main_result.copy()
            df_merged['과세여부'] = '과세'
            df_merged['입수량'] = 1
            df_merged['SKU코드'] = df_merged['재고관리코드']
        
        # 거래처 매핑
        client_map = {
            '쿠팡': '쿠팡 주식회사', 
            '고도몰5': '고래미자사몰_현금영수증(고도몰)', 
            '스마트스토어': '스토어팜',
            '배민상회': '주식회사 우아한형제들(배민상회)',
            '이지웰몰': '주식회사 현대이지웰'
        }
        
        # 이카운트 업로드용 데이터
        df_ecount_upload = pd.DataFrame()
        
        df_ecount_upload['일자'] = datetime.now().strftime("%Y%m%d")
        df_ecount_upload['거래처명'] = df_merged['쇼핑몰'].map(client_map).fillna(df_merged['쇼핑몰'])
        df_ecount_upload['출하창고'] = '고래미'
        df_ecount_upload['거래유형'] = np.where(df_merged['과세여부'] == '면세', 12, 11)
        df_ecount_upload['적요_전표'] = '오전/온라인'
        df_ecount_upload['품목코드'] = df_merged['재고관리코드']
        
        # 수량 계산
        is_box_order = df_merged['SKU상품명'].str.contains("BOX", na=False)
        입수량 = pd.to_numeric(df_merged['입수량'], errors='coerce').fillna(1)
        base_quantity = np.where(is_box_order, df_merged['주문수량'] * 입수량, df_merged['주문수량'])
        is_3_pack = df_merged['SKU상품명'].str.contains("3개입|3개", na=False)
        final_quantity = np.where(is_3_pack, base_quantity * 3, base_quantity)
        df_ecount_upload['박스'] = np.where(is_box_order, df_merged['주문수량'], np.nan)
        df_ecount_upload['수량'] = final_quantity.astype(int)
        
        # 금액 계산
        df_merged['실결제금액'] = pd.to_numeric(df_merged['실결제금액'], errors='coerce').fillna(0)
        공급가액 = np.where(df_merged['과세여부'] == '과세', df_merged['실결제금액'] / 1.1, df_merged['실결제금액'])
        df_ecount_upload['공급가액'] = 공급가액
        df_ecount_upload['부가세'] = df_merged['실결제금액'] - df_ecount_upload['공급가액']
        
        df_ecount_upload['쇼핑몰고객명'] = df_merged['수령자명']
        df_ecount_upload['original_order'] = df_merged['original_order']
        
        # 이카운트 컬럼 정리
        ecount_columns = [
            '일자', '순번', '거래처코드', '거래처명', '담당자', '출하창고', '거래유형', '통화', '환율', 
            '적요_전표', '미수금', '총합계', '연결전표', '품목코드', '품목명', '규격', '박스', '수량', 
            '단가', '외화금액', '공급가액', '부가세', '적요_품목', '생산전표생성', '시리얼/로트', 
            '관리항목', '쇼핑몰고객명', 'original_order'
        ]
        
        for col in ecount_columns:
            if col not in df_ecount_upload:
                df_ecount_upload[col] = ''
        
        for col in ['공급가액', '부가세']:
            df_ecount_upload[col] = df_ecount_upload[col].round().astype('Int64')
        
        # 정렬
        sort_order = [
            '고래미자사몰_현금영수증(고도몰)', 
            '스토어팜', 
            '쿠팡 주식회사',
            '주식회사 우아한형제들(배민상회)',
            '주식회사 현대이지웰'
        ]
        
        df_ecount_upload['거래처명_sort'] = pd.Categorical(
            df_ecount_upload['거래처명'], 
            categories=sort_order, 
            ordered=True
        )
        
        df_ecount_upload = df_ecount_upload.sort_values(
            by=['거래처명_sort', '거래유형', 'original_order'],
            ascending=[True, True, True]
        ).drop(columns=['거래처명_sort', 'original_order'])
        
        df_ecount_upload = df_ecount_upload[ecount_columns[:-1]]
        
        # 기록용 데이터 준비
        df_for_record = df_merged.copy()
        df_for_record['주문일자'] = order_date
        df_for_record['거래처명'] = df_for_record['쇼핑몰'].map(client_map).fillna(df_for_record['쇼핑몰'])
        df_for_record['공급가액'] = 공급가액.round().astype('Int64')
        df_for_record['부가세'] = (df_merged['실결제금액'] - 공급가액).round().astype('Int64')
        df_for_record['거래유형'] = np.where(df_merged['과세여부'] == '면세', 12, 11)
        df_for_record['품목코드'] = df_merged['재고관리코드']

        return (df_main_result.drop(columns=['original_order']), 
                df_quantity_summary, 
                df_packing_list_final, 
                df_ecount_upload, 
                df_for_record,
                True, 
                "모든 파일 처리가 성공적으로 완료되었습니다.", 
                warnings)

    except Exception as e:
        import traceback
        st.error(f"처리 중 오류: {e}")
        st.error(traceback.format_exc())
        return None, None, None, None, None, False, f"오류: {e}", []

# --------------------------------------------------------------------------
# 메인 앱
# --------------------------------------------------------------------------
def main():
    # 사이드바
    with st.sidebar:
        st.title("⚙️ 설정")
        
        user_name = st.text_input("사용자 이름", value=st.session_state.get('user_name', ''))
        if user_name:
            st.session_state['user_name'] = user_name
        
        st.divider()
        st.info("""
        **v2.0 Lite**
        - 핵심 기능 최적화
        - 빠른 처리 속도
        - 안정적인 작동
        """)
    
    # 메인
    st.title("📑 주문 처리 자동화 시스템 v2.0 Lite")
    st.caption("빠르고 안정적인 주문 데이터 처리")
    
    # 탭
    tab1, tab2 = st.tabs(["📤 데이터 처리", "📊 간단 대시보드"])
    
    with tab1:
        st.header("1. 원본 파일 업로드")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            file1 = st.file_uploader("스마트스토어", type=['xlsx', 'xls'])
        with col2:
            file2 = st.file_uploader("이카운트", type=['xlsx', 'xls'])
        with col3:
            file3 = st.file_uploader("고도몰", type=['xlsx', 'xls'])
        
        st.divider()
        
        if st.button("🚀 처리 시작", type="primary", use_container_width=True):
            if file1 and file2 and file3:
                # 마스터 데이터 로드
                with st.spinner('마스터 데이터 로드 중...'):
                    df_master = load_master_from_sharepoint()
                
                if df_master is None:
                    st.error("마스터 데이터 로드 실패!")
                    return
                
                # 파일 처리
                with st.spinner('파일 처리 중...'):
                    result = process_all_files(file1, file2, file3, df_master)
                
                if result[5]:  # success
                    df_main, df_qty, df_pack, df_ecount, df_for_record, success, message, warnings = result
                    
                    st.success(message)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    
                    # 로컬 저장
                    saved, count = save_record_locally(df_for_record)
                    if saved:
                        st.info(f"✅ {count}건 저장 완료")
                    
                    # 경고
                    if warnings:
                        with st.expander("⚠️ 경고 메시지"):
                            for w in warnings:
                                st.write(w)
                    
                    # 결과 표시
                    st.subheader("📥 다운로드")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            "📥 이카운트 업로드용",
                            to_excel_formatted(df_ecount, 'ecount_upload'),
                            f"ecount_{timestamp}.xlsx"
                        )
                        st.download_button(
                            "📥 포장 리스트",
                            to_excel_formatted(df_pack, 'packing_list'),
                            f"packing_{timestamp}.xlsx"
                        )
                    
                    with col2:
                        st.download_button(
                            "📥 출고 수량",
                            to_excel_formatted(df_qty, 'quantity_summary'),
                            f"quantity_{timestamp}.xlsx"
                        )
                        st.download_button(
                            "📥 최종 보정 리스트",
                            to_excel_formatted(df_main),
                            f"final_{timestamp}.xlsx"
                        )
                    
                    # 미리보기
                    with st.expander("데이터 미리보기"):
                        st.dataframe(df_ecount.head(10))
                else:
                    st.error(result[6])
            else:
                st.warning("3개 파일을 모두 업로드하세요.")
    
    with tab2:
        st.header("📊 간단 대시보드")
        
        df_record = load_record_data()
        
        if not df_record.empty:
            create_simple_dashboard(df_record)
            
            # AI 분석 (옵션)
            if st.button("🤖 AI 분석"):
                with st.spinner("분석 중..."):
                    analysis = analyze_with_gemini(df_record, "trend")
                    st.write(analysis)
        else:
            st.info("아직 데이터가 없습니다.")

if __name__ == "__main__":
    main()
