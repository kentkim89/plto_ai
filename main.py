import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import requests
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import tempfile
import os
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import google.generativeai as genai
import json
from typing import Dict, List, Tuple, Any
import base64
from urllib.parse import urlparse, parse_qs

# --------------------------------------------------------------------------
# Streamlit 페이지 설정
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="주문 처리 자동화 시스템 v2.0", 
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'About': "주문 처리 자동화 시스템 v2.0 - SharePoint & AI Powered"
    }
)

# --------------------------------------------------------------------------
# SharePoint 연결 함수
# --------------------------------------------------------------------------
@st.cache_resource
def get_sharepoint_context():
    """SharePoint 인증 컨텍스트 생성"""
    try:
        site_url = f"https://goremi.sharepoint.com/sites/{st.secrets['sharepoint_files']['site_name']}"
        client_id = st.secrets["sharepoint"]["client_id"]
        client_secret = st.secrets["sharepoint"]["client_secret"]
        
        credentials = ClientCredential(client_id, client_secret)
        ctx = ClientContext(site_url).with_credentials(credentials)
        
        # 연결 테스트
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        
        return ctx
    except Exception as e:
        st.error(f"SharePoint 연결 오류: {e}")
        return None

def extract_file_path_from_url(sharepoint_url):
    """SharePoint URL에서 파일 경로 추출"""
    try:
        # URL에서 파일 ID 추출
        parsed = urlparse(sharepoint_url)
        
        # 직접 다운로드 URL 형식으로 변환
        if "sharepoint.com/:x:" in sharepoint_url:  # Excel 파일
            # 공유 링크를 다운로드 URL로 변환
            download_url = sharepoint_url.replace("/:x:/", "/_layouts/15/download.aspx?share=")
            download_url = download_url.split("?")[0] + "?share=" + sharepoint_url.split("/")[-1].split("?")[0]
            return download_url
        else:
            return sharepoint_url
    except Exception as e:
        st.error(f"URL 파싱 오류: {e}")
        return None

@st.cache_data(ttl=600)  # 10분 캐시
def load_master_from_sharepoint():
    """SharePoint에서 마스터 데이터 로드"""
    try:
        ctx = get_sharepoint_context()
        if not ctx:
            raise Exception("SharePoint 컨텍스트 생성 실패")
        
        # 파일 경로 설정
        site_name = st.secrets["sharepoint_files"]["site_name"]
        file_name = st.secrets["sharepoint_files"]["file_name"]
        
        # 상대 URL 구성
        file_url = f"/sites/{site_name}/Shared Documents/{file_name}"
        
        try:
            # SharePoint에서 파일 다운로드
            response = File.open_binary(ctx, file_url)
            
            # BytesIO 객체로 변환
            bytes_file_obj = io.BytesIO()
            bytes_file_obj.write(response.content)
            bytes_file_obj.seek(0)
            
            # Excel 파일 읽기
            df_master = pd.read_excel(bytes_file_obj)
            df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
            
            st.success(f"✅ 마스터 데이터 로드 완료: {len(df_master)}개 품목")
            return df_master
            
        except Exception as e:
            st.warning(f"SharePoint 파일 접근 실패: {e}")
            
            # 대체 방법: 직접 URL 접근
            master_url = st.secrets["sharepoint_files"]["plto_master_data_file_url"]
            response = requests.get(master_url)
            if response.status_code == 200:
                df_master = pd.read_excel(io.BytesIO(response.content))
                df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
                st.success(f"✅ 마스터 데이터 로드 완료 (직접 URL): {len(df_master)}개 품목")
                return df_master
            else:
                raise Exception(f"파일 다운로드 실패: {response.status_code}")
                
    except Exception as e:
        st.error(f"마스터 데이터 로드 실패: {e}")
        
        # 최종 폴백: 로컬 파일
        try:
            if os.path.exists("master_data.csv"):
                df_master = pd.read_csv("master_data.csv")
                df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
                st.warning(f"⚠️ 로컬 백업 파일 사용: {len(df_master)}개 품목")
                return df_master
        except:
            pass
            
        return None

def load_record_data_from_sharepoint():
    """SharePoint에서 기록 데이터 로드"""
    try:
        ctx = get_sharepoint_context()
        if not ctx:
            return pd.DataFrame()
        
        site_name = st.secrets["sharepoint_files"]["site_name"]
        record_file_name = st.secrets["sharepoint_files"].get("record_file_name", "plto_record_data.xlsx")
        
        file_url = f"/sites/{site_name}/Shared Documents/{record_file_name}"
        
        try:
            response = File.open_binary(ctx, file_url)
            bytes_file_obj = io.BytesIO()
            bytes_file_obj.write(response.content)
            bytes_file_obj.seek(0)
            
            df_record = pd.read_excel(bytes_file_obj)
            return df_record
        except Exception as e:
            # 파일이 없거나 비어있으면 빈 DataFrame 반환
            st.info(f"기록 파일이 없거나 비어있습니다. 새로 생성됩니다.")
            return pd.DataFrame()
            
    except Exception as e:
        st.error(f"기록 데이터 로드 실패: {e}")
        return pd.DataFrame()

def save_record_to_sharepoint(df_new_records):
    """SharePoint에 기록 데이터 저장"""
    try:
        ctx = get_sharepoint_context()
        if not ctx:
            return False, 0
        
        site_name = st.secrets["sharepoint_files"]["site_name"]
        record_file_name = st.secrets["sharepoint_files"].get("record_file_name", "plto_record_data.xlsx")
        
        # 기존 데이터 로드
        df_existing = load_record_data_from_sharepoint()
        
        # 헤더 정의
        expected_columns = [
            '처리일시', '주문일자', '쇼핑몰', '거래처명', '품목코드', 'SKU상품명', 
            '주문수량', '실결제금액', '공급가액', '부가세', '수령자명', 
            '과세여부', '거래유형', '처리자', '원본파일명'
        ]
        
        # 기존 데이터가 비어있으면 헤더 설정
        if df_existing.empty:
            df_existing = pd.DataFrame(columns=expected_columns)
        
        # 새 레코드 준비
        df_new_records['처리일시'] = datetime.now()
        df_new_records['처리자'] = st.session_state.get('user_name', 'Unknown')
        df_new_records['원본파일명'] = st.session_state.get('current_files', '')
        
        # 중복 체크 (주문일자, 쇼핑몰, 수령자명, 품목코드 기준)
        if not df_existing.empty:
            merge_keys = ['주문일자', '쇼핑몰', '수령자명', '품목코드']
            
            # 각 컬럼이 존재하는지 확인
            for key in merge_keys:
                if key not in df_existing.columns:
                    df_existing[key] = ''
                if key not in df_new_records.columns:
                    df_new_records[key] = ''
            
            df_existing['check_key'] = df_existing[merge_keys].astype(str).agg('_'.join, axis=1)
            df_new_records['check_key'] = df_new_records[merge_keys].astype(str).agg('_'.join, axis=1)
            
            # 중복되지 않는 레코드만 필터링
            new_keys = set(df_new_records['check_key']) - set(df_existing['check_key'])
            df_new_records = df_new_records[df_new_records['check_key'].isin(new_keys)]
            
            # check_key 컬럼 제거
            df_existing = df_existing.drop('check_key', axis=1, errors='ignore')
            df_new_records = df_new_records.drop('check_key', axis=1, errors='ignore')
        
        # 누락된 컬럼 추가
        for col in expected_columns:
            if col not in df_new_records.columns:
                df_new_records[col] = ''
            if col not in df_existing.columns:
                df_existing[col] = ''
        
        # 데이터 결합
        df_combined = pd.concat([df_existing, df_new_records], ignore_index=True)
        
        # 날짜 형식 정리
        if '처리일시' in df_combined.columns:
            df_combined['처리일시'] = pd.to_datetime(df_combined['처리일시'], errors='coerce')
        if '주문일자' in df_combined.columns:
            df_combined['주문일자'] = pd.to_datetime(df_combined['주문일자'], errors='coerce')
        
        # 정렬
        df_combined = df_combined.sort_values('처리일시', ascending=False)
        
        # 임시 파일로 저장
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
                df_combined.to_excel(writer, index=False, sheet_name='Records')
                
                # 워크시트 가져오기
                worksheet = writer.sheets['Records']
                
                # 헤더 서식 적용
                header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                header_font = openpyxl.styles.Font(color="FFFFFF", bold=True)
                
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                
                # 열 너비 자동 조정
                for column_cells in worksheet.columns:
                    length = max(len(str(cell.value or '')) for cell in column_cells)
                    worksheet.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 50)
            
            tmp.flush()
            
            # SharePoint에 업로드
            with open(tmp.name, 'rb') as file_content:
                file_url = f"/sites/{site_name}/Shared Documents/{record_file_name}"
                
                folder = ctx.web.get_folder_by_server_relative_path(f"sites/{site_name}/Shared Documents")
                target_file = folder.upload_file(record_file_name, file_content.read())
                ctx.execute_query()
            
            os.unlink(tmp.name)
        
        return True, len(df_new_records)
        
    except Exception as e:
        st.error(f"기록 저장 실패: {e}")
        import traceback
        st.error(traceback.format_exc())
        return False, 0

# --------------------------------------------------------------------------
# Gemini AI 분석 함수
# --------------------------------------------------------------------------
def initialize_gemini():
    """Gemini AI 초기화"""
    try:
        genai.configure(api_key=st.secrets["gemini"]["api_key"])
        model = genai.GenerativeModel('gemini-pro')
        return model
    except Exception as e:
        st.error(f"Gemini AI 초기화 실패: {e}")
        return None

def analyze_with_gemini(df_data, analysis_type="trend"):
    """Gemini AI를 사용한 데이터 분석"""
    model = initialize_gemini()
    if not model or df_data.empty:
        return None
    
    try:
        # 데이터 요약 생성
        summary_stats = df_data.describe().to_string()
        
        # 상위 품목 정보
        top_products = df_data.groupby('SKU상품명')['실결제금액'].sum().nlargest(10).to_string()
        
        # 쇼핑몰별 매출
        mall_sales = df_data.groupby('쇼핑몰')['실결제금액'].sum().to_string()
        
        # 분석 유형별 프롬프트
        prompts = {
            "trend": f"""
            다음 판매 데이터를 분석하여 주요 트렌드와 인사이트를 한국어로 제공해주세요:
            
            [통계 요약]
            {summary_stats}
            
            [베스트셀러 TOP 10]
            {top_products}
            
            [쇼핑몰별 매출]
            {mall_sales}
            
            다음 항목들을 포함해주세요:
            1. 📈 전체적인 판매 트렌드
            2. 🏆 베스트셀러 상품 분석
            3. 🛍️ 쇼핑몰별 특징
            4. 📊 계절성 또는 주기적 패턴
            5. 💡 개선 제안사항
            
            각 항목을 명확하게 구분하고, 이모지를 활용하여 읽기 쉽게 작성해주세요.
            """,
            
            "forecast": f"""
            다음 판매 데이터를 바탕으로 향후 예측을 한국어로 제공해주세요:
            
            [통계 요약]
            {summary_stats}
            
            [상위 품목]
            {top_products}
            
            다음을 포함해주세요:
            1. 📅 다음 주/월 예상 판매량
            2. 📦 주의가 필요한 재고 품목
            3. 🚀 성장 가능성이 높은 카테고리
            4. ⚠️ 리스크 요인
            
            구체적인 수치와 함께 실행 가능한 제안을 해주세요.
            """,
            
            "anomaly": f"""
            다음 판매 데이터에서 이상 패턴이나 특이사항을 한국어로 찾아주세요:
            
            [통계 요약]
            {summary_stats}
            
            [쇼핑몰별 매출]
            {mall_sales}
            
            다음을 확인해주세요:
            1. 🔍 비정상적인 주문 패턴
            2. 📉 급격한 변화가 있는 품목
            3. ⚡ 주의가 필요한 거래
            4. 🔧 데이터 품질 이슈
            
            발견된 이상 항목에 대한 대응 방안도 제시해주세요.
            """
        }
        
        prompt = prompts.get(analysis_type, prompts["trend"])
        response = model.generate_content(prompt)
        
        return response.text
        
    except Exception as e:
        st.error(f"AI 분석 실패: {e}")
        return None

# --------------------------------------------------------------------------
# 데이터 시각화 함수
# --------------------------------------------------------------------------
def create_dashboard(df_record):
    """대시보드 생성"""
    if df_record.empty:
        st.warning("📊 분석할 데이터가 없습니다.")
        return
    
    # 데이터 전처리
    df_record['주문일자'] = pd.to_datetime(df_record['주문일자'], errors='coerce')
    df_record = df_record.dropna(subset=['주문일자'])
    
    # 색상 팔레트
    colors = px.colors.qualitative.Set3
    
    # 메트릭 계산
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_revenue = df_record['실결제금액'].sum()
        st.metric("총 매출", f"₩{total_revenue:,.0f}")
    
    with col2:
        total_orders = len(df_record)
        st.metric("총 주문 수", f"{total_orders:,}")
    
    with col3:
        avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
        st.metric("평균 주문 금액", f"₩{avg_order_value:,.0f}")
    
    with col4:
        unique_customers = df_record['수령자명'].nunique()
        st.metric("고유 고객 수", f"{unique_customers:,}")
    
    # 차트 생성
    tab1, tab2, tab3, tab4 = st.tabs(["📈 일별 트렌드", "🏪 쇼핑몰별 분석", "📦 상품별 분석", "🤖 AI 인사이트"])
    
    with tab1:
        col1, col2 = st.columns(2)
        
        with col1:
            # 일별 매출 트렌드
            daily_sales = df_record.groupby(df_record['주문일자'].dt.date)['실결제금액'].sum().reset_index()
            
            fig = px.line(daily_sales, x='주문일자', y='실결제금액',
                         title='일별 매출 트렌드',
                         labels={'실결제금액': '매출 (원)', '주문일자': '날짜'},
                         color_discrete_sequence=[colors[0]])
            fig.update_layout(hovermode='x unified', height=400)
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # 주별 매출 트렌드
            df_record['주차'] = df_record['주문일자'].dt.isocalendar().week
            df_record['연도'] = df_record['주문일자'].dt.year
            weekly_sales = df_record.groupby(['연도', '주차'])['실결제금액'].sum().reset_index()
            weekly_sales['연도_주차'] = weekly_sales['연도'].astype(str) + '-W' + weekly_sales['주차'].astype(str).str.zfill(2)
            
            fig2 = px.bar(weekly_sales, x='연도_주차', y='실결제금액',
                         title='주별 매출 트렌드',
                         labels={'실결제금액': '매출 (원)', '연도_주차': '연도-주차'},
                         color_discrete_sequence=[colors[1]])
            fig2.update_layout(height=400)
            st.plotly_chart(fig2, use_container_width=True)
        
        # 시간대별 분석
        if '처리일시' in df_record.columns:
            df_record['처리시간'] = pd.to_datetime(df_record['처리일시'], errors='coerce').dt.hour
            hourly_orders = df_record.groupby('처리시간').size().reset_index(name='주문수')
            
            fig3 = px.bar(hourly_orders, x='처리시간', y='주문수',
                         title='시간대별 주문 분포',
                         labels={'처리시간': '시간', '주문수': '주문 건수'},
                         color_discrete_sequence=[colors[2]])
            fig3.update_layout(height=300)
            st.plotly_chart(fig3, use_container_width=True)
    
    with tab2:
        col1, col2 = st.columns(2)
        
        with col1:
            # 쇼핑몰별 매출 비중
            mall_sales = df_record.groupby('쇼핑몰')['실결제금액'].sum().reset_index()
            
            fig = px.pie(mall_sales, values='실결제금액', names='쇼핑몰',
                        title='쇼핑몰별 매출 비중',
                        color_discrete_sequence=colors)
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # 쇼핑몰별 평균 주문 금액
            mall_avg = df_record.groupby('쇼핑몰')['실결제금액'].mean().reset_index()
            mall_avg.columns = ['쇼핑몰', '평균주문금액']
            
            fig2 = px.bar(mall_avg, x='쇼핑몰', y='평균주문금액',
                          title='쇼핑몰별 평균 주문 금액',
                          labels={'평균주문금액': '평균 금액 (원)'},
                          color='쇼핑몰',
                          color_discrete_sequence=colors)
            fig2.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig2, use_container_width=True)
        
        # 쇼핑몰별 일별 트렌드
        mall_daily = df_record.groupby([df_record['주문일자'].dt.date, '쇼핑몰'])['실결제금액'].sum().reset_index()
        
        fig3 = px.line(mall_daily, x='주문일자', y='실결제금액', color='쇼핑몰',
                      title='쇼핑몰별 일별 매출 트렌드',
                      color_discrete_sequence=colors)
        fig3.update_layout(height=400)
        st.plotly_chart(fig3, use_container_width=True)
    
    with tab3:
        col1, col2 = st.columns(2)
        
        with col1:
            # TOP 10 상품 (매출 기준)
            product_sales = df_record.groupby('SKU상품명').agg({
                '주문수량': 'sum',
                '실결제금액': 'sum'
            }).reset_index()
            
            top_products = product_sales.nlargest(10, '실결제금액')
            
            fig = px.bar(top_products, x='실결제금액', y='SKU상품명',
                         orientation='h', title='TOP 10 베스트셀러 (매출 기준)',
                         labels={'실결제금액': '매출 (원)', 'SKU상품명': '상품명'},
                         color='실결제금액',
                         color_continuous_scale='Blues')
            fig.update_layout(yaxis={'categoryorder':'total ascending'}, height=500)
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # TOP 10 상품 (수량 기준)
            top_qty = product_sales.nlargest(10, '주문수량')
            
            fig2 = px.bar(top_qty, x='주문수량', y='SKU상품명',
                          orientation='h', title='TOP 10 베스트셀러 (수량 기준)',
                          labels={'주문수량': '판매 수량', 'SKU상품명': '상품명'},
                          color='주문수량',
                          color_continuous_scale='Greens')
            fig2.update_layout(yaxis={'categoryorder':'total ascending'}, height=500)
            st.plotly_chart(fig2, use_container_width=True)
        
        # 상품별 판매 트리맵
        fig3 = px.treemap(product_sales.nlargest(20, '주문수량'), 
                         path=['SKU상품명'], values='주문수량',
                         title='상품별 판매 수량 분포 (TOP 20)',
                         color='실결제금액',
                         color_continuous_scale='RdYlBu')
        fig3.update_layout(height=500)
        st.plotly_chart(fig3, use_container_width=True)
    
    with tab4:
        st.subheader("🤖 AI 기반 데이터 분석")
        
        col1, col2 = st.columns([1, 3])
        
        with col1:
            analysis_type = st.selectbox(
                "분석 유형 선택",
                ["trend", "forecast", "anomaly"],
                format_func=lambda x: {
                    "trend": "📈 트렌드 분석",
                    "forecast": "🔮 판매 예측",
                    "anomaly": "⚠️ 이상 패턴 감지"
                }[x]
            )
            
            if st.button("🚀 AI 분석 실행", type="primary", use_container_width=True):
                st.session_state['run_analysis'] = True
        
        with col2:
            if st.session_state.get('run_analysis', False):
                with st.spinner("AI가 데이터를 분석 중입니다... ⏳"):
                    analysis_result = analyze_with_gemini(df_record, analysis_type)
                    if analysis_result:
                        st.markdown("### 📊 분석 결과")
                        st.markdown(analysis_result)
                        st.session_state['run_analysis'] = False
                    else:
                        st.error("AI 분석을 수행할 수 없습니다.")
                        st.session_state['run_analysis'] = False

# --------------------------------------------------------------------------
# 기존 함수들 (수정 및 유지)
# --------------------------------------------------------------------------
def to_excel_formatted(df, format_type=None):
    """데이터프레임을 서식이 적용된 엑셀 파일 형식의 BytesIO 객체로 변환하는 함수"""
    output = io.BytesIO()
    df_to_save = df.fillna('')
    
    if format_type == 'ecount_upload':
        df_to_save = df_to_save.rename(columns={'적요_전표': '적요', '적요_품목': '적요.1'})

    df_to_save.to_excel(output, index=False, sheet_name='Sheet1')
    
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    # 공통 서식: 모든 셀 가운데 정렬
    center_alignment = Alignment(horizontal='center', vertical='center')
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = center_alignment

    # 파일별 특수 서식
    for column_cells in sheet.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            try:
                if cell.value:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        sheet.column_dimensions[column].width = adjusted_width
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    pink_fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")

    if format_type == 'packing_list':
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                cell.border = thin_border
        
        bundle_start_row = 2
        for row_num in range(2, sheet.max_row + 2):
            current_bundle_cell = sheet.cell(row=row_num, column=1) if row_num <= sheet.max_row else None
            
            if (current_bundle_cell and current_bundle_cell.value) or (row_num > sheet.max_row):
                if row_num > 2:
                    bundle_end_row = row_num - 1
                    prev_bundle_num_str = str(sheet.cell(row=bundle_start_row, column=1).value)
                    
                    if prev_bundle_num_str.isdigit():
                        prev_bundle_num = int(prev_bundle_num_str)
                        if prev_bundle_num % 2 != 0:
                            for r in range(bundle_start_row, bundle_end_row + 1):
                                for c in range(1, sheet.max_column + 1):
                                    sheet.cell(row=r, column=c).fill = pink_fill
                    
                    if bundle_start_row < bundle_end_row:
                        sheet.merge_cells(start_row=bundle_start_row, start_column=1, end_row=bundle_end_row, end_column=1)
                        sheet.merge_cells(start_row=bundle_start_row, start_column=4, end_row=bundle_end_row, end_column=4)
                
                bundle_start_row = row_num

    if format_type == 'quantity_summary':
        for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column)):
            for cell in row:
                cell.border = thin_border
            if row_idx > 0 and row_idx % 2 != 0:
                for cell in row:
                    cell.fill = pink_fill
    
    final_output = io.BytesIO()
    workbook.save(final_output)
    final_output.seek(0)
    
    return final_output.getvalue()

def extract_order_date(df_ecount):
    """이카운트 데이터에서 주문일자 추출"""
    try:
        # 일자 컬럼이 있으면 사용, 없으면 오늘 날짜 사용
        if '일자' in df_ecount.columns:
            order_date = pd.to_datetime(df_ecount['일자'].iloc[0], format='%Y%m%d', errors='coerce')
            if pd.isna(order_date):
                order_date = datetime.now().date()
            else:
                order_date = order_date.date()
        else:
            order_date = datetime.now().date()
        return order_date
    except:
        return datetime.now().date()

def process_all_files(file1, file2, file3, df_master):
    try:
        df_smartstore = pd.read_excel(file1)
        df_ecount_orig = pd.read_excel(file2)
        df_godomall = pd.read_excel(file3)

        # 파일명 저장
        st.session_state['current_files'] = f"{file1.name}, {file2.name}, {file3.name}"

        # 주문일자 추출
        order_date = extract_order_date(df_ecount_orig)

        df_ecount_orig['original_order'] = range(len(df_ecount_orig))
        
        # 컬럼명 호환성 처리
        if '회 할인 금액' in df_godomall.columns and '회원 할인 금액' not in df_godomall.columns:
            df_godomall.rename(columns={'회 할인 금액': '회원 할인 금액'}, inplace=True)
        if '자체옵션코드' in df_godomall.columns:
            df_godomall.rename(columns={'자체옵션코드': '재고관리코드'}, inplace=True)
        
        # 1단계: 데이터 클리닝 강화
        cols_to_numeric = ['상품별 품목금액', '총 배송 금액', '회원 할인 금액', '쿠폰 할인 금액', '사용된 마일리지', '총 결제 금액']
        for col in cols_to_numeric:
            if col in df_godomall.columns: 
                df_godomall[col] = pd.to_numeric(df_godomall[col].astype(str).str.replace('[원,]', '', regex=True), errors='coerce').fillna(0)
        
        # 2단계: 배송비 중복 계산 방지
        df_godomall['보정된_배송비'] = np.where(
            df_godomall.duplicated(subset=['수취인 이름']), 
            0, 
            df_godomall['총 배송 금액']
        )
        
        df_godomall['수정될_금액_고도몰'] = (
            df_godomall['상품별 품목금액'] + df_godomall['보정된_배송비'] - df_godomall['회원 할인 금액'] - 
            df_godomall['쿠폰 할인 금액'] - df_godomall['사용된 마일리지']
        )
        
        # 3단계: 결제 금액 검증 및 알림 기능 추가
        godomall_warnings = []
        grouped_godomall = df_godomall.groupby('수취인 이름')
        
        for name, group in grouped_godomall:
            calculated_total = group['수정될_금액_고도몰'].sum()
            actual_total = group['총 결제 금액'].iloc[0]
            discrepancy = calculated_total - actual_total
            
            if abs(discrepancy) > 1:
                warning_msg = f"- [고도몰 금액 불일치] **{name}**님의 주문의 계산된 금액과 실제 결제 금액이 **{discrepancy:,.0f}원** 만큼 차이납니다."
                godomall_warnings.append(warning_msg)

        # 기존 처리 로직 계속...
        df_final = df_ecount_orig.copy().rename(columns={'금액': '실결제금액'})
        
        # 스마트스토어 병합 준비
        key_cols_smartstore = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = df_smartstore.rename(columns={'실결제금액': '수정될_금액_스토어'})[key_cols_smartstore + ['수정될_금액_스토어']].drop_duplicates(subset=key_cols_smartstore, keep='first')
        
        # 고도몰 병합
        key_cols_godomall = ['재고관리코드', '수취인 이름', '상품수량']
        godomall_prices_for_merge = df_godomall[key_cols_godomall + ['수정될_금액_고도몰']].rename(
            columns={'수취인 이름': '수령자명', '상품수량': '주문수량'}
        )
        godomall_prices_for_merge = godomall_prices_for_merge.drop_duplicates(
            subset=['재고관리코드', '수령자명', '주문수량'], keep='first'
        )
        
        # 데이터 타입 통일
        for col in ['재고관리코드', '수령자명']:
            df_final[col] = df_final[col].astype(str).str.strip()
            smartstore_prices[col] = smartstore_prices[col].astype(str).str.strip()
            godomall_prices_for_merge[col] = godomall_prices_for_merge[col].astype(str).str.strip()
        
        for col in ['주문수량']:
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0).astype(int)
            smartstore_prices[col] = pd.to_numeric(smartstore_prices[col], errors='coerce').fillna(0).astype(int)
            godomall_prices_for_merge[col] = pd.to_numeric(godomall_prices_for_merge[col], errors='coerce').fillna(0).astype(int)

        df_final['실결제금액'] = pd.to_numeric(df_final['실결제금액'], errors='coerce').fillna(0).astype(int)
        
        # 데이터 병합
        df_final = pd.merge(df_final, smartstore_prices, on=key_cols_smartstore, how='left')
        df_final = pd.merge(df_final, godomall_prices_for_merge, 
                            on=['재고관리코드', '수령자명', '주문수량'], 
                            how='left')

        # 경고 메시지 생성
        warnings = [f"- [금액보정 실패] **{row['쇼핑몰']}** / {row['수령자명']} / {row['SKU상품명']}" 
                   for _, row in df_final[(df_final['쇼핑몰'] == '스마트스토어') & (df_final['수정될_금액_스토어'].isna()) | 
                                          (df_final['쇼핑몰'] == '고도몰5') & (df_final['수정될_금액_고도몰'].isna())].iterrows()]
        warnings.extend(godomall_warnings)

        # 최종 결제 금액 업데이트
        df_final['실결제금액'] = np.where(df_final['쇼핑몰'] == '고도몰5', 
                                          df_final['수정될_금액_고도몰'].fillna(df_final['실결제금액']), 
                                          df_final['실결제금액'])
        df_final['실결제금액'] = np.where(df_final['쇼핑몰'] == '스마트스토어', 
                                          df_final['수정될_금액_스토어'].fillna(df_final['실결제금액']), 
                                          df_final['실결제금액'])
        
        df_main_result = df_final[['재고관리코드', 'SKU상품명', '주문수량', '실결제금액', '쇼핑몰', '수령자명', 'original_order']]
        
        # 동명이인 경고 추가
        homonym_warnings = []
        name_groups = df_main_result.groupby('수령자명')['original_order'].apply(list)
        for name, orders in name_groups.items():
            if len(orders) > 1 and (max(orders) - min(orders) + 1) != len(orders):
                homonym_warnings.append(f"- [동명이인 의심] **{name}** 님의 주문이 떨어져서 입력되었습니다.")
        warnings.extend(homonym_warnings)

        # 수량 요약 및 포장 리스트 생성
        df_quantity_summary = df_main_result.groupby('SKU상품명', as_index=False)['주문수량'].sum().rename(columns={'주문수량': '개수'})
        df_packing_list = df_main_result.sort_values(by='original_order')[['SKU상품명', '주문수량', '수령자명', '쇼핑몰']].copy()
        is_first_item = df_packing_list['수령자명'] != df_packing_list['수령자명'].shift(1)
        df_packing_list['묶음번호'] = is_first_item.cumsum()
        df_packing_list_final = df_packing_list.copy()
        df_packing_list_final['묶음번호'] = df_packing_list_final['묶음번호'].where(is_first_item, '')
        df_packing_list_final = df_packing_list_final[['묶음번호', 'SKU상품명', '주문수량', '수령자명', '쇼핑몰']]

        # 마스터 데이터 병합
        df_merged = pd.merge(df_main_result, df_master[['SKU코드', '과세여부', '입수량']], 
                            left_on='재고관리코드', right_on='SKU코드', how='left')
        
        unmastered = df_merged[df_merged['SKU코드'].isna()]
        for _, row in unmastered.iterrows():
            warnings.append(f"- [미등록 상품] **{row['재고관리코드']}** / {row['SKU상품명']}")

        # 거래처 매핑
        client_map = {
            '쿠팡': '쿠팡 주식회사', 
            '고도몰5': '고래미자사몰_현금영수증(고도몰)', 
            '스마트스토어': '스토어팜',
            '배민상회': '주식회사 우아한형제들(배민상회)',
            '이지웰몰': '주식회사 현대이지웰'
        }
        
        # 이카운트 업로드용 데이터 생성
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
        
        df_ecount_upload['거래유형'] = pd.to_numeric(df_ecount_upload['거래유형'])
        
        # 정렬
        sort_order = [
            '고래미자사몰_현금영수증(고도몰)', 
            '스토어팜', 
            '쿠팡 주식회사',
            '주식회사 우아한형제들(배민상회)',
            '주식회사 현대이지웰'
        ]
        
        df_ecount_upload['거래처명_sort'] = pd.Categorical(df_ecount_upload['거래처명'], categories=sort_order, ordered=True)
        
        df_ecount_upload = df_ecount_upload.sort_values(
            by=['거래처명_sort', '거래유형', 'original_order'],
            ascending=[True, True, True]
        ).drop(columns=['거래처명_sort', 'original_order'])
        
        df_ecount_upload = df_ecount_upload[ecount_columns[:-1]]

        # SharePoint 기록용 데이터 준비
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
        st.error(f"처리 중 오류가 발생했습니다: {e}")
        st.error(traceback.format_exc())
        return None, None, None, None, None, False, f"오류가 발생했습니다: {e}", []

# --------------------------------------------------------------------------
# Streamlit 메인 앱
# --------------------------------------------------------------------------
def main():
    # 사이드바 설정
    with st.sidebar:
        st.title("⚙️ 설정")
        
        # 사용자 이름 입력
        user_name = st.text_input("사용자 이름", value=st.session_state.get('user_name', ''))
        if user_name:
            st.session_state['user_name'] = user_name
        
        st.divider()
        
        # SharePoint 연결 상태
        st.subheader("📡 연결 상태")
        ctx = get_sharepoint_context()
        if ctx:
            st.success("✅ SharePoint 연결됨")
        else:
            st.error("❌ SharePoint 연결 실패")
        
        # AI 상태
        if st.secrets.get("gemini", {}).get("api_key"):
            st.success("✅ Gemini AI 활성화")
        else:
            st.warning("⚠️ AI 기능 비활성화")
        
        st.divider()
        
        # 정보
        st.info("""
        **v2.0 새로운 기능:**
        - SharePoint 통합
        - AI 기반 분석
        - 실시간 대시보드
        - 자동 데이터 기록
        
        **연결 정보:**
        - Site: goremi.sharepoint.com
        - 데이터: plto_master_data.xlsx
        - 기록: plto_record_data.xlsx
        """)
    
    # 메인 컨텐츠
    st.title("📑 주문 처리 자동화 시스템 v2.0")
    st.caption("SharePoint & AI Powered | 실시간 데이터 분석 | 자동 기록 시스템")
    
    # 탭 구성
    tab1, tab2, tab3 = st.tabs(["📤 데이터 처리", "📊 대시보드", "📈 상세 분석"])
    
    with tab1:
        st.header("1. 원본 엑셀 파일 3개 업로드")
        col1, col2, col3 = st.columns(3)
        with col1:
            file1 = st.file_uploader("1️⃣ 스마트스토어 (금액확인용)", type=['xlsx', 'xls', 'csv'])
        with col2:
            file2 = st.file_uploader("2️⃣ 이카운트 다운로드 (주문목록)", type=['xlsx', 'xls', 'csv'])
        with col3:
            file3 = st.file_uploader("3️⃣ 고도몰 (금액확인용)", type=['xlsx', 'xls', 'csv'])

        st.divider()
        
        st.header("2. 처리 결과 확인 및 다운로드")
        
        if st.button("🚀 모든 데이터 처리 및 파일 생성 실행", type="primary", use_container_width=True):
            if file1 and file2 and file3:
                # 마스터 데이터 로드
                with st.spinner('SharePoint에서 마스터 데이터를 불러오는 중...'):
                    df_master = load_master_from_sharepoint()
                    
                if df_master is None:
                    st.error("🚨 마스터 데이터를 불러올 수 없습니다!")
                    return
                
                # 파일 처리
                with st.spinner('모든 파일을 읽고 데이터를 처리하며 엑셀 서식을 적용 중입니다...'):
                    result = process_all_files(file1, file2, file3, df_master)
                    
                if result[5]:  # success
                    df_main, df_qty, df_pack, df_ecount, df_for_record, success, message, warnings = result
                    
                    st.success(message)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                    # SharePoint에 기록 저장
                    with st.spinner('SharePoint에 처리 결과를 기록 중...'):
                        saved, record_count = save_record_to_sharepoint(df_for_record)
                        if saved:
                            st.success(f"✅ SharePoint에 {record_count}건의 신규 데이터가 저장되었습니다.")
                        else:
                            st.warning("⚠️ SharePoint 저장 실패 - 로컬 백업을 권장합니다.")

                    # 경고 메시지 표시
                    if warnings:
                        st.warning("⚠️ 확인 필요 항목")
                        with st.expander("자세한 목록 보기..."):
                            for warning_message in warnings:
                                st.markdown(warning_message)
                    
                    # 결과 탭
                    tab_erp, tab_pack, tab_qty, tab_main = st.tabs([
                        "🏢 이카운트 업로드용", 
                        "📋 포장 리스트", 
                        "📦 출고수량 요약", 
                        "✅ 최종 보정 리스트"
                    ])
                    
                    with tab_erp:
                        st.dataframe(df_ecount.astype(str), use_container_width=True)
                        st.download_button(
                            "📥 다운로드", 
                            to_excel_formatted(df_ecount, format_type='ecount_upload'), 
                            f"이카운트_업로드용_{timestamp}.xlsx"
                        )

                    with tab_pack:
                        st.dataframe(df_pack, use_container_width=True)
                        st.download_button(
                            "📥 다운로드", 
                            to_excel_formatted(df_pack, format_type='packing_list'), 
                            f"물류팀_전달용_포장리스트_{timestamp}.xlsx"
                        )

                    with tab_qty:
                        st.dataframe(df_qty, use_container_width=True)
                        st.download_button(
                            "📥 다운로드", 
                            to_excel_formatted(df_qty, format_type='quantity_summary'), 
                            f"물류팀_전달용_출고수량_{timestamp}.xlsx"
                        )
                    
                    with tab_main:
                        st.dataframe(df_main, use_container_width=True)
                        st.download_button(
                            "📥 다운로드", 
                            to_excel_formatted(df_main), 
                            f"최종_실결제금액_보정완료_{timestamp}.xlsx"
                        )
                else:
                    st.error(result[6])  # error message
            else:
                st.warning("⚠️ 3개의 엑셀 파일을 모두 업로드해야 실행할 수 있습니다.")
    
    with tab2:
        st.header("📊 실시간 대시보드")
        
        # 데이터 로드
        with st.spinner("SharePoint에서 기록 데이터를 불러오는 중..."):
            df_record = load_record_data_from_sharepoint()
        
        if not df_record.empty:
            create_dashboard(df_record)
        else:
            st.info("📊 아직 기록된 데이터가 없습니다. 데이터를 처리하면 자동으로 대시보드가 생성됩니다.")
    
    with tab3:
        st.header("📈 상세 데이터 분석")
        
        with st.spinner("SharePoint에서 기록 데이터를 불러오는 중..."):
            df_record = load_record_data_from_sharepoint()
        
        if not df_record.empty:
            # 필터링 옵션
            col1, col2, col3 = st.columns(3)
            
            with col1:
                date_range = st.date_input(
                    "날짜 범위",
                    value=(datetime.now() - timedelta(days=30), datetime.now()),
                    format="YYYY-MM-DD"
                )
            
            with col2:
                selected_malls = st.multiselect(
                    "쇼핑몰 선택",
                    options=df_record['쇼핑몰'].unique(),
                    default=df_record['쇼핑몰'].unique()
                )
            
            with col3:
                top_n = st.number_input("TOP N 상품", min_value=5, max_value=50, value=10)
            
            # 필터링 적용
            df_filtered = df_record[
                (pd.to_datetime(df_record['주문일자']).dt.date >= date_range[0]) &
                (pd.to_datetime(df_record['주문일자']).dt.date <= date_range[1]) &
                (df_record['쇼핑몰'].isin(selected_malls))
            ]
            
            if not df_filtered.empty:
                # 상세 통계
                st.subheader("📊 상세 통계")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.metric("필터링된 매출", f"₩{df_filtered['실결제금액'].sum():,.0f}")
                    st.metric("평균 주문 금액", f"₩{df_filtered['실결제금액'].mean():,.0f}")
                
                with col2:
                    st.metric("총 주문 건수", f"{len(df_filtered):,}")
                    st.metric("고유 상품 수", f"{df_filtered['SKU상품명'].nunique():,}")
                
                # 데이터 테이블
                st.subheader("📋 상세 데이터")
                st.dataframe(
                    df_filtered.sort_values('주문일자', ascending=False),
                    use_container_width=True,
                    height=400
                )
                
                # 데이터 다운로드
                csv = df_filtered.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "📥 CSV 다운로드",
                    csv,
                    "filtered_data.csv",
                    "text/csv",
                    key='download-csv'
                )
            else:
                st.warning("선택한 조건에 맞는 데이터가 없습니다.")
        else:
            st.info("📊 아직 기록된 데이터가 없습니다.")

if __name__ == "__main__":
    main()
