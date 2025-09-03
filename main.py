import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from typing import Optional, Tuple, List, Dict
import hashlib
import json

# SharePoint 관련 imports는 try-except로 처리
try:
    from office365.runtime.auth.client_credential import ClientCredential
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File
    SHAREPOINT_AVAILABLE = True
except ImportError:
    SHAREPOINT_AVAILABLE = False
    st.warning("SharePoint 라이브러리가 설치되지 않았습니다. 로컬 모드로 실행됩니다.")

# Gemini AI import도 optional로 처리
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

# --------------------------------------------------------------------------
# SharePoint 연결 설정 (Optional)
# --------------------------------------------------------------------------

@st.cache_resource
def init_sharepoint_context():
    """SharePoint 컨텍스트 초기화"""
    if not SHAREPOINT_AVAILABLE:
        return None
    
    try:
        # secrets 체크
        if "sharepoint" not in st.secrets:
            return None
            
        tenant_id = st.secrets["sharepoint"]["tenant_id"]
        client_id = st.secrets["sharepoint"]["client_id"]
        client_secret = st.secrets["sharepoint"]["client_secret"]
        site_url = "https://goremi.sharepoint.com/sites/data"
        
        credentials = ClientCredential(client_id, client_secret)
        ctx = ClientContext(site_url).with_credentials(credentials)
        return ctx
    except Exception as e:
        st.warning(f"SharePoint 연결 실패: {e}")
        return None

@st.cache_data(ttl=600)
def load_master_data_from_sharepoint():
    """SharePoint에서 마스터 데이터 로드 또는 로컬 파일 사용"""
    if SHAREPOINT_AVAILABLE:
        try:
            ctx = init_sharepoint_context()
            if ctx and "sharepoint_files" in st.secrets:
                file_url = st.secrets["sharepoint_files"]["plto_master_data_file_url"]
                response = File.open_binary(ctx, file_url)
                df_master = pd.read_excel(io.BytesIO(response.content))
                df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
                return df_master
        except Exception as e:
            st.info(f"SharePoint 접속 실패, 로컬 파일 사용: {e}")
    
    # 로컬 파일 로드
    return load_local_master_data("master_data.csv")

def load_local_master_data(file_path="master_data.csv"):
    """로컬 백업 마스터 데이터 로드"""
    try:
        df_master = pd.read_csv(file_path)
        df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
        return df_master
    except Exception as e:
        st.error(f"마스터 데이터 파일을 찾을 수 없습니다: {e}")
        return pd.DataFrame()

def save_to_sharepoint_records(df_main_result, df_ecount_upload):
    """처리 결과를 SharePoint에 저장 (Optional)"""
    if not SHAREPOINT_AVAILABLE:
        return False, "SharePoint 기능이 비활성화되어 있습니다."
    
    try:
        ctx = init_sharepoint_context()
        if not ctx:
            return False, "SharePoint 연결 실패"
        
        if "sharepoint_files" not in st.secrets or "plto_record_data_file_url" not in st.secrets["sharepoint_files"]:
            return False, "레코드 파일 URL이 설정되지 않았습니다."
        
        record_file_url = st.secrets["sharepoint_files"]["plto_record_data_file_url"]
        
        # 기존 레코드 파일 다운로드 시도
        try:
            response = File.open_binary(ctx, record_file_url)
            existing_df = pd.read_excel(io.BytesIO(response.content))
        except:
            existing_df = pd.DataFrame()
        
        # 새 데이터 준비
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
        
        # 중복 체크용 해시 생성
        new_records['unique_hash'] = new_records.apply(
            lambda x: hashlib.md5(
                f"{x['주문일자']}_{x['재고관리코드']}_{x['수령자명']}_{x['쇼핑몰']}".encode()
            ).hexdigest(), axis=1
        )
        
        # 기존 데이터와 병합
        if not existing_df.empty and 'unique_hash' in existing_df.columns:
            new_unique_records = new_records[~new_records['unique_hash'].isin(existing_df['unique_hash'])]
            combined_df = pd.concat([existing_df, new_unique_records], ignore_index=True)
        else:
            combined_df = new_records
        
        # Excel 파일로 저장
        output = io.BytesIO()
        combined_df.to_excel(output, index=False, sheet_name='Records')
        output.seek(0)
        
        # SharePoint에 업로드
        target_folder = ctx.web.get_folder_by_server_relative_url("/sites/data/Shared Documents")
        target_folder.upload_file("plto_record_data.xlsx", output.read()).execute_query()
        
        return True, f"성공적으로 {len(new_records)}개의 레코드를 저장했습니다."
        
    except Exception as e:
        return False, f"SharePoint 저장 실패: {e}"

# --------------------------------------------------------------------------
# AI 분석 기능 (Optional)
# --------------------------------------------------------------------------

@st.cache_resource
def init_gemini():
    """Gemini AI 초기화"""
    if not GEMINI_AVAILABLE:
        return None
    
    try:
        if "GEMINI_API_KEY" in st.secrets:
            genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
            model = genai.GenerativeModel('gemini-pro')
            return model
    except Exception as e:
        st.warning(f"Gemini AI 초기화 실패: {e}")
    return None

def analyze_sales_data_with_ai(df_records):
    """AI를 사용한 판매 데이터 분석"""
    if not GEMINI_AVAILABLE:
        return "AI 분석 기능이 비활성화되어 있습니다."
    
    try:
        model = init_gemini()
        if not model or df_records.empty:
            return None
        
        # 데이터 요약 준비
        summary = {
            "total_orders": len(df_records),
            "total_revenue": float(df_records['실결제금액'].sum()),
            "unique_products": int(df_records['SKU상품명'].nunique()),
            "unique_customers": int(df_records['수령자명'].nunique()),
            "date_range": f"{df_records['주문일자'].min()} ~ {df_records['주문일자'].max()}",
            "top_products": df_records.groupby('SKU상품명')['주문수량'].sum().nlargest(5).to_dict(),
            "channel_distribution": {k: float(v) for k, v in df_records.groupby('쇼핑몰')['실결제금액'].sum().to_dict().items()}
        }
        
        prompt = f"""
        다음 온라인 쇼핑몰 판매 데이터를 분석하고 인사이트를 제공해주세요:
        
        {json.dumps(summary, ensure_ascii=False, indent=2, default=str)}
        
        다음 항목들을 포함해서 분석해주세요:
        1. 전체적인 판매 트렌드
        2. 베스트셀러 상품 분석
        3. 채널별 판매 성과
        4. 개선 제안사항
        
        간결하고 실용적인 인사이트를 제공해주세요.
        """
        
        response = model.generate_content(prompt)
        return response.text
        
    except Exception as e:
        return f"AI 분석 중 오류 발생: {e}"

def load_record_data_from_sharepoint():
    """SharePoint에서 기록 데이터 로드"""
    if not SHAREPOINT_AVAILABLE:
        return pd.DataFrame()
    
    try:
        ctx = init_sharepoint_context()
        if not ctx or "sharepoint_files" not in st.secrets:
            return pd.DataFrame()
        
        if "plto_record_data_file_url" not in st.secrets["sharepoint_files"]:
            return pd.DataFrame()
            
        record_file_url = st.secrets["sharepoint_files"]["plto_record_data_file_url"]
        response = File.open_binary(ctx, record_file_url)
        df_records = pd.read_excel(io.BytesIO(response.content))
        
        if '주문일자' in df_records.columns:
            df_records['주문일자'] = pd.to_datetime(df_records['주문일자'], format='%Y%m%d', errors='coerce')
        
        return df_records
    except:
        return pd.DataFrame()

def create_analytics_dashboard(df_records):
    """분석 대시보드 생성"""
    if df_records.empty:
        st.warning("분석할 데이터가 없습니다.")
        return
    
    # 날짜별 집계
    df_daily = df_records.groupby('주문일자').agg({
        '실결제금액': 'sum',
        '주문수량': 'sum',
        '수령자명': 'nunique'
    }).reset_index()
    df_daily.columns = ['날짜', '매출액', '판매수량', '고객수']
    
    # 상품별 판매 TOP 10
    df_product_top = df_records.groupby('SKU상품명').agg({
        '주문수량': 'sum',
        '실결제금액': 'sum'
    }).nlargest(10, '주문수량').reset_index()
    
    # 채널별 매출
    df_channel = df_records.groupby('쇼핑몰')['실결제금액'].sum().reset_index()
    
    # 대시보드 레이아웃
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_revenue = df_records['실결제금액'].sum()
        st.metric("총 매출", f"₩{total_revenue:,.0f}")
    
    with col2:
        total_orders = len(df_records)
        st.metric("총 주문수", f"{total_orders:,}")
    
    with col3:
        avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
        st.metric("평균 주문 금액", f"₩{avg_order_value:,.0f}")
    
    with col4:
        unique_customers = df_records['수령자명'].nunique()
        st.metric("고객수", f"{unique_customers:,}")
    
    # 차트 생성
    tab1, tab2, tab3, tab4 = st.tabs(["📈 일별 트렌드", "🏆 베스트셀러", "🛒 채널 분석", "🤖 AI 인사이트"])
    
    with tab1:
        if not df_daily.empty:
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Scatter(
                x=df_daily['날짜'], 
                y=df_daily['매출액'],
                mode='lines+markers',
                name='매출액',
                line=dict(color='#1f77b4', width=2)
            ))
            fig_trend.update_layout(
                title="일별 매출 트렌드",
                xaxis_title="날짜",
                yaxis_title="매출액 (원)",
                hovermode='x unified'
            )
            st.plotly_chart(fig_trend, use_container_width=True)
    
    with tab2:
        if not df_product_top.empty:
            fig_products = px.bar(
                df_product_top, 
                x='주문수량', 
                y='SKU상품명',
                orientation='h',
                title="상품별 판매 수량 TOP 10",
                color='실결제금액',
                color_continuous_scale='Blues'
            )
            st.plotly_chart(fig_products, use_container_width=True)
    
    with tab3:
        if not df_channel.empty:
            fig_channel = px.pie(
                df_channel, 
                values='실결제금액', 
                names='쇼핑몰',
                title="채널별 매출 비중"
            )
            st.plotly_chart(fig_channel, use_container_width=True)
    
    with tab4:
        if GEMINI_AVAILABLE:
            with st.spinner("AI가 데이터를 분석 중입니다..."):
                ai_insights = analyze_sales_data_with_ai(df_records)
                if ai_insights:
                    st.markdown("### 🤖 AI 판매 분석 리포트")
                    st.markdown(ai_insights)
                else:
                    st.info("AI 분석을 사용할 수 없습니다.")
        else:
            st.info("AI 분석 기능을 사용하려면 google-generativeai를 설치하세요.")

# --------------------------------------------------------------------------
# 기존 함수들
# --------------------------------------------------------------------------

def to_excel_formatted(df, format_type=None):
    """데이터프레임을 서식이 적용된 엑셀 파일 형식의 BytesIO 객체로 변환하는 함수"""
    output = io.BytesIO()
    df_to_save = df.fillna('')
    
    if format_type == 'ecount_upload':
        df_to_save = df_to_save.rename(columns={'적요_전표': '적요', '적요_품목': '적요.1'})

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_save.to_excel(writer, index=False, sheet_name='Sheet1')
    
    output.seek(0)
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
        adjusted_width = min((max_length + 2) * 1.2, 50)  # 최대 너비 제한
        sheet.column_dimensions[column].width = adjusted_width
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
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
                        sheet.merge_cells(start_row=bundle_start_row, start_column=1, 
                                        end_row=bundle_end_row, end_column=1)
                        sheet.merge_cells(start_row=bundle_start_row, start_column=4, 
                                        end_row=bundle_end_row, end_column=4)
                
                bundle_start_row = row_num

    if format_type == 'quantity_summary':
        for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=sheet.max_row, 
                                                     min_col=1, max_col=sheet.max_column)):
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

        df_ecount_orig['original_order'] = range(len(df_ecount_orig))
        
        # 컬럼명 호환성 처리
        if '회 할인 금액' in df_godomall.columns and '회원 할인 금액' not in df_godomall.columns:
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
        
        # 결제 금액 검증
        godomall_warnings = []
        grouped_godomall = df_godomall.groupby('수취인 이름')
        
        for name, group in grouped_godomall:
            calculated_total = group['수정될_금액_고도몰'].sum()
            actual_total = group['총 결제 금액'].iloc[0]
            discrepancy = calculated_total - actual_total
            
            if abs(discrepancy) > 1:
                warning_msg = f"- [고도몰 금액 불일치] **{name}**님의 주문 금액 차이: **{discrepancy:,.0f}원**"
                godomall_warnings.append(warning_msg)

        # 기존 처리 로직
        df_final = df_ecount_orig.copy().rename(columns={'금액': '실결제금액'})
        
        # 스마트스토어 병합
        key_cols_smartstore = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = df_smartstore.rename(columns={'실결제금액': '수정될_금액_스토어'})[
            key_cols_smartstore + ['수정될_금액_스토어']
        ].drop_duplicates(subset=key_cols_smartstore, keep='first')
        
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
                            on=['재고관리코드', '수령자명', '주문수량'], how='left')

        # 경고 메시지 생성
        warnings = []
        failed_corrections = df_final[
            ((df_final['쇼핑몰'] == '스마트스토어') & (df_final['수정될_금액_스토어'].isna())) |
            ((df_final['쇼핑몰'] == '고도몰5') & (df_final['수정될_금액_고도몰'].isna()))
        ]
        
        for _, row in failed_corrections.iterrows():
            warnings.append(f"- [금액보정 실패] **{row['쇼핑몰']}** / {row['수령자명']} / {row['SKU상품명']}")
        
        warnings.extend(godomall_warnings)

        # 최종 결제 금액 업데이트
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
        
        # 동명이인 경고
        name_groups = df_main_result.groupby('수령자명')['original_order'].apply(list)
        for name, orders in name_groups.items():
            if len(orders) > 1 and (max(orders) - min(orders) + 1) != len(orders):
                warnings.append(f"- [동명이인 의심] **{name}** 님의 주문이 떨어져서 입력되었습니다.")

        # 요약 데이터 생성
        df_quantity_summary = df_main_result.groupby('SKU상품명', as_index=False)['주문수량'].sum().rename(
            columns={'주문수량': '개수'}
        )
        
        df_packing_list = df_main_result.sort_values(by='original_order')[
            ['SKU상품명', '주문수량', '수령자명', '쇼핑몰']
        ].copy()
        
        is_first_item = df_packing_list['수령자명'] != df_packing_list['수령자명'].shift(1)
        df_packing_list['묶음번호'] = is_first_item.cumsum()
        df_packing_list_final = df_packing_list.copy()
        df_packing_list_final['묶음번호'] = df_packing_list_final['묶음번호'].where(is_first_item, '')
        df_packing_list_final = df_packing_list_final[['묶음번호', 'SKU상품명', '주문수량', '수령자명', '쇼핑몰']]

        # 이카운트 업로드 데이터 생성
        df_merged = pd.merge(
            df_main_result, 
            df_master[['SKU코드', '과세여부', '입수량']], 
            left_on='재고관리코드', 
            right_on='SKU코드', 
            how='left'
        )
        
        # 미등록 상품 경고
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
        
        # 이카운트 데이터 생성
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

        return (df_main_result.drop(columns=['original_order']), 
                df_quantity_summary, 
                df_packing_list_final, 
                df_ecount_upload, 
                True, 
                "모든 파일 처리가 성공적으로 완료되었습니다.", 
                warnings)

    except Exception as e:
        import traceback
        error_msg = f"처리 중 오류가 발생했습니다: {str(e)}\n{traceback.format_exc()}"
        return None, None, None, None, False, error_msg, []

# --------------------------------------------------------------------------
# Streamlit 앱 UI 구성
# --------------------------------------------------------------------------

# 페이지 설정
st.set_page_config(
    page_title="주문 처리 자동화 v2.0",
    layout="wide",
    page_icon="📊"
)

# 사이드바
with st.sidebar:
    st.title("📊 Order Pro v2.0")
    st.markdown("---")
    
    menu_option = st.radio(
        "메뉴 선택",
        ["📑 주문 처리", "📈 판매 분석", "⚙️ 설정"],
        index=0
    )
    
    st.markdown("---")
    st.caption("연결 상태")
    
    # SharePoint 상태
    if SHAREPOINT_AVAILABLE:
        ctx = init_sharepoint_context()
        if ctx:
            st.success("✅ SharePoint 연결")
        else:
            st.warning("⚠️ SharePoint 오프라인")
    else:
        st.info("💾 로컬 모드")
    
    # AI 상태
    if GEMINI_AVAILABLE:
        if "GEMINI_API_KEY" in st.secrets:
            st.success("✅ AI 활성화")
        else:
            st.warning("⚠️ AI 키 필요")
    else:
        st.info("🤖 AI 비활성화")

# 메인 화면
if menu_option == "📑 주문 처리":
    st.title("📑 주문 처리 자동화")
    
    if SHAREPOINT_AVAILABLE and init_sharepoint_context():
        st.info("💡 SharePoint와 연동하여 자동으로 데이터를 처리합니다.")
    else:
        st.info("💡 로컬 모드로 실행 중입니다. master_data.csv 파일을 사용합니다.")
    
    st.write("---")
    st.header("1. 원본 엑셀 파일 3개 업로드")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        file1 = st.file_uploader("1️⃣ 스마트스토어", type=['xlsx', 'xls'])
    with col2:
        file2 = st.file_uploader("2️⃣ 이카운트", type=['xlsx', 'xls'])
    with col3:
        file3 = st.file_uploader("3️⃣ 고도몰", type=['xlsx', 'xls'])
    
    st.write("---")
    st.header("2. 처리 실행")
    
    if st.button("🚀 데이터 처리 시작", type="primary", disabled=not (file1 and file2 and file3)):
        if file1 and file2 and file3:
            try:
                # 마스터 데이터 로드
                with st.spinner('마스터 데이터를 불러오는 중...'):
                    df_master = load_master_data_from_sharepoint()
                
                if df_master.empty:
                    st.error("마스터 데이터를 불러올 수 없습니다.")
                else:
                    # 파일 처리
                    with st.spinner('파일을 처리하는 중...'):
                        result = process_all_files(file1, file2, file3, df_master)
                        df_main, df_qty, df_pack, df_ecount, success, message, warnings = result
                    
                    if success:
                        st.success(message)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        
                        # SharePoint 저장 (옵션)
                        if SHAREPOINT_AVAILABLE and init_sharepoint_context():
                            with st.spinner('SharePoint에 저장 중...'):
                                save_success, save_msg = save_to_sharepoint_records(df_main, df_ecount)
                                if save_success:
                                    st.success(f"✅ {save_msg}")
                                else:
                                    st.info(f"ℹ️ {save_msg}")
                        
                        # 경고 표시
                        if warnings:
                            with st.expander("⚠️ 확인 필요 항목"):
                                for w in warnings:
                                    st.markdown(w)
                        
                        # 결과 표시
                        tabs = st.tabs(["🏢 이카운트", "📋 포장리스트", "📦 수량요약", "✅ 최종결과"])
                        
                        with tabs[0]:
                            st.dataframe(df_ecount.astype(str), use_container_width=True)
                            st.download_button(
                                "📥 다운로드",
                                to_excel_formatted(df_ecount, 'ecount_upload'),
                                f"이카운트_{timestamp}.xlsx"
                            )
                        
                        with tabs[1]:
                            st.dataframe(df_pack, use_container_width=True)
                            st.download_button(
                                "📥 다운로드",
                                to_excel_formatted(df_pack, 'packing_list'),
                                f"포장리스트_{timestamp}.xlsx"
                            )
                        
                        with tabs[2]:
                            st.dataframe(df_qty, use_container_width=True)
                            st.download_button(
                                "📥 다운로드",
                                to_excel_formatted(df_qty, 'quantity_summary'),
                                f"수량요약_{timestamp}.xlsx"
                            )
                        
                        with tabs[3]:
                            st.dataframe(df_main, use_container_width=True)
                            st.download_button(
                                "📥 다운로드",
                                to_excel_formatted(df_main),
                                f"최종결과_{timestamp}.xlsx"
                            )
                    else:
                        st.error(message)
                        
            except Exception as e:
                st.error(f"처리 중 오류: {e}")
        else:
            st.warning("3개 파일을 모두 업로드하세요.")

elif menu_option == "📈 판매 분석":
    st.title("📈 판매 데이터 분석")
    
    if not SHAREPOINT_AVAILABLE:
        st.warning("SharePoint가 연결되지 않아 분석 기능을 사용할 수 없습니다.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            period = st.selectbox("분석 기간", ["최근 7일", "최근 30일", "전체"])
        
        if st.button("📊 분석 시작", type="primary"):
            with st.spinner("데이터 로드 중..."):
                df_records = load_record_data_from_sharepoint()
                
                if not df_records.empty:
                    # 기간 필터
                    if period != "전체":
                        days = 7 if period == "최근 7일" else 30
                        cutoff = datetime.now() - timedelta(days=days)
                        df_records = df_records[df_records['주문일자'] >= cutoff]
                    
                    if not df_records.empty:
                        create_analytics_dashboard(df_records)
                    else:
                        st.info("선택 기간에 데이터가 없습니다.")
                else:
                    st.info("분석할 데이터가 없습니다.")

elif menu_option == "⚙️ 설정":
    st.title("⚙️ 시스템 설정")
    
    st.header("연결 정보")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("SharePoint")
        if SHAREPOINT_AVAILABLE:
            if "sharepoint" in st.secrets:
                st.text_input("Tenant ID", value=st.secrets["sharepoint"]["tenant_id"][:10] + "...", disabled=True)
                st.text_input("Client ID", value=st.secrets["sharepoint"]["client_id"][:10] + "...", disabled=True)
            else:
                st.info("SharePoint 설정이 없습니다.")
        else:
            st.info("SharePoint 라이브러리가 설치되지 않았습니다.")
    
    with col2:
        st.subheader("AI")
        if GEMINI_AVAILABLE:
            if "GEMINI_API_KEY" in st.secrets:
                st.text_input("API Key", value=st.secrets["GEMINI_API_KEY"][:10] + "...", disabled=True)
            else:
                st.info("AI API 키가 설정되지 않았습니다.")
        else:
            st.info("AI 라이브러리가 설치되지 않았습니다.")
    
    if st.button("🔄 연결 테스트"):
        with st.spinner("테스트 중..."):
            # SharePoint 테스트
            if SHAREPOINT_AVAILABLE:
                ctx = init_sharepoint_context()
                if ctx:
                    st.success("✅ SharePoint 연결 성공")
                else:
                    st.error("❌ SharePoint 연결 실패")
            
            # AI 테스트
            if GEMINI_AVAILABLE:
                model = init_gemini()
                if model:
                    st.success("✅ AI 연결 성공")
                else:
                    st.error("❌ AI 연결 실패")
    
    st.header("캐시 관리")
    if st.button("🗑️ 캐시 초기화"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.success("캐시를 초기화했습니다.")
        st.rerun()
