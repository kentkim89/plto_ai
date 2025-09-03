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
import plotly.express as px
import plotly.graph_objects as go
import google.generativeai as genai
from typing import Optional, Tuple, List, Dict
import hashlib
import json

# --------------------------------------------------------------------------
# SharePoint 연결 설정
# --------------------------------------------------------------------------

@st.cache_resource
def init_sharepoint_context():
    """SharePoint 컨텍스트 초기화"""
    try:
        tenant_id = st.secrets["sharepoint"]["tenant_id"]
        client_id = st.secrets["sharepoint"]["client_id"]
        client_secret = st.secrets["sharepoint"]["client_secret"]
        site_url = "https://goremi.sharepoint.com/sites/data"
        
        credentials = ClientCredential(client_id, client_secret)
        ctx = ClientContext(site_url).with_credentials(credentials)
        return ctx
    except Exception as e:
        st.error(f"SharePoint 연결 실패: {e}")
        return None

@st.cache_data(ttl=600)  # 10분 캐시
def load_master_data_from_sharepoint():
    """SharePoint에서 마스터 데이터 로드"""
    try:
        ctx = init_sharepoint_context()
        if not ctx:
            return load_local_master_data("master_data.csv")
        
        file_url = st.secrets["sharepoint_files"]["plto_master_data_file_url"]
        
        # SharePoint에서 파일 다운로드
        response = File.open_binary(ctx, file_url)
        
        # BytesIO로 변환 후 pandas로 읽기
        df_master = pd.read_excel(io.BytesIO(response.content))
        df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
        
        return df_master
    except Exception as e:
        st.warning(f"SharePoint 연결 실패, 로컬 파일 사용: {e}")
        return load_local_master_data("master_data.csv")

def load_local_master_data(file_path="master_data.csv"):
    """로컬 백업 마스터 데이터 로드"""
    try:
        df_master = pd.read_csv(file_path)
        df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
        return df_master
    except:
        st.error("로컬 마스터 데이터도 찾을 수 없습니다!")
        return pd.DataFrame()

def save_to_sharepoint_records(df_main_result, df_ecount_upload):
    """처리 결과를 SharePoint의 plto_record_data.xlsx에 저장"""
    try:
        ctx = init_sharepoint_context()
        if not ctx:
            return False, "SharePoint 연결 실패"
        
        record_file_url = st.secrets["sharepoint_files"]["plto_record_data_file_url"]
        
        # 기존 레코드 파일 다운로드 시도
        try:
            response = File.open_binary(ctx, record_file_url)
            existing_df = pd.read_excel(io.BytesIO(response.content))
        except:
            # 파일이 없으면 새로 생성
            existing_df = pd.DataFrame()
        
        # 새 데이터 준비
        new_records = pd.DataFrame()
        
        # 주문 날짜 추출 (이카운트 업로드 데이터의 일자 사용)
        order_date = df_ecount_upload['일자'].iloc[0] if not df_ecount_upload.empty else datetime.now().strftime("%Y%m%d")
        
        # 기록할 데이터 구성
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
        
        # 기존 데이터와 병합 (중복 제거)
        if not existing_df.empty and 'unique_hash' in existing_df.columns:
            # 중복되지 않는 새 레코드만 추가
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
# AI 분석 기능
# --------------------------------------------------------------------------

@st.cache_resource
def init_gemini():
    """Gemini AI 초기화"""
    try:
        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
        model = genai.GenerativeModel('gemini-pro')
        return model
    except Exception as e:
        st.error(f"Gemini AI 초기화 실패: {e}")
        return None

def analyze_sales_data_with_ai(df_records):
    """AI를 사용한 판매 데이터 분석"""
    try:
        model = init_gemini()
        if not model or df_records.empty:
            return None
        
        # 데이터 요약 준비
        summary = {
            "total_orders": len(df_records),
            "total_revenue": df_records['실결제금액'].sum(),
            "unique_products": df_records['SKU상품명'].nunique(),
            "unique_customers": df_records['수령자명'].nunique(),
            "date_range": f"{df_records['주문일자'].min()} ~ {df_records['주문일자'].max()}",
            "top_products": df_records.groupby('SKU상품명')['주문수량'].sum().nlargest(5).to_dict(),
            "channel_distribution": df_records.groupby('쇼핑몰')['실결제금액'].sum().to_dict()
        }
        
        prompt = f"""
        다음 온라인 쇼핑몰 판매 데이터를 분석하고 인사이트를 제공해주세요:
        
        {json.dumps(summary, ensure_ascii=False, indent=2)}
        
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
    try:
        ctx = init_sharepoint_context()
        if not ctx:
            return pd.DataFrame()
        
        record_file_url = st.secrets["sharepoint_files"]["plto_record_data_file_url"]
        response = File.open_binary(ctx, record_file_url)
        df_records = pd.read_excel(io.BytesIO(response.content))
        
        # 날짜 형식 정규화
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
        with st.spinner("AI가 데이터를 분석 중입니다..."):
            ai_insights = analyze_sales_data_with_ai(df_records)
            if ai_insights:
                st.markdown("### 🤖 AI 판매 분석 리포트")
                st.markdown(ai_insights)
            else:
                st.info("AI 분석을 사용할 수 없습니다.")

# --------------------------------------------------------------------------
# 기존 함수들 (수정 없음)
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

def process_all_files(file1, file2, file3, df_master):
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

        # 기존 처리 로직
        df_final = df_ecount_orig.copy().rename(columns={'금액': '실결제금액'})
        
        # 스마트스토어 병합 준비
        key_cols_smartstore = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = df_smartstore.rename(columns={'실결제금액': '수정될_금액_스토어'})[key_cols_smartstore + ['수정될_금액_스토어']].drop_duplicates(subset=key_cols_smartstore, keep='first')
        
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

        # 이카운트 업로드 데이터 생성
        df_merged = pd.merge(df_main_result, df_master[['SKU코드', '과세여부', '입수량']], 
                            left_on='재고관리코드', right_on='SKU코드', how='left')
        
        unmastered = df_merged[df_merged['SKU코드'].isna()]
        for _, row in unmastered.iterrows():
            warnings.append(f"- [미등록 상품] **{row['재고관리코드']}** / {row['SKU상품명']}")

        client_map = {
            '쿠팡': '쿠팡 주식회사', 
            '고도몰5': '고래미자사몰_현금영수증(고도몰)', 
            '스마트스토어': '스토어팜',
            '배민상회': '주식회사 우아한형제들(배민상회)',
            '이지웰몰': '주식회사 현대이지웰'
        }
        
        df_ecount_upload = pd.DataFrame()
        
        df_ecount_upload['일자'] = datetime.now().strftime("%Y%m%d")
        df_ecount_upload['거래처명'] = df_merged['쇼핑몰'].map(client_map).fillna(df_merged['쇼핑몰'])
        df_ecount_upload['출하창고'] = '고래미'
        df_ecount_upload['거래유형'] = np.where(df_merged['과세여부'] == '면세', 12, 11)
        df_ecount_upload['적요_전표'] = '오전/온라인'
        df_ecount_upload['품목코드'] = df_merged['재고관리코드']
        
        is_box_order = df_merged['SKU상품명'].str.contains("BOX", na=False)
        입수량 = pd.to_numeric(df_merged['입수량'], errors='coerce').fillna(1)
        base_quantity = np.where(is_box_order, df_merged['주문수량'] * 입수량, df_merged['주문수량'])
        is_3_pack = df_merged['SKU상품명'].str.contains("3개입|3개", na=False)
        final_quantity = np.where(is_3_pack, base_quantity * 3, base_quantity)
        df_ecount_upload['박스'] = np.where(is_box_order, df_merged['주문수량'], np.nan)
        df_ecount_upload['수량'] = final_quantity.astype(int)
        
        df_merged['실결제금액'] = pd.to_numeric(df_merged['실결제금액'], errors='coerce').fillna(0)
        공급가액 = np.where(df_merged['과세여부'] == '과세', df_merged['실결제금액'] / 1.1, df_merged['실결제금액'])
        df_ecount_upload['공급가액'] = 공급가액
        df_ecount_upload['부가세'] = df_merged['실결제금액'] - df_ecount_upload['공급가액']
        
        df_ecount_upload['쇼핑몰고객명'] = df_merged['수령자명']
        df_ecount_upload['original_order'] = df_merged['original_order']
        
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

        return df_main_result.drop(columns=['original_order']), df_quantity_summary, df_packing_list_final, df_ecount_upload, True, "모든 파일 처리가 성공적으로 완료되었습니다.", warnings

    except Exception as e:
        import traceback
        st.error(f"처리 중 오류가 발생했습니다: {e}")
        st.error(traceback.format_exc())
        return None, None, None, None, False, f"오류가 발생했습니다: {e}", []

# --------------------------------------------------------------------------
# Streamlit 앱 UI 구성
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="주문 처리 자동화 Pro v2.0", 
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# 사이드바 메뉴
with st.sidebar:
    st.title("📊 Order Pro v2.0")
    st.markdown("---")
    
    menu_option = st.radio(
        "메뉴 선택",
        ["📑 주문 처리", "📈 판매 분석", "⚙️ 설정"],
        index=0
    )
    
    st.markdown("---")
    st.caption("SharePoint 연동 상태")
    try:
        ctx = init_sharepoint_context()
        if ctx:
            st.success("✅ 연결됨")
        else:
            st.error("❌ 연결 실패")
    except:
        st.warning("⚠️ 확인 필요")

# 메인 콘텐츠
if menu_option == "📑 주문 처리":
    st.title("📑 주문 처리 자동화")
    st.info("💡 SharePoint와 연동하여 마스터 데이터를 자동으로 불러오고, 처리 결과를 자동 저장합니다.")
    
    st.write("---")
    st.header("1. 원본 엑셀 파일 3개 업로드")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        file1 = st.file_uploader("1️⃣ 스마트스토어 (금액확인용)", type=['xlsx', 'xls', 'csv'])
    with col2:
        file2 = st.file_uploader("2️⃣ 이카운트 다운로드 (주문목록)", type=['xlsx', 'xls', 'csv'])
    with col3:
        file3 = st.file_uploader("3️⃣ 고도몰 (금액확인용)", type=['xlsx', 'xls', 'csv'])
    
    st.write("---")
    st.header("2. 처리 결과 확인 및 다운로드")
    
    if st.button("🚀 모든 데이터 처리 및 파일 생성 실행", type="primary"):
        if file1 and file2 and file3:
            try:
                # SharePoint에서 마스터 데이터 로드
                with st.spinner('SharePoint에서 마스터 데이터를 불러오는 중...'):
                    df_master = load_master_data_from_sharepoint()
                
                if df_master.empty:
                    st.error("마스터 데이터를 불러올 수 없습니다.")
                else:
                    with st.spinner('모든 파일을 처리하는 중입니다...'):
                        df_main, df_qty, df_pack, df_ecount, success, message, warnings = process_all_files(
                            file1, file2, file3, df_master
                        )
                    
                    if success:
                        st.success(message)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        
                        # SharePoint에 기록 저장
                        with st.spinner('SharePoint에 처리 결과를 저장하는 중...'):
                            save_success, save_message = save_to_sharepoint_records(df_main, df_ecount)
                            if save_success:
                                st.success(f"✅ {save_message}")
                            else:
                                st.warning(f"⚠️ {save_message}")
                        
                        # 경고 메시지 표시
                        if warnings:
                            st.warning("⚠️ 확인 필요 항목")
                            with st.expander("자세한 목록 보기..."):
                                for warning_message in warnings:
                                    st.markdown(warning_message)
                        
                        # 결과 탭 표시
                        tab_erp, tab_pack, tab_qty, tab_main = st.tabs([
                            "🏢 **이카운트 업로드용**", 
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
                        st.error(message)
                        
            except Exception as e:
                st.error(f"🚨 처리 중 오류가 발생했습니다: {e}")
        else:
            st.warning("⚠️ 3개의 엑셀 파일을 모두 업로드해야 실행할 수 있습니다.")

elif menu_option == "📈 판매 분석":
    st.title("📈 AI 기반 판매 데이터 분석")
    st.info("💡 SharePoint에 저장된 판매 기록을 분석하고 AI 인사이트를 제공합니다.")
    
    # 분석 기간 선택
    col1, col2 = st.columns(2)
    with col1:
        analysis_period = st.selectbox(
            "분석 기간",
            ["최근 7일", "최근 30일", "최근 90일", "전체 기간", "사용자 지정"],
            index=1
        )
    
    with col2:
        if analysis_period == "사용자 지정":
            date_range = st.date_input(
                "날짜 범위",
                value=(datetime.now() - timedelta(days=30), datetime.now()),
                max_value=datetime.now()
            )
    
    if st.button("📊 분석 시작", type="primary"):
        with st.spinner("SharePoint에서 데이터를 불러오는 중..."):
            df_records = load_record_data_from_sharepoint()
            
            if not df_records.empty:
                # 기간 필터링
                if analysis_period != "전체 기간":
                    today = pd.Timestamp.now()
                    if analysis_period == "최근 7일":
                        start_date = today - timedelta(days=7)
                    elif analysis_period == "최근 30일":
                        start_date = today - timedelta(days=30)
                    elif analysis_period == "최근 90일":
                        start_date = today - timedelta(days=90)
                    elif analysis_period == "사용자 지정":
                        start_date = pd.Timestamp(date_range[0])
                        today = pd.Timestamp(date_range[1])
                    
                    df_records = df_records[
                        (df_records['주문일자'] >= start_date) & 
                        (df_records['주문일자'] <= today)
                    ]
                
                if not df_records.empty:
                    create_analytics_dashboard(df_records)
                else:
                    st.warning("선택한 기간에 데이터가 없습니다.")
            else:
                st.warning("분석할 데이터가 없습니다. 먼저 주문 처리를 실행해주세요.")

elif menu_option == "⚙️ 설정":
    st.title("⚙️ 시스템 설정")
    
    st.header("SharePoint 연결 정보")
    col1, col2 = st.columns(2)
    
    with col1:
        st.text_input("Tenant ID", value=st.secrets["sharepoint"]["tenant_id"], disabled=True)
        st.text_input("Client ID", value=st.secrets["sharepoint"]["client_id"], disabled=True)
    
    with col2:
        st.text_input("Site Name", value=st.secrets["sharepoint_files"]["site_name"], disabled=True)
        st.text_input("Master File", value=st.secrets["sharepoint_files"]["file_name"], disabled=True)
    
    st.header("AI 설정")
    st.text_input("Gemini API Key", value=st.secrets["GEMINI_API_KEY"][:10] + "...", disabled=True)
    
    if st.button("🔄 연결 테스트"):
        with st.spinner("테스트 중..."):
            # SharePoint 테스트
            ctx = init_sharepoint_context()
            if ctx:
                st.success("✅ SharePoint 연결 성공")
            else:
                st.error("❌ SharePoint 연결 실패")
            
            # AI 테스트
            model = init_gemini()
            if model:
                st.success("✅ Gemini AI 연결 성공")
            else:
                st.error("❌ Gemini AI 연결 실패")
    
    st.header("캐시 관리")
    if st.button("🗑️ 캐시 초기화"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.success("캐시가 초기화되었습니다.")
        st.rerun()
