import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment
from datetime import datetime
import requests
from urllib.parse import urlparse

# --------------------------------------------------------------------------
# 페이지 설정
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="주문 처리 자동화 v2.0",
    layout="wide",
    page_icon="📊"
)

# --------------------------------------------------------------------------
# SharePoint 데이터 로드 함수
# --------------------------------------------------------------------------

@st.cache_data(ttl=600)  # 10분 캐시
def load_master_data_from_sharepoint():
    """SharePoint URL에서 직접 마스터 데이터 로드"""
    try:
        # Streamlit secrets에서 SharePoint URL 가져오기
        if "sharepoint_files" not in st.secrets:
            st.error("SharePoint 설정이 없습니다. secrets.toml 파일을 확인하세요.")
            return pd.DataFrame()
        
        file_url = st.secrets["sharepoint_files"]["plto_master_data_file_url"]
        
        # SharePoint 공유 링크를 다운로드 가능한 링크로 변환
        # SharePoint 링크 형식: https://goremi.sharepoint.com/:x:/s/data/...
        # 다운로드 링크 형식: https://goremi.sharepoint.com/sites/data/_layouts/15/download.aspx?share=...
        
        if "sharepoint.com/:x:" in file_url:
            # 공유 링크에서 share ID 추출
            parts = file_url.split('/')
            share_id = parts[-1].split('?')[0]
            base_url = file_url.split('/:x:')[0]
            site_path = "/sites/data"
            download_url = f"{base_url}{site_path}/_layouts/15/download.aspx?share={share_id}"
        else:
            # 이미 다운로드 가능한 링크인 경우
            download_url = file_url
        
        # 파일 다운로드
        response = requests.get(download_url, timeout=30)
        response.raise_for_status()
        
        # Excel 파일 읽기
        df_master = pd.read_excel(io.BytesIO(response.content))
        df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
        
        st.success("✅ SharePoint에서 마스터 데이터를 성공적으로 불러왔습니다.")
        return df_master
        
    except Exception as e:
        st.error(f"❌ SharePoint 데이터 로드 실패: {e}")
        
        # 대체 방법: 사용자가 직접 업로드
        st.warning("SharePoint 연결 실패. 마스터 데이터를 직접 업로드해주세요.")
        return pd.DataFrame()

# --------------------------------------------------------------------------
# 기존 처리 함수들
# --------------------------------------------------------------------------

def to_excel_formatted(df, format_type=None):
    """데이터프레임을 서식이 적용된 엑셀 파일로 변환"""
    output = io.BytesIO()
    df_to_save = df.fillna('')
    
    if format_type == 'ecount_upload':
        df_to_save = df_to_save.rename(columns={'적요_전표': '적요', '적요_품목': '적요.1'})

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_save.to_excel(writer, index=False, sheet_name='Sheet1')
    
    output.seek(0)
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    # 가운데 정렬
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
    
    # 테두리와 색상
    thin_border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    pink_fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")

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
                        sheet.merge_cells(
                            start_row=bundle_start_row, start_column=1,
                            end_row=bundle_end_row, end_column=1
                        )
                        sheet.merge_cells(
                            start_row=bundle_start_row, start_column=4,
                            end_row=bundle_end_row, end_column=4
                        )
                
                bundle_start_row = row_num

    elif format_type == 'quantity_summary':
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
        
        # 경고 메시지 수집
        warnings = []
        
        # 고도몰 금액 검증
        for name, group in df_godomall.groupby('수취인 이름'):
            calculated = group['수정될_금액_고도몰'].sum()
            actual = group['총 결제 금액'].iloc[0]
            diff = calculated - actual
            if abs(diff) > 1:
                warnings.append(f"- [금액 불일치] {name}님: {diff:,.0f}원 차이")

        # 메인 데이터 처리
        df_final = df_ecount_orig.copy().rename(columns={'금액': '실결제금액'})
        
        # 스마트스토어 병합
        key_cols = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = df_smartstore.rename(
            columns={'실결제금액': '수정될_금액_스토어'}
        )[key_cols + ['수정될_금액_스토어']].drop_duplicates(subset=key_cols, keep='first')
        
        # 고도몰 병합
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
        
        df_final['주문수량'] = pd.to_numeric(df_final['주문수량'], errors='coerce').fillna(0).astype(int)
        smartstore_prices['주문수량'] = pd.to_numeric(smartstore_prices['주문수량'], errors='coerce').fillna(0).astype(int)
        godomall_prices['주문수량'] = pd.to_numeric(godomall_prices['주문수량'], errors='coerce').fillna(0).astype(int)
        
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
        
        # 결과 데이터프레임
        df_main_result = df_final[[
            '재고관리코드', 'SKU상품명', '주문수량', '실결제금액', 
            '쇼핑몰', '수령자명', 'original_order'
        ]]
        
        # 수량 요약
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
        
        # 미등록 상품 경고
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
        
        # 이카운트 업로드 데이터
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
        import traceback
        return None, None, None, None, False, f"❌ 오류: {str(e)}\n{traceback.format_exc()}", []

# --------------------------------------------------------------------------
# 메인 앱
# --------------------------------------------------------------------------

st.title("📊 주문 처리 자동화 시스템 v2.0")
st.markdown("### SharePoint 연동 버전")
st.markdown("---")

# 사이드바
with st.sidebar:
    st.header("📌 상태")
    
    # SharePoint 연결 상태 체크
    if "sharepoint_files" in st.secrets:
        st.success("✅ SharePoint 설정 완료")
    else:
        st.error("❌ SharePoint 미설정")
        st.info("secrets.toml 파일 설정 필요")
    
    # 캐시 초기화 버튼
    if st.button("🔄 캐시 초기화"):
        st.cache_data.clear()
        st.success("캐시가 초기화되었습니다.")
        st.rerun()
    
    st.markdown("---")
    st.caption("© 2024 PLTO System")

# 탭 생성
tab1, tab2, tab3 = st.tabs(["📑 주문 처리", "📈 통계", "⚙️ 설정"])

with tab1:
    st.header("📤 파일 업로드")
    
    # 마스터 데이터 로드 섹션
    with st.expander("📊 마스터 데이터 상태", expanded=True):
        if st.button("🔄 SharePoint에서 마스터 데이터 새로고침"):
            st.cache_data.clear()
            
        df_master = load_master_data_from_sharepoint()
        
        if not df_master.empty:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("총 SKU", f"{len(df_master):,}개")
            with col2:
                st.metric("과세 상품", f"{len(df_master[df_master['과세여부']=='과세']):,}개")
            with col3:
                st.metric("면세 상품", f"{len(df_master[df_master['과세여부']=='면세']):,}개")
        else:
            st.warning("⚠️ 마스터 데이터가 없습니다. 수동으로 업로드하세요.")
            manual_master = st.file_uploader("마스터 데이터 업로드", type=['xlsx', 'xls', 'csv'], key="master")
            if manual_master:
                try:
                    if manual_master.name.endswith('.csv'):
                        df_master = pd.read_csv(manual_master)
                    else:
                        df_master = pd.read_excel(manual_master)
                    df_master = df_master.drop_duplicates(subset=['SKU코드'], keep='first')
                    st.success(f"✅ 마스터 데이터 로드 완료: {len(df_master)}개 SKU")
                except Exception as e:
                    st.error(f"파일 읽기 실패: {e}")
    
    st.markdown("---")
    
    # 주문 파일 업로드
    col1, col2, col3 = st.columns(3)
    with col1:
        file1 = st.file_uploader("📱 스마트스토어", type=['xlsx', 'xls'], key="smart")
    with col2:
        file2 = st.file_uploader("📋 이카운트", type=['xlsx', 'xls'], key="ecount")
    with col3:
        file3 = st.file_uploader("🏪 고도몰", type=['xlsx', 'xls'], key="godo")
    
    # 처리 버튼
    if st.button("🚀 처리 시작", type="primary", use_container_width=True):
        if file1 and file2 and file3:
            if not df_master.empty:
                # 파일 처리
                with st.spinner('데이터 처리 중... 잠시만 기다려주세요.'):
                    result = process_all_files(file1, file2, file3, df_master)
                    df_main, df_qty, df_pack, df_ecount, success, message, warnings = result
                
                if success:
                    st.balloons()
                    st.success(message)
                    
                    # 경고 표시
                    if warnings:
                        with st.expander(f"⚠️ 확인 필요 항목 ({len(warnings)}건)", expanded=True):
                            for w in warnings:
                                st.write(w)
                    
                    # 결과 저장
                    st.session_state['result'] = {
                        'main': df_main,
                        'qty': df_qty,
                        'pack': df_pack,
                        'ecount': df_ecount
                    }
                    
                    # 다운로드 섹션
                    st.markdown("---")
                    st.subheader("📥 다운로드")
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.download_button(
                            "💼 이카운트",
                            to_excel_formatted(df_ecount, 'ecount_upload'),
                            f"이카운트_{timestamp}.xlsx",
                            mime="application/vnd.ms-excel",
                            use_container_width=True
                        )
                    
                    with col2:
                        st.download_button(
                            "📦 포장리스트",
                            to_excel_formatted(df_pack, 'packing_list'),
                            f"포장_{timestamp}.xlsx",
                            mime="application/vnd.ms-excel",
                            use_container_width=True
                        )
                    
                    with col3:
                        st.download_button(
                            "📊 수량요약",
                            to_excel_formatted(df_qty, 'quantity_summary'),
                            f"수량_{timestamp}.xlsx",
                            mime="application/vnd.ms-excel",
                            use_container_width=True
                        )
                    
                    with col4:
                        st.download_button(
                            "✅ 최종결과",
                            to_excel_formatted(df_main),
                            f"결과_{timestamp}.xlsx",
                            mime="application/vnd.ms-excel",
                            use_container_width=True
                        )
                    
                    # 미리보기
                    st.markdown("---")
                    st.subheader("📋 데이터 미리보기")
                    
                    preview_option = st.selectbox(
                        "확인할 데이터 선택",
                        ["이카운트 업로드", "포장 리스트", "수량 요약", "최종 결과"]
                    )
                    
                    if preview_option == "이카운트 업로드":
                        st.dataframe(df_ecount, use_container_width=True)
                    elif preview_option == "포장 리스트":
                        st.dataframe(df_pack, use_container_width=True)
                    elif preview_option == "수량 요약":
                        st.dataframe(df_qty, use_container_width=True)
                    else:
                        st.dataframe(df_main, use_container_width=True)
                    
                else:
                    st.error(message)
            else:
                st.error("❌ 마스터 데이터가 필요합니다!")
        else:
            st.warning("⚠️ 3개 파일을 모두 업로드해주세요!")

with tab2:
    st.header("📊 통계 분석")
    
    if 'result' in st.session_state:
        df = st.session_state['result']['main']
        
        # 기본 메트릭
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_revenue = df['실결제금액'].sum()
            st.metric("💰 총 매출", f"₩{total_revenue:,.0f}")
        
        with col2:
            total_orders = len(df)
            st.metric("📦 총 주문수", f"{total_orders:,}")
        
        with col3:
            avg_order = total_revenue / total_orders if total_orders > 0 else 0
            st.metric("💵 평균 주문액", f"₩{avg_order:,.0f}")
        
        with col4:
            unique_customers = df['수령자명'].nunique()
            st.metric("👥 고객수", f"{unique_customers:,}")
        
        st.markdown("---")
        
        # 채널별 분석
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🛒 채널별 매출")
            channel_stats = df.groupby('쇼핑몰')['실결제금액'].sum().sort_values(ascending=True)
            st.bar_chart(channel_stats)
            
            # 상세 테이블
            channel_detail = df.groupby('쇼핑몰').agg({
                '실결제금액': 'sum',
                '주문수량': 'sum',
                '수령자명': 'nunique'
            }).round(0)
            channel_detail.columns = ['매출액', '판매수량', '고객수']
            st.dataframe(channel_detail, use_container_width=True)
        
        with col2:
            st.subheader("🏆 TOP 10 상품")
            top_products = df.groupby('SKU상품명')['주문수량'].sum().nlargest(10).sort_values(ascending=True)
            st.bar_chart(top_products)
            
            # 상세 테이블
            product_detail = df.groupby('SKU상품명').agg({
                '주문수량': 'sum',
                '실결제금액': 'sum'
            }).nlargest(10, '주문수량')
            product_detail.columns = ['판매수량', '매출액']
            st.dataframe(product_detail, use_container_width=True)
        
    else:
        st.info("📌 먼저 주문 처리를 실행해주세요.")

with tab3:
    st.header("⚙️ 시스템 설정")
    
    st.subheader("📁 SharePoint 설정")
    
    if "sharepoint_files" in st.secrets:
        st.code(f"""
[sharepoint_files]
plto_master_data_file_url = "{st.secrets['sharepoint_files']['plto_master_data_file_url'][:50]}..."
site_name = "{st.secrets['sharepoint_files']['site_name']}"
file_name = "{st.secrets['sharepoint_files']['file_name']}"
        """)
        st.success("✅ SharePoint 설정이 정상적으로 구성되어 있습니다.")
    else:
        st.warning("⚠️ SharePoint 설정이 없습니다.")
        st.info("""
        Streamlit secrets.toml 파일에 다음과 같이 설정하세요:
        
        ```toml
        [sharepoint_files]
        plto_master_data_file_url = "SharePoint 공유 링크"
        site_name = "data"
        file_name = "plto_master_data.xlsx"
        ```
        """)
    
    st.markdown("---")
    
    st.subheader("📝 사용 방법")
    st.markdown("""
    1. **마스터 데이터 확인**: SharePoint에서 자동으로 불러옵니다
    2. **파일 업로드**: 스마트스토어, 이카운트, 고도몰 파일을 각각 업로드
    3. **처리 실행**: '처리 시작' 버튼 클릭
    4. **결과 다운로드**: 생성된 4개의 파일 다운로드
    5. **통계 확인**: 통계 탭에서 판매 분석 확인
    """)
    
    st.markdown("---")
    
    st.subheader("🔧 문제 해결")
    st.markdown("""
    - **SharePoint 연결 실패**: secrets.toml 파일 확인
    - **마스터 데이터 없음**: 수동으로 업로드 가능
    - **처리 오류**: 파일 형식과 컬럼명 확인
    - **캐시 문제**: 사이드바의 '캐시 초기화' 클릭
    """)
