import streamlit as st
import pandas as pd
import numpy as np
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

st.set_page_config(page_title="E1 수주/출고 자동 입력 헬퍼", layout="wide")

st.title("📦 E1(JD Edwards) 입력 자동화 & 재고 거름망 시스템")
st.caption("세일즈 리포트의 매출/반품 건을 확인하고, Inventory On-Hand Report와 매칭하여 E1에 즉시 복붙 가능한 클립보드를 생성합니다.")

# ---------------------------------------------------------
# 1. 파일 업로드 영역
# ---------------------------------------------------------
col_up1, col_up2 = st.columns(2)

with col_up1:
    sales_file = st.file_uploader("1. 세일즈 리포트 업로드 (.xlsx)", type=["xlsx", "xls", "csv"], key="sales")

with col_up2:
    onhand_file = st.file_uploader("2. Inventory On-Hand Report 업로드 (.csv, .xlsx)", type=["csv", "xlsx", "xls"], key="onhand")

# ---------------------------------------------------------
# 2. 캐싱 처리된 데이터 매칭 엔진
# ---------------------------------------------------------
@st.cache_data(show_spinner="데이터 매칭 처리 중...")
def process_data(sales_file_bytes, sales_file_name, onhand_file_bytes, onhand_file_name):
    # A. 세일즈 리포트 로드
    if sales_file_name.endswith(('.xlsx', '.xls')):
        df_sales = pd.read_excel(sales_file_bytes, sheet_name='일일출고')
    else:
        df_sales = pd.read_csv(sales_file_bytes)

    def is_order_completed(val):
        if pd.isna(val):
            return False
        val_str = str(val).strip().split('.')[0]
        return len(val_str) == 6 and val_str.isdigit()

    df_sales['is_completed'] = df_sales['Order #'].apply(is_order_completed)
    df_sales['category_clean'] = df_sales['구분'].astype(str).str.strip()
    is_not_move = ~df_sales['category_clean'].str.contains('이동|재고이동', na=False)
    
    df_sales_valid = df_sales[
        (~df_sales['is_completed']) & 
        (pd.to_numeric(df_sales['수량'], errors='coerce') > 0) & 
        is_not_move
    ].copy()

    # B. Inventory On-Hand Report 로드
    df_onhand_raw = None
    if onhand_file_name.endswith('.csv'):
        encodings = ['utf-8-sig', 'cp949', 'euc-kr', 'utf-8', 'latin1']
        for enc in encodings:
            try:
                onhand_file_bytes.seek(0)
                temp_df = pd.read_csv(onhand_file_bytes, encoding=enc, header=None, nrows=10)
                header_row_idx = None
                for idx, row in temp_df.iterrows():
                    row_str = " ".join(row.dropna().astype(str))
                    if 'Branch' in row_str and 'Item Number' in row_str:
                        header_row_idx = idx
                        break
                
                if header_row_idx is not None:
                    onhand_file_bytes.seek(0)
                    df_onhand_raw = pd.read_csv(onhand_file_bytes, encoding=enc, header=header_row_idx)
                    break
            except Exception:
                continue
    else:
        df_onhand_raw = pd.read_excel(onhand_file_bytes, header=2)

    if df_onhand_raw is None:
        return None

    # C. 재고 데이터 전처리
    df_onhand_raw.columns = df_onhand_raw.columns.astype(str).str.strip()
    df_onhand = df_onhand_raw.copy()
    df_onhand['On-Hand Qty'] = pd.to_numeric(df_onhand['On-Hand Qty'], errors='coerce').fillna(0)
    df_onhand['Item Number'] = df_onhand['Item Number'].astype(str).str.strip()
    df_onhand['Lot Number'] = df_onhand['Lot Number'].astype(str).str.strip()
    df_onhand['Location'] = df_onhand['Location'].astype(str).str.strip()
    df_onhand['Lot Expiration Date'] = pd.to_datetime(df_onhand['Lot Expiration Date'], errors='coerce')

    inventory_pool = {}
    for _, row in df_onhand[df_onhand['On-Hand Qty'] > 0].iterrows():
        key = (row['Item Number'], row['Location'], row['Lot Number'])
        inventory_pool[key] = inventory_pool.get(key, 0) + row['On-Hand Qty']

    processed_rows = []

    for idx, s_row in df_sales_valid.iterrows():
        item_code = str(s_row['제품코드']).strip()
        req_qty = float(s_row['수량'])
        unit_price = float(s_row.get('단가', 0))
        sales_lot = str(s_row.get('LOT', '')).strip() if pd.notna(s_row.get('LOT')) else ''
        category = str(s_row.get('구분', '')).strip()

        raw_date = s_row.get('Date', '')
        clean_date = pd.to_datetime(raw_date, errors='coerce').strftime('%Y-%m-%d') if pd.notna(raw_date) else str(raw_date).split(' ')[0]

        target_location = 'RET' if '반품' in category else 'PRI'

        if not df_onhand[df_onhand['Item Number'] == item_code].empty:
            e1_item_code = item_code
        elif not df_onhand[df_onhand['Item Number'] == 'K' + item_code].empty:
            e1_item_code = 'K' + item_code
        else:
            e1_item_code = item_code

        item_inv = df_onhand[
            (df_onhand['Item Number'] == e1_item_code) & 
            (df_onhand['Location'] == target_location)
        ].sort_values(by='Lot Expiration Date')

        sales_base = {
            '구분': category,
            'Date': clean_date,
            'Customer': str(s_row.get('Customer', '')).strip(),
            'bill to': str(s_row.get('bill to ', s_row.get('Ship to ', ''))).strip(),
            'Ship to': str(s_row.get('Ship to ', '')).strip(),
            '제품코드': item_code,
            '제품명': s_row.get('제품명', ''),
            '단가': unit_price,
            'Total Amount': int(req_qty * unit_price),
            '매입확인': str(s_row.get('매입확인', '')).strip(),
            'Channel': str(s_row.get('Channel', '')).strip(),
            'Location': target_location
        }

        # 1차 매칭
        key_sales = (e1_item_code, target_location, sales_lot)
        if sales_lot and key_sales in inventory_pool and inventory_pool[key_sales] > 0:
            avail = inventory_pool[key_sales]
            use_qty = min(req_qty, avail)
            status_flag = "NORMAL" if use_qty == req_qty else "SPLIT"

            row_data = sales_base.copy()
            row_data.update({
                '수량': int(use_qty),
                'LOT': sales_lot,
                '상태구분': status_flag,
                '상태메시지': '정상 매칭' if status_flag == "NORMAL" else '수량 부족으로 LOT 분할'
            })
            processed_rows.append(row_data)

            inventory_pool[key_sales] -= use_qty
            req_qty -= use_qty

        # 2차 매칭
        if req_qty > 0:
            for _, inv_row in item_inv.iterrows():
                cur_lot = inv_row['Lot Number']
                key_cur = (e1_item_code, target_location, cur_lot)
                avail = inventory_pool.get(key_cur, 0)

                if avail <= 0:
                    continue

                use_qty = min(req_qty, avail)
                
                row_data = sales_base.copy()
                row_data.update({
                    '수량': int(use_qty),
                    'LOT': cur_lot,
                    '상태구분': 'REPLACED' if sales_lot != cur_lot else 'SPLIT',
                    '상태메시지': f'LOT 자동 변경 ({sales_lot} ➡️ {cur_lot})'
                })
                processed_rows.append(row_data)

                inventory_pool[key_cur] -= use_qty
                req_qty -= use_qty

                if req_qty <= 0:
                    break

        # 3차 매칭 (재고 부족)
        if req_qty > 0:
            row_data = sales_base.copy()
            row_data.update({
                '수량': int(req_qty),
                'LOT': sales_lot if sales_lot else '재고없음',
                '상태구분': 'SHORTAGE',
                '상태메시지': f'🚨 E1 {target_location} 재고 부족 ({int(req_qty)}개 부족)'
            })
            processed_rows.append(row_data)

    return pd.DataFrame(processed_rows)


# ---------------------------------------------------------
# 3. 데이터 표시 (초고속 AgGrid 전용 그리드 적용)
# ---------------------------------------------------------
if sales_file and onhand_file:
    df_result = process_data(sales_file, sales_file.name, onhand_file, onhand_file.name)

    if df_result is None or df_result.empty:
        st.error("처리 가능한 데이터가 없습니다.")
        st.stop()

    st.markdown("---")
    st.subheader("1️⃣ 세일즈 리포트 확인 그리드 (초고속 AgGrid)")
    st.caption("💡 각 열 헤더의 **필터 아이콘**으로 엑셀처럼 필터링하고, **셀/범위를 드래그**하면 우측 하단에 **합계/평균이 자동 계산**됩니다.")

    sales_report_cols = [
        '구분', 'Date', 'Customer', 'bill to', 'Ship to', 
        '제품코드', '제품명', '수량', '단가', 'Total Amount', '매입확인', 'LOT', '상태메시지'
    ]
    df_sales_disp = df_result[sales_report_cols].copy()

    # 🚀 AgGrid 옵션 설정 (엑셀 기능 활성화)
    gb1 = GridOptionsBuilder.from_dataframe(df_sales_disp)
    gb1.configure_default_column(filterable=True, sortable=True, resizable=True)
    gb1.configure_selection(selection_mode="multiple", use_checkbox=True)
    gb1.configure_grid_options(enableRangeSelection=True, enableStatusBar=True) # 범위 선택 및 하단 상태바 활성화
    
    gridOptions1 = gb1.build()

    AgGrid(
        df_sales_disp,
        gridOptions=gridOptions1,
        height=350,
        theme='balham', # 엑셀과 비슷한 깔끔한 테마
        fit_columns_on_grid_load=False
    )

    # ---------------------------------------------------------
    # 2nd GRID: E1 복붙용 그리드
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("2️⃣ E1 입력창 복붙용 클립보드 그리드")
    st.caption("📋 원하는 행을 드래그하여 선택 후 `Ctrl+C` 하여 E1에 붙여넣으세요.")

    df_e1 = pd.DataFrame()
    df_e1['Line Number'] = [''] * len(df_result)
    df_e1['Item  Number'] = df_result['제품코드']
    df_e1['Description'] = [''] * len(df_result)
    df_e1['Quantity Ordered'] = df_result['수량'].astype(int)
    df_e1['Unit Price'] = df_result['단가']
    df_e1['Extended Price'] = (df_result['수량'] * df_result['단가']).astype(int)
    df_e1['Last Status'] = [''] * len(df_result)
    df_e1['Lot Number'] = df_result['LOT']
    df_e1['Requested Date'] = [''] * len(df_result)
    df_e1['Location'] = df_result['Location']

    gb2 = GridOptionsBuilder.from_dataframe(df_e1)
    gb2.configure_default_column(filterable=True, sortable=True, resizable=True)
    gb2.configure_selection(selection_mode="multiple", use_checkbox=True)
    gb2.configure_grid_options(enableRangeSelection=True, enableStatusBar=True)
    
    gridOptions2 = gb2.build()

    AgGrid(
        df_e1,
        gridOptions=gridOptions2,
        height=350,
        theme='balham',
        fit_columns_on_grid_load=False
    )

    # 📥 TSV 보조 다운로드
    tsv_data = df_e1.to_csv(sep='\t', index=False, header=False).encode('utf-8-sig')
    st.download_button(
        label="📥 E1 복붙용 데이터 파일 다운로드 (.tsv)",
        data=tsv_data,
        file_name="E1_Upload_Data.tsv",
        mime="text/tab-separated-values"
    )
