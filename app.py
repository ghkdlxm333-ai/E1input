import streamlit as st
import pandas as pd

st.set_page_config(page_title="E1 수주/출고 자동 매칭 시스템", layout="wide")

st.title("📦 E1(JD Edwards) 실재고 연동 Auto-LOT 매칭 도구")
st.caption("세일즈 리포트와 E1 Inventory On-Hand Report를 기반으로 재고 부족 에러 없는 E1 입력 데이터를 생성합니다.")

# ---------------------------------------------------------
# 1. 파일 업로드 영역
# ---------------------------------------------------------
col_up1, col_up2 = st.columns(2)

with col_up1:
    sales_file = st.file_uploader("1. 세일즈 리포트 업로드 (.xlsx)", type=["xlsx", "xls", "csv"], key="sales")

with col_up2:
    onhand_file = st.file_uploader("2. Inventory On-Hand Report 업로드 (.csv, .xlsx)", type=["csv", "xlsx", "xls"], key="onhand")

# ---------------------------------------------------------
# 2. 데이터 처리 및 Auto-LOT 매칭 로직
# ---------------------------------------------------------
if sales_file and onhand_file:
    # A. 세일즈 리포트 로드
    try:
        if sales_file.name.endswith(('.xlsx', '.xls')):
            df_sales = pd.read_excel(sales_file, sheet_name='일일출고')
        else:
            df_sales = pd.read_csv(sales_file)
    except Exception as e:
        st.error(f"세일즈 리포트를 읽는 중 오류가 발생했습니다: {e}")
        st.stop()

    # B. Inventory On-Hand Report 로드 및 헤더 자동 인식 (인코딩 자동 시도)
    df_onhand_raw = None
    try:
        if onhand_file.name.endswith('.csv'):
            encodings = ['utf-8-sig', 'cp949', 'euc-kr', 'utf-8', 'latin1']
            for enc in encodings:
                try:
                    onhand_file.seek(0)
                    # 헤더 위치 자동 검색 (Branch 또는 Item Number가 포함된 행 탐지)
                    temp_df = pd.read_csv(onhand_file, encoding=enc, header=None, nrows=10)
                    header_row_idx = None
                    for idx, row in temp_df.iterrows():
                        row_str = " ".join(row.dropna().astype(str))
                        if 'Branch' in row_str and 'Item Number' in row_str:
                            header_row_idx = idx
                            break
                    
                    if header_row_idx is not None:
                        onhand_file.seek(0)
                        df_onhand_raw = pd.read_csv(onhand_file, encoding=enc, header=header_row_idx)
                        break
                except Exception:
                    continue

            if df_onhand_raw is None:
                st.error("Inventory On-Hand Report 파일의 인코딩 및 헤더 위치를 해석할 수 없습니다.")
                st.stop()
        else:
            df_onhand_raw = pd.read_excel(onhand_file, header=2)
    except Exception as e:
        st.error(f"Inventory On-Hand Report를 읽는 중 오류가 발생했습니다: {e}")
        st.stop()

    # C. 재고 데이터 전처리 (Location == 'PRI' & On-Hand Qty > 0)
    df_onhand_raw.columns = df_onhand_raw.columns.astype(str).str.strip()
    
    req_cols = ['Item Number', 'Location', 'Lot Number', 'Lot Expiration Date', 'On-Hand Qty']
    if not all(col in df_onhand_raw.columns for col in req_cols):
        st.error(f"Inventory On-Hand Report에 필수 컬럼이 누락되었습니다. 필요한 컬럼: {req_cols}")
        st.stop()

    df_pri = df_onhand_raw[
        (df_onhand_raw['Location'] == 'PRI') & 
        (pd.to_numeric(df_onhand_raw['On-Hand Qty'], errors='coerce') > 0)
    ].copy()

    df_pri['On-Hand Qty'] = pd.to_numeric(df_pri['On-Hand Qty'])
    df_pri['Item Number'] = df_pri['Item Number'].astype(str).str.strip()
    df_pri['Lot Number'] = df_pri['Lot Number'].astype(str).str.strip()
    df_pri['Lot Expiration Date'] = pd.to_datetime(df_pri['Lot Expiration Date'], errors='coerce')

    # D. 세일즈 데이터 필터링 (수량 > 0)
    df_sales['수량'] = pd.to_numeric(df_sales['수량'], errors='coerce').fillna(0)
    df_sales_valid = df_sales[df_sales['수량'] > 0].copy()

    # ---------------------------------------------------------
    # 3. Smart LOT Matching Engine
    # ---------------------------------------------------------
    inventory_pool = {}
    for _, row in df_pri.iterrows():
        key = (row['Item Number'], row['Lot Number'])
        inventory_pool[key] = inventory_pool.get(key, 0) + row['On-Hand Qty']

    matched_rows = []
    shortage_alerts = []

    for idx, s_row in df_sales_valid.iterrows():
        item_code = str(s_row['제품코드']).strip()
        
        # 제품코드 접두사 K 교정
        if not df_pri[df_pri['Item Number'] == item_code].empty:
            e1_item_code = item_code
        elif not df_pri[df_pri['Item Number'] == 'K' + item_code].empty:
            e1_item_code = 'K' + item_code
        else:
            e1_item_code = item_code

        req_qty = float(s_row['수량'])
        unit_price = float(s_row.get('단가', 0))
        sales_lot = str(s_row.get('LOT', '')).strip() if pd.notna(s_row.get('LOT')) else ''

        item_inventory = df_pri[df_pri['Item Number'] == e1_item_code].sort_values(by='Lot Expiration Date')

        # 1차: 세일즈 LOT에 가용재고 존재 여부 확인
        if sales_lot and (e1_item_code, sales_lot) in inventory_pool and inventory_pool[(e1_item_code, sales_lot)] > 0:
            avail = inventory_pool[(e1_item_code, sales_lot)]
            use_qty = min(req_qty, avail)
            
            matched_rows.append({
                'Item Number': item_code,
                'Quantity Ordered': use_qty,
                'Unit Price': unit_price,
                'Extended Price': use_qty * unit_price,
                'Last Status': '',
                'Lot Number': sales_lot,
                'Requested Date': s_row.get('Date', ''),
                'Location': 'PRI',
                'Branch/Plant': 'KW',
                'Sold To': s_row.get('bill to ', s_row.get('Ship to ', '')),
                'Ship To': s_row.get('Ship to ', ''),
                'Customer': s_row.get('Customer', ''),
                '매칭상태': '세일즈 LOT 일치'
            })
            
            inventory_pool[(e1_item_code, sales_lot)] -= use_qty
            req_qty -= use_qty

        # 2차: FEFO(유통기한 임박순) 기준 자동 차감 및 행 분할(Split)
        if req_qty > 0:
            for _, inv_row in item_inventory.iterrows():
                cur_lot = inv_row['Lot Number']
                avail = inventory_pool.get((e1_item_code, cur_lot), 0)

                if avail <= 0:
                    continue

                use_qty = min(req_qty, avail)
                
                matched_rows.append({
                    'Item Number': item_code,
                    'Quantity Ordered': use_qty,
                    'Unit Price': unit_price,
                    'Extended Price': use_qty * unit_price,
                    'Last Status': '',
                    'Lot Number': cur_lot,
                    'Requested Date': s_row.get('Date', ''),
                    'Location': 'PRI',
                    'Branch/Plant': 'KW',
                    'Sold To': s_row.get('bill to ', s_row.get('Ship to ', '')),
                    'Ship To': s_row.get('Ship to ', ''),
                    'Customer': s_row.get('Customer', ''),
                    '매칭상태': 'FEFO 자동대체' if sales_lot != cur_lot else '세일즈 LOT 분할매칭'
                })

                inventory_pool[(e1_item_code, cur_lot)] -= use_qty
                req_qty -= use_qty

                if req_qty <= 0:
                    break

        # 3차: 재고 부족 경고 수집
        if req_qty > 0:
            shortage_alerts.append({
                '제품코드': item_code,
                '제품명': s_row.get('제품명', ''),
                '부족수량': req_qty,
                '고객사': s_row.get('Customer', '')
            })

    df_result = pd.DataFrame(matched_rows)

    # ---------------------------------------------------------
    # 4. 결과 화면 및 복붙 데이터 제공
    # ---------------------------------------------------------
    if shortage_alerts:
        st.error("⚠️ [재고 부족 경고] 아래 품목은 Inventory On-Hand Report의 PRI 로케이션 재고가 부족하여 매칭되지 못했습니다.")
        st.dataframe(pd.DataFrame(shortage_alerts))

    st.markdown("---")
    st.subheader("📋 1. E1 입력 데이터 필터링 (거래처 / 날짜별)")

    customers = ['전체'] + list(df_result['Customer'].dropna().unique())
    selected_customer = st.selectbox("거래처(Customer) 선택:", customers)

    df_filtered = df_result if selected_customer == '전체' else df_result[df_result['Customer'] == selected_customer]

    if not df_filtered.empty:
        col_h1, col_h2, col_h3 = st.columns(3)
        with col_h1:
            st.info(f"**Sold To:** `{df_filtered['Sold To'].iloc[0]}`")
        with col_h2:
            st.info(f"**Ship To:** `{df_filtered['Ship To'].iloc[0]}`")
        with col_h3:
            st.info("**Branch/Plant:** `KW`")

    st.subheader("📋 2. E1 Detail 그리드 매칭 결과")
    e1_grid_cols = ['Item Number', 'Quantity Ordered', 'Unit Price', 'Extended Price', 'Last Status', 'Lot Number', 'Location', 'Branch/Plant', '매칭상태']
    st.dataframe(df_filtered[e1_grid_cols], use_container_width=True)

    st.subheader("📋 3. E1 붙여넣기용 클립보드 데이터 (TSV)")
    tsv_cols = ['Item Number', 'Quantity Ordered', 'Unit Price', 'Extended Price', 'Last Status', 'Lot Number', 'Requested Date', 'Location', 'Branch/Plant']
    
    df_tsv = df_filtered[tsv_cols].copy()
    df_tsv['Requested Date'] = pd.to_datetime(df_tsv['Requested Date'], errors='coerce').dt.strftime('%Y/%m/%d').fillna('')

    tsv_data = df_tsv.to_csv(sep='\t', index=False, header=False)

    st.write("👉 아래 텍스트 박스를 **클릭** 후 `Ctrl+A` ➡️ `Ctrl+C` 하시고, E1 Detail 그리드의 `Item Number` 첫 칸에서 `Ctrl+V` 하세요.")
    st.text_area("E1 복붙용 데이터 (Tab 구분)", value=tsv_data, height=200)
