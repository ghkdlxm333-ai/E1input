import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="E1 수주/출고 자동 입력 헬퍼", layout="wide")

st.title("📦 E1(JD Edwards) 입력 자동화 & 재고 거름망 시스템")
st.caption("세일즈 리포트 건 중 이미 완료된 건(Order # 6자리)은 자동 제외하며, Inventory On-Hand Report와 매칭하여 바로 복붙 가능한 데이터를 생성합니다.")

# ---------------------------------------------------------
# 1. 파일 업로드 영역
# ---------------------------------------------------------
col_up1, col_up2 = st.columns(2)

with col_up1:
    sales_file = st.file_uploader("1. 세일즈 리포트 업로드 (.xlsx)", type=["xlsx", "xls", "csv"], key="sales")

with col_up2:
    onhand_file = st.file_uploader("2. Inventory On-Hand Report 업로드 (.csv, .xlsx)", type=["csv", "xlsx", "xls"], key="onhand")

# ---------------------------------------------------------
# 2. 데이터 처리 및 매칭 로직
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

    # Order # 6자리 이미 입력 완료된 건 자동 제외 로직
    def is_order_completed(val):
        if pd.isna(val):
            return False
        val_str = str(val).strip().split('.')[0] # 소수점 제거
        return len(val_str) == 6 and val_str.isdigit()

    # 미입력건만 추출
    df_sales['is_completed'] = df_sales['Order #'].apply(is_order_completed)
    df_sales_valid = df_sales[(~df_sales['is_completed']) & (pd.to_numeric(df_sales['수량'], errors='coerce') > 0)].copy()

    st.sidebar.metric("총 출고건수", len(df_sales))
    st.sidebar.metric("입력 필요건수 (Order # 없음)", len(df_sales_valid))
    st.sidebar.metric("완료건수 (자동제외)", len(df_sales) - len(df_sales_valid))

    # B. Inventory On-Hand Report 로드 (인코딩 자동 파악)
    df_onhand_raw = None
    try:
        if onhand_file.name.endswith('.csv'):
            encodings = ['utf-8-sig', 'cp949', 'euc-kr', 'utf-8', 'latin1']
            for enc in encodings:
                try:
                    onhand_file.seek(0)
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
        else:
            df_onhand_raw = pd.read_excel(onhand_file, header=2)
    except Exception as e:
        st.error(f"Inventory On-Hand Report를 읽는 중 오류가 발생했습니다: {e}")
        st.stop()

    if df_onhand_raw is None:
        st.error("Inventory On-Hand Report 읽기에 실패했습니다.")
        st.stop()

    # C. 재고 전처리
    df_onhand_raw.columns = df_onhand_raw.columns.astype(str).str.strip()
    df_onhand = df_onhand_raw.copy()
    df_onhand['On-Hand Qty'] = pd.to_numeric(df_onhand['On-Hand Qty'], errors='coerce').fillna(0)
    df_onhand['Item Number'] = df_onhand['Item Number'].astype(str).str.strip()
    df_onhand['Lot Number'] = df_onhand['Lot Number'].astype(str).str.strip()
    df_onhand['Location'] = df_onhand['Location'].astype(str).str.strip()
    df_onhand['Lot Expiration Date'] = pd.to_datetime(df_onhand['Lot Expiration Date'], errors='coerce')

    # ---------------------------------------------------------
    # 3. 매칭 엔진 (구분별 Location 및 LOT 수량 검증)
    # ---------------------------------------------------------
    # (Item Number, Location, Lot Number) -> Qty
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

        # Location 자동 결정
        if '반품' in category:
            target_location = 'RET'
        elif '이동' in category or '재고이동' in category:
            target_location = 'REP2' # 기본 REP2 설정
        else:
            target_location = 'PRI' # 기본 매출

        # E1 Item Number 매핑 (K 접두사 대응)
        if not df_onhand[df_onhand['Item Number'] == item_code].empty:
            e1_item_code = item_code
        elif not df_onhand[df_onhand['Item Number'] == 'K' + item_code].empty:
            e1_item_code = 'K' + item_code
        else:
            e1_item_code = item_code

        # 해당 품목 & Location 재고 목록 (FEFO 순)
        item_inv = df_onhand[
            (df_onhand['Item Number'] == e1_item_code) & 
            (df_onhand['Location'] == target_location)
        ].sort_values(by='Lot Expiration Date')

        # 1차 시도: 세일즈 LOT 매칭
        key_sales = (e1_item_code, target_location, sales_lot)
        if sales_lot and key_sales in inventory_pool and inventory_pool[key_sales] > 0:
            avail = inventory_pool[key_sales]
            use_qty = min(req_qty, avail)

            status_flag = "NORMAL" if use_qty == req_qty else "SPLIT"

            processed_rows.append({
                'Line Number': '',
                'Item  Number': item_code,
                'Description': s_row.get('제품명', ''),
                'Quantity Ordered': use_qty,
                'Unit Price': unit_price,
                'Extended Price': use_qty * unit_price,
                'Last Status': '',
                'Lot Number': sales_lot,
                'Requested Date': s_row.get('Date', ''),
                'Location': target_location,
                # 조회 필터용 속성
                'Date': s_row.get('Date', ''),
                'Channel': s_row.get('Channel', ''),
                'Customer': s_row.get('Customer', ''),
                'Sold To (bill to)': s_row.get('bill to ', s_row.get('Ship to ', '')),
                'Ship To': s_row.get('Ship to ', ''),
                '상태구분': status_flag,
                '상태메시지': '정상 매칭' if status_flag == "NORMAL" else '수량 부족으로 LOT 분할'
            })

            inventory_pool[key_sales] -= use_qty
            req_qty -= use_qty

        # 2차 시도: FEFO 대체 및 분할
        if req_qty > 0:
            for _, inv_row in item_inv.iterrows():
                cur_lot = inv_row['Lot Number']
                key_cur = (e1_item_code, target_location, cur_lot)
                avail = inventory_pool.get(key_cur, 0)

                if avail <= 0:
                    continue

                use_qty = min(req_qty, avail)
                
                processed_rows.append({
                    'Line Number': '',
                    'Item  Number': item_code,
                    'Description': s_row.get('제품명', ''),
                    'Quantity Ordered': use_qty,
                    'Unit Price': unit_price,
                    'Extended Price': use_qty * unit_price,
                    'Last Status': '',
                    'Lot Number': cur_lot,
                    'Requested Date': s_row.get('Date', ''),
                    'Location': target_location,
                    'Date': s_row.get('Date', ''),
                    'Channel': s_row.get('Channel', ''),
                    'Customer': s_row.get('Customer', ''),
                    'Sold To (bill to)': s_row.get('bill to ', s_row.get('Ship to ', '')),
                    'Ship To': s_row.get('Ship to ', ''),
                    '상태구분': 'REPLACED' if sales_lot != cur_lot else 'SPLIT',
                    '상태메시지': f'LOT 자동 변경 ({sales_lot} ➡️ {cur_lot})'
                })

                inventory_pool[key_cur] -= use_qty
                req_qty -= use_qty

                if req_qty <= 0:
                    break

        # 3차: 재고 전량 부족
        if req_qty > 0:
            processed_rows.append({
                'Line Number': '',
                'Item  Number': item_code,
                'Description': s_row.get('제품명', ''),
                'Quantity Ordered': req_qty,
                'Unit Price': unit_price,
                'Extended Price': req_qty * unit_price,
                'Last Status': '',
                'Lot Number': sales_lot if sales_lot else '재고없음',
                'Requested Date': s_row.get('Date', ''),
                'Location': target_location,
                'Date': s_row.get('Date', ''),
                'Channel': s_row.get('Channel', ''),
                'Customer': s_row.get('Customer', ''),
                'Sold To (bill to)': s_row.get('bill to ', s_row.get('Ship to ', '')),
                'Ship To': s_row.get('Ship to ', ''),
                '상태구분': 'SHORTAGE',
                '상태메시지': f'🚨 E1 {target_location} 재고 부족 ({req_qty}개 부족)'
            })

    df_result = pd.DataFrame(processed_rows)

    # ---------------------------------------------------------
    # 4. 조회 필터 및 E1 클립보드 출력
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("🔍 E1 입력 대상 조회 & 검색 필터")

    f_col1, f_col2, f_col3, f_col4 = st.columns(4)

    with f_col1:
        dates = ['전체'] + sorted(list(df_result['Date'].dropna().astype(str).unique()))
        sel_date = st.selectbox("날짜 선택:", dates)

    with f_col2:
        channels = ['전체'] + list(df_result['Channel'].dropna().unique())
        sel_channel = st.selectbox("채널 선택:", channels)

    with f_col3:
        customers = ['전체'] + list(df_result['Customer'].dropna().unique())
        sel_customer = st.selectbox("Customer 선택:", customers)

    with f_col4:
        sold_tos = ['전체'] + list(df_result['Sold To (bill to)'].dropna().astype(str).unique())
        sel_sold_to = st.selectbox("Sold To (bill to) 선택:", sold_tos)

    # 필터링 적용
    df_filtered = df_result.copy()
    if sel_date != '전체':
        df_filtered = df_filtered[df_filtered['Date'].astype(str) == sel_date]
    if sel_channel != '전체':
        df_filtered = df_filtered[df_filtered['Channel'] == sel_channel]
    if sel_customer != '전체':
        df_filtered = df_filtered[df_filtered['Customer'] == sel_customer]
    if sel_sold_to != '전체':
        df_filtered = df_filtered[df_filtered['Sold To (bill to)'].astype(str) == sel_sold_to]

    # Header 상단 배치
    if not df_filtered.empty:
        h1, h2, h3 = st.columns(3)
        with h1:
            st.info(f"**Sold To:** `{df_filtered['Sold To (bill to)'].iloc[0]}`")
        with h2:
            st.info(f"**Ship To:** `{df_filtered['Ship To'].iloc[0]}`")
        with h3:
            st.info(f"**선택 건수:** `{len(df_filtered)}행`")

    # 테이블 조건부 서식 함수
    def highlight_status(row):
        flag = row['상태구분']
        if flag == 'SHORTAGE':
            return ['background-color: #ffcdd2; color: #b71c1c; font-weight: bold;'] * len(row) # 빨강 (재고부족)
        elif flag in ['SPLIT', 'REPLACED']:
            return ['background-color: #fff9c4; color: #f57f17; font-weight: bold;'] * len(row) # 노랑 (LOT변경/분할)
        else:
            return [''] * len(row)

    st.subheader("📋 세일즈 및 E1 매칭 확인 그리드")
    
    # 노출용 컬럼 정리
    display_cols = [
        'Item  Number', 'Description', 'Quantity Ordered', 'Unit Price', 'Extended Price', 
        'Lot Number', 'Location', '상태메시지'
    ]
    
    st.dataframe(
        df_filtered[display_cols].style.apply(highlight_status, axis=1),
        use_container_width=True
    )

    # E1 입력창 순서 10개 컬럼
    e1_grid_cols = [
        'Line Number', 'Item  Number', 'Description', 'Quantity Ordered', 
        'Unit Price', 'Extended Price', 'Last Status', 'Lot Number', 
        'Requested Date', 'Location'
    ]

    # Date 포맷 정리
    df_e1_copy = df_filtered[e1_grid_cols].copy()
    df_e1_copy['Requested Date'] = pd.to_datetime(df_e1_copy['Requested Date'], errors='coerce').dt.strftime('%Y/%m/%d').fillna('')

    # TSV 생성 (Header 제외, Tab 구분자)
    tsv_output = df_e1_copy.to_csv(sep='\t', index=False, header=False)

    st.markdown("---")
    st.subheader("📋 E1 붙여넣기용 클립보드 (E1 Grid 동일 규격)")
    st.write("👉 아래 박스 클릭 ➡️ `Ctrl+A` ➡️ `Ctrl+C` 후 E1 그리드 첫 칸(`Line Number` 또는 `Item Number`)에 `Ctrl+V` 하세요.")
    
    st.text_area("E1 복붙 텍스트 (Tab 구분)", value=tsv_output, height=200)
