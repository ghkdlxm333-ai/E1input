import streamlit as st
import pandas as pd
import numpy as np

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
    if sales_file_name.endswith(('.xlsx', '.xls')):
        df_sales = pd.read_excel(sales_file_bytes, sheet_name='일일출고')
    else:
        df_sales = pd.read_csv(sales_file_bytes)

    def clean_num(val):
        if pd.isna(val) or val is None:
            return 0
        s = str(val).replace(',', '').strip()
        try:
            return int(float(s))
        except ValueError:
            return 0

    def clean_str(val):
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip()
        return "" if s.lower() in ["none", "nan", "<na>"] else s

    def is_order_completed(val):
        if pd.isna(val):
            return False
        val_str = str(val).strip().split('.')[0]
        return len(val_str) == 6 and val_str.isdigit()

    df_sales['is_completed'] = df_sales['Order #'].apply(is_order_completed)
    df_sales['category_clean'] = df_sales['구분'].astype(str).str.strip()
    is_not_move = ~df_sales['category_clean'].str.contains('이동|재고이동', na=False)
    
    df_sales['수량_num'] = df_sales['수량'].apply(clean_num)
    
    df_sales_valid = df_sales[
        (~df_sales['is_completed']) & 
        (df_sales['수량_num'] > 0) & 
        is_not_move
    ].copy()

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

    df_onhand_raw.columns = df_onhand_raw.columns.astype(str).str.strip()
    df_onhand = df_onhand_raw.copy()
    df_onhand['On-Hand Qty'] = df_onhand['On-Hand Qty'].apply(clean_num)
    df_onhand['Item Number'] = df_onhand['Item Number'].apply(clean_str)
    df_onhand['Lot Number'] = df_onhand['Lot Number'].apply(clean_str)
    df_onhand['Location'] = df_onhand['Location'].apply(clean_str)
    df_onhand['Lot Expiration Date'] = pd.to_datetime(df_onhand['Lot Expiration Date'], errors='coerce')

    inventory_pool = {}
    for _, row in df_onhand[df_onhand['On-Hand Qty'] > 0].iterrows():
        key = (row['Item Number'], row['Location'], row['Lot Number'])
        inventory_pool[key] = inventory_pool.get(key, 0) + row['On-Hand Qty']

    processed_rows = []

    for idx, s_row in df_sales_valid.iterrows():
        item_code = clean_str(s_row.get('제품코드', ''))
        req_qty = clean_num(s_row.get('수량', 0))
        unit_price = clean_num(s_row.get('단가', 0))
        sales_lot = clean_str(s_row.get('LOT', ''))
        category = clean_str(s_row.get('구분', ''))

        raw_date = s_row.get('Date', '')
        clean_date = pd.to_datetime(raw_date, errors='coerce').strftime('%Y-%m-%d') if pd.notna(raw_date) else str(raw_date).split(' ')[0]
        if clean_date.lower() in ['nat', 'none']:
            clean_date = ""

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
            'Customer': clean_str(s_row.get('Customer', '')),
            'bill to': clean_str(s_row.get('bill to ', s_row.get('Ship to ', ''))),
            'Ship to': clean_str(s_row.get('Ship to ', '')),
            '제품코드': item_code,
            '제품명': clean_str(s_row.get('제품명', '')),
            '수량': req_qty,
            '단가': unit_price,
            'Total Amount': req_qty * unit_price,
            '매입확인': clean_str(s_row.get('매입확인', '')),
            'Channel': clean_str(s_row.get('Channel', '')),
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
                'Total Amount': int(use_qty * unit_price),
                'LOT': sales_lot,
                '상태구분': status_flag,
                '상태메시지': '정상 매칭' if status_flag == "NORMAL" else f'⚠️ LOT 분할 ({sales_lot})'
            })
            processed_rows.append(row_data)

            inventory_pool[key_sales] -= use_qty
            req_qty -= use_qty

        # 2차 매칭 (FIFO 선입선출)
        if req_qty > 0:
            for _, inv_row in item_inv.iterrows():
                cur_lot = inv_row['Lot Number']
                key_cur = (e1_item_code, target_location, cur_lot)
                avail = inventory_pool.get(key_cur, 0)

                if avail <= 0:
                    continue

                use_qty = min(req_qty, avail)
                msg = f'⚠️ LOT 분할 선입선출 ({cur_lot})'
                st_flag = 'SPLIT'

                row_data = sales_base.copy()
                row_data.update({
                    '수량': int(use_qty),
                    'Total Amount': int(use_qty * unit_price),
                    'LOT': cur_lot,
                    '상태구분': st_flag,
                    '상태메시지': msg
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
                'Total Amount': int(req_qty * unit_price),
                'LOT': sales_lot if sales_lot else '재고없음',
                '상태구분': 'SHORTAGE',
                '상태메시지': f'🚨 E1 {target_location} 재고 부족 ({int(req_qty)}개 부족)'
            })
            processed_rows.append(row_data)

    return pd.DataFrame(processed_rows)


# ---------------------------------------------------------
# 행 색상 강조 스타일 지정 함수
# ---------------------------------------------------------
def highlight_status(row):
    status = str(row.get('상태구분', ''))
    if status == 'SPLIT':
        return ['background-color: #FFF3CD; color: #856404; font-weight: bold;'] * len(row) # 노란색 (LOT 분할만 강조)
    elif status == 'SHORTAGE':
        return ['background-color: #F8D7DA; color: #721C24; font-weight: bold;'] * len(row) # 빨간색 (재고 부족)
    return [''] * len(row)


# ---------------------------------------------------------
# 3. 메인 화면 구성 및 인터랙션
# ---------------------------------------------------------
if sales_file and onhand_file:
    df_result = process_data(sales_file, sales_file.name, onhand_file, onhand_file.name)

    if df_result is None or df_result.empty:
        st.error("처리 가능한 데이터가 없습니다.")
        st.stop()

    # --- 사이드바 필터 설정 ---
    st.sidebar.header("🔍 조회 조건 필터")

    # 1. 구분 필터 (중복 선택 가능 - 거래처보다 위로 배치)
    categories = sorted([cat for cat in df_result['구분'].unique() if pd.notna(cat) and str(cat).strip() != ''])
    selected_cats = st.sidebar.multiselect("🏷️ 구분 (중복 선택 가능):", options=categories, default=categories)

    # 2. 거래처 선택
    customers = sorted([c for c in df_result['Customer'].unique() if pd.notna(c) and str(c).strip() != ''])
    selected_cust = st.sidebar.radio("🏢 거래처 (Customer) 선택:", options=["📊 전체 모아보기"] + customers)

    # 3. Date 날짜 범위 필터
    df_result['Date_dt'] = pd.to_datetime(df_result['Date'], errors='coerce')
    valid_dates = df_result['Date_dt'].dropna()

    if not valid_dates.empty:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        selected_date_range = st.sidebar.date_input(
            "📅 Date (날짜 범위 필터):",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )
    else:
        selected_date_range = None

    # --- 데이터 필터링 적용 ---
    df_curr = df_result.copy()

    # 구분 필터링
    if selected_cats:
        df_curr = df_curr[df_curr['구분'].isin(selected_cats)]
    else:
        df_curr = df_curr.iloc[0:0]

    # 거래처 필터링
    if selected_cust != "📊 전체 모아보기":
        df_curr = df_curr[df_curr['Customer'] == selected_cust]

    # Date 필터링
    if selected_date_range and len(selected_date_range) == 2:
        start_d, end_d = selected_date_range
        df_curr = df_curr[
            (df_curr['Date_dt'].dt.date >= start_d) & 
            (df_curr['Date_dt'].dt.date <= end_d)
        ]

    df_curr = df_curr.reset_index(drop=True)

    tot_qty = df_curr['수량'].sum() if not df_curr.empty else 0
    tot_amt = df_curr['Total Amount'].sum() if not df_curr.empty else 0
    
    st.subheader(f"📌 [{selected_cust}] 세일즈 리포트 매칭 결과")
    st.info(f"💡 총 **{len(df_curr):,}건**  |  **총 수량:** `{tot_qty:,} 개`  |  **총 금액:** `{tot_amt:,} 원`  |  🟡 **노란색:** LOT 분할  |  🔴 **빨간색:** 재고 부족 (복붙 제외됨)")

    # ---------------------------------------------------------
    # 1️⃣ 상단 세일즈 리포트 결과 (상태 점검용)
    # ---------------------------------------------------------
    sales_cols = ['구분', 'Date', 'Customer', 'bill to', 'Ship to', '제품코드', '제품명', '수량', '단가', 'Total Amount', '매입확인', 'LOT', 'Location', '상태메시지', '상태구분']
    df_sales_disp = df_curr[sales_cols].copy()

    styled_sales_df = df_sales_disp.style.apply(highlight_status, axis=1)

    st.dataframe(
        styled_sales_df,
        use_container_width=True,
        hide_index=True,
        column_config={"상태구분": None}
    )

    # ---------------------------------------------------------
    # 2️⃣ 하단 E1 입력창 복붙용 클립보드 (요약 및 표)
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("2️⃣ E1 입력창 복붙용 클립보드 표")

    # 재고 부족(SHORTAGE) 건 제외
    df_e1_valid = df_curr[df_curr['상태구분'] != 'SHORTAGE'].copy().reset_index(drop=True)

    e1_tot_cnt = len(df_e1_valid)
    e1_tot_qty = df_e1_valid['수량'].sum() if not df_e1_valid.empty else 0
    e1_tot_amt = df_e1_valid['Total Amount'].sum() if not df_e1_valid.empty else 0

    # 복붙용 상단 집계 요약 카드
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("📋 E1 복붙 대상 건수", f"{e1_tot_cnt:,} 건")
    col_m2.metric("📦 E1 복붙 총 수량", f"{e1_tot_qty:,} 개")
    col_m3.metric("💰 E1 복붙 총 금액", f"{e1_tot_amt:,} 원")

    df_e1 = pd.DataFrame()
    if not df_e1_valid.empty:
        df_e1['Line Number'] = ""
        df_e1['Item  Number'] = df_e1_valid['제품코드'].astype(str)
        df_e1['Description'] = ""
        df_e1['Quantity Ordered'] = df_e1_valid['수량'].astype(int)
        df_e1['Unit Price'] = df_e1_valid['단가'].astype(int)
        df_e1['Extended Price'] = df_e1_valid['Total Amount'].astype(int)
        df_e1['Last Status'] = ""
        df_e1['Lot Number'] = df_e1_valid['LOT'].astype(str)
        df_e1['Requested Date'] = ""
        df_e1['Location'] = df_e1_valid['Location'].astype(str)
        df_e1['상태구분'] = df_e1_valid['상태구분']

        # 결측치 최종 제거
        df_e1 = df_e1.replace({'None': '', 'nan': '', 'NaN': '', np.nan: ''})

    styled_e1_df = df_e1.style.apply(highlight_status, axis=1)

    st.success(f"✅ E1 입력 준비 완료. 아래 표를 마우스로 드래그 후 `Ctrl+C` 하세요.")

    st.dataframe(
        styled_e1_df,
        use_container_width=True,
        hide_index=True,
        column_config={"상태구분": None}
    )

    tsv_data = df_e1.drop(columns=['상태구분'], errors='ignore').to_csv(sep='\t', index=False, header=False).encode('utf-8-sig')
    st.download_button(
        label="📥 E1 복붙용 파일 다운로드 (.tsv)",
        data=tsv_data,
        file_name=f"E1_Upload_{selected_cust}.tsv",
        mime="text/tab-separated-values"
    )
