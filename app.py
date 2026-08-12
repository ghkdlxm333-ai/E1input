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
# 2. 캐싱 처리된 데이터 매칭 엔진 (None 값 원천 차단)
# ---------------------------------------------------------
@st.cache_data(show_spinner="데이터 매칭 처리 중...")
def process_data(sales_file_bytes, sales_file_name, onhand_file_bytes, onhand_file_name):
    if sales_file_name.endswith(('.xlsx', '.xls')):
        df_sales = pd.read_excel(sales_file_bytes, sheet_name='일일출고')
    else:
        df_sales = pd.read_csv(sales_file_bytes)

    # 문자열 처리 및 숫자 변환 보정 함수
    def clean_num(val):
        if pd.isna(val):
            return 0.0
        s = str(val).replace(',', '').strip()
        try:
            return float(s)
        except ValueError:
            return 0.0

    def clean_str(val):
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip()
        return "" if s.lower() == "none" or s.lower() == "nan" else s

    def is_order_completed(val):
        if pd.isna(val):
            return False
        val_str = str(val).strip().split('.')[0]
        return len(val_str) == 6 and val_str.isdigit()

    df_sales['is_completed'] = df_sales['Order #'].apply(is_order_completed)
    df_sales['category_clean'] = df_sales['구분'].astype(str).str.strip()
    is_not_move = ~df_sales['category_clean'].str.contains('이동|재고이동', na=False)
    
    # 수량 전처리
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
        if clean_date.lower() == 'nat' or clean_date.lower() == 'none':
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
            '수량': int(req_qty),
            '단가': int(unit_price),
            'Total Amount': int(req_qty * unit_price),
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
                '상태메시지': '정상 매칭' if status_flag == "NORMAL" else f'⚠️ 수량 부족으로 LOT 분할 ({sales_lot})'
            })
            processed_rows.append(row_data)

            inventory_pool[key_sales] -= use_qty
            req_qty -= use_qty

        # 2차 매칭 (FIFO)
        if req_qty > 0:
            for _, inv_row in item_inv.iterrows():
                cur_lot = inv_row['Lot Number']
                key_cur = (e1_item_code, target_location, cur_lot)
                avail = inventory_pool.get(key_cur, 0)

                if avail <= 0:
                    continue

                use_qty = min(req_qty, avail)
                
                if sales_lot != cur_lot and sales_lot != '':
                    msg = f'⚠️ LOT 변경 ({sales_lot} ➡️ {cur_lot})'
                    st_flag = 'REPLACED'
                else:
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
# 3. 메인 화면 구성
# ---------------------------------------------------------
if sales_file and onhand_file:
    df_result = process_data(sales_file, sales_file.name, onhand_file, onhand_file.name)

    if df_result is None or df_result.empty:
        st.error("처리 가능한 데이터가 없습니다.")
        st.stop()

    customers = sorted([c for c in df_result['Customer'].unique() if pd.notna(c) and str(c).strip() != ''])
    
    st.sidebar.header("🏢 거래처 (Customer) 선택")
    selected_cust = st.sidebar.radio(
        "입력할 거래처를 선택하세요:",
        options=["📊 전체 모아보기"] + customers
    )

    if selected_cust == "📊 전체 모아보기":
        df_curr = df_result.copy()
    else:
        df_curr = df_result[df_result['Customer'] == selected_cust].copy()

    tot_qty = df_curr['수량'].sum()
    tot_amt = df_curr['Total Amount'].sum()
    
    st.subheader(f"📌 [{selected_cust}] 세일즈 리포트 확인 및 선택")
    st.info(f"💡 총 **{len(df_curr):,}건**  |  **총 수량:** `{tot_qty:,} 개`  |  **총 금액:** `{tot_amt:,} 원`")

    # ---------------------------------------------------------
    # 1️⃣ 순정 Streamlit data_editor
    # ---------------------------------------------------------
    sales_cols = ['구분', 'Date', 'Customer', 'bill to', 'Ship to', '제품코드', '제품명', '수량', '단가', 'Total Amount', '매입확인', 'LOT', 'Location', '상태메시지']
    
    df_sales_disp = df_curr[sales_cols].copy()
    df_sales_disp.insert(0, "선택", True)

    column_config = {
        "선택": st.column_config.CheckboxColumn("선택", default=True, width="small"),
        "구분": st.column_config.TextColumn("구분", width="small"),
        "Date": st.column_config.TextColumn("Date", width="small"),
        "Customer": st.column_config.TextColumn("Customer", width="medium"),
        "bill to": st.column_config.TextColumn("bill to", width="medium"),
        "Ship to": st.column_config.TextColumn("Ship to", width="medium"),
        "제품코드": st.column_config.TextColumn("제품코드", width="small"),
        "제품명": st.column_config.TextColumn("제품명", width="large"),
        "수량": st.column_config.NumberColumn("수량", width="small", format="%d"),
        "단가": st.column_config.NumberColumn("단가", width="small", format="%d"),
        "Total Amount": st.column_config.NumberColumn("Total Amount", width="medium", format="%d"),
        "매입확인": st.column_config.TextColumn("매입확인", width="small"),
        "LOT": st.column_config.TextColumn("LOT", width="medium"),
        "Location": st.column_config.TextColumn("Location", width="small"),
        "상태메시지": st.column_config.TextColumn("상태메시지", width="large"),
    }

    edited_df = st.data_editor(
        df_sales_disp,
        column_config=column_config,
        hide_index=True,
        use_container_width=True,
        key=f"editor_{selected_cust}"
    )

    st.caption("💡 **Tip:** 표 상단 왼쪽 체크박스를 선택/해제하여 입력할 항목을 제어할 수 있습니다.")

    # ---------------------------------------------------------
    # 2️⃣ 복붙용 클립보드 생성 영역 (None 완전 제거 보장)
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("2️⃣ E1 입력창 복붙용 클립보드 생성")

    if st.button("📋 선택한 행으로 E1 복붙용 클립보드 표 생성하기", type="primary", use_container_width=True):
        df_selected = edited_df[edited_df["선택"] == True].copy()

        if df_selected.empty:
            st.warning("⚠️ 위 표에서 E1에 입력할 행을 1개 이상 체크해 주세요.")
        else:
            df_e1 = pd.DataFrame()
            
            # None 방지를 위한 안전 변환 함수
            def safe_col(df, col_name, is_num=False):
                if col_name not in df.columns:
                    return [0 if is_num else ""] * len(df)
                s = df[col_name].fillna(0 if is_num else "")
                if is_num:
                    return pd.to_numeric(s, errors='coerce').fillna(0).astype(int)
                return s.astype(str).replace({'None': '', 'nan': '', 'NaN': ''})

            df_e1['Line Number'] = [''] * len(df_selected)
            df_e1['Item  Number'] = safe_col(df_selected, '제품코드')
            df_e1['Description'] = [''] * len(df_selected)
            df_e1['Quantity Ordered'] = safe_col(df_selected, '수량', is_num=True)
            df_e1['Unit Price'] = safe_col(df_selected, '단가', is_num=True)
            df_e1['Extended Price'] = df_e1['Quantity Ordered'] * df_e1['Unit Price']
            df_e1['Last Status'] = [''] * len(df_selected)
            df_e1['Lot Number'] = safe_col(df_selected, 'LOT')
            df_e1['Requested Date'] = [''] * len(df_selected)
            df_e1['Location'] = safe_col(df_selected, 'Location')

            st.success(f"✅ 선택한 **{len(df_e1):,}개 행**이 정상 변환되었습니다. 마우스로 드래그 후 `Ctrl+C` 하세요.")

            st.dataframe(
                df_e1,
                use_container_width=True,
                hide_index=True
            )

            tsv_data = df_e1.to_csv(sep='\t', index=False, header=False).encode('utf-8-sig')
            st.download_button(
                label="📥 선택 건 E1 복붙용 파일 다운로드 (.tsv)",
                data=tsv_data,
                file_name=f"E1_Upload_{selected_cust}.tsv",
                mime="text/tab-separated-values"
            )
