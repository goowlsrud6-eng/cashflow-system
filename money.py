import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re
import requests
import io

# -----------------------------------------------------------
# 1. 페이지 설정
# -----------------------------------------------------------
st.set_page_config(page_title="자금 관리 대시보드 Pro", layout="wide")

def fmt_num(x):
    return f"{x:,.2f}"

def fmt_krw(x):
    return f"{x:,.0f}"

# -----------------------------------------------------------
# 2. 데이터 로딩 및 전처리 (초고속 캐싱 최적화)
# -----------------------------------------------------------
def clean_currency(x):
    if isinstance(x, str):
        clean_str = re.sub(r'[^\d.-]', '', x)
        try: return float(clean_str) if clean_str else 0.0
        except: return 0.0
    return float(x) if pd.notnull(x) else 0.0

# 엑셀 파싱 핵심 로직 (캐싱 안함, 데이터만 처리)
def parse_excel_data(file_bytes):
    try:
        xls = pd.ExcelFile(file_bytes)
        
        # 1. 다이렉트
        if '다이렉트' in xls.sheet_names:
            df_d = pd.read_excel(xls, sheet_name='다이렉트')
            df_d.columns = df_d.columns.str.strip()
            if '잔금_금액' in df_d.columns: df_d['잔금_금액'] = df_d['잔금_금액'].apply(clean_currency).fillna(0)
            if '잔금_날짜' in df_d.columns: df_d['잔금_날짜'] = pd.to_datetime(df_d['잔금_날짜'], errors='coerce')
            if '실지급_날짜' in df_d.columns: df_d['실지급_날짜'] = pd.to_datetime(df_d['실지급_날짜'], errors='coerce')
            if '화폐단위' in df_d.columns: df_d['화폐단위'] = df_d['화폐단위'].astype(str).str.upper().str.strip()
            if '구분' not in df_d.columns: df_d['구분'] = 'Direct'
        else:
            df_d = pd.DataFrame()

        # 2. YIWU
        if 'YIWU' in xls.sheet_names:
            df_y = pd.read_excel(xls, sheet_name='YIWU')
            df_y.columns = df_y.columns.str.strip()
            if '잔금' in df_y.columns and '잔금_금액' not in df_y.columns: df_y.rename(columns={'잔금': '잔금_금액'}, inplace=True)
            if '잔금_금액' in df_y.columns:
                df_y['잔금_금액'] = df_y['잔금_금액'].apply(clean_currency).fillna(0)
                comm_col = next((c for c in df_y.columns if '수수료' in str(c)), None)
                if comm_col:
                    def apply_fee(row):
                        val = row['잔금_금액']
                        status = str(row[comm_col]).strip()
                        if '별도' in status: return val * 1.1
                        return val
                    df_y['잔금_금액'] = df_y.apply(apply_fee, axis=1)
            if '잔금_날짜' in df_y.columns: df_y['잔금_날짜'] = pd.to_datetime(df_y['잔금_날짜'], errors='coerce')
        else:
            df_y = pd.DataFrame()

        # 3. 송금내역 (YIWU)
        yiwu_balance = 0.0
        df_l = pd.DataFrame()
        target_sheet = '송금내역 (YIWU)'
        if target_sheet not in xls.sheet_names:
             target_sheet = next((s for s in xls.sheet_names if '송금' in s and 'YIWU' in s), target_sheet)
        if target_sheet in xls.sheet_names:
            df_l = pd.read_excel(xls, sheet_name=target_sheet)
            if '잔고' not in str(list(df_l.columns)): df_l = pd.read_excel(xls, sheet_name=target_sheet, header=1)
            df_l.columns = df_l.columns.str.strip()
            bal_col = next((c for c in df_l.columns if '잔고' in str(c) and 'CNY' in str(c)), None)
            if bal_col:
                balances = df_l[bal_col].apply(clean_currency)
                if not balances.dropna().empty: yiwu_balance = balances.dropna().iloc[-1]
            if '날짜' in df_l.columns: df_l['날짜'] = pd.to_datetime(df_l['날짜'], errors='coerce')
                
        # 4. 환전내역
        df_ex = pd.DataFrame()
        if '환전내역' in xls.sheet_names:
            df_ex = pd.read_excel(xls, sheet_name='환전내역')
            df_ex.columns = df_ex.columns.str.strip()
            date_col = next((c for c in df_ex.columns if '날짜' in str(c) or '일자' in str(c)), None)
            if date_col: df_ex['날짜'] = pd.to_datetime(df_ex[date_col], errors='coerce')
            curr_col = next((c for c in df_ex.columns if '화폐' in str(c) or '통화' in str(c) or '구분' in str(c)), None)
            if curr_col: df_ex.rename(columns={curr_col: '화폐'}, inplace=True)
            amt_col = next((c for c in df_ex.columns if '환전' in str(c) and '원화' not in str(c)), None)
            if not amt_col: amt_col = next((c for c in df_ex.columns if '외화' in str(c)), None)
            if amt_col: df_ex['환전금액'] = df_ex[amt_col].apply(clean_currency).fillna(0)
            
        return df_d, df_y, yiwu_balance, df_l, df_ex
    except Exception as e:
        st.error(f"데이터 로드 에러: {e}")
        return pd.DataFrame(), pd.DataFrame(), 0.0, pd.DataFrame(), pd.DataFrame()

# 다운로드 및 파싱을 하나로 묶어 강력한 캐싱 적용 (10분 유지)
@st.cache_data(ttl=600, show_spinner="☁️ 구글 드라이브에서 엑셀 파일을 불러오고 분석하는 중입니다... (최초 1회만 소요)")
def get_drive_data(url):
    file_id_match = re.search(r'/d/([a-zA-Z0-9_-]+)', url) or re.search(r'id=([a-zA-Z0-9_-]+)', url)
    if not file_id_match: return None
    
    file_id = file_id_match.group(1)
    download_url = f"https://drive.google.com/uc?id={file_id}&export=download"
    try:
        response = requests.get(download_url)
        if response.status_code == 200:
            file_bytes = io.BytesIO(response.content)
            return parse_excel_data(file_bytes)
    except Exception as e:
        st.error(f"구글 드라이브 연동 실패: {e}")
    return None

@st.cache_data(ttl=60)
def get_local_data(file):
    return parse_excel_data(file)

# -----------------------------------------------------------
# 3. 잔고 자동 계산 로직
# -----------------------------------------------------------
def calculate_realtime_balances(df_d, df_ex, df_l, base_date, base_cny, base_usd):
    cny_bal = base_cny
    usd_bal = base_usd

    if not df_ex.empty and '날짜' in df_ex.columns:
        new_ex = df_ex[df_ex['날짜'] > base_date]
        for _, row in new_ex.iterrows():
            amt = clean_currency(row.get('환전금액', 0))
            curr = str(row.get('화폐', '')).upper()
            if 'CNY' in curr: cny_bal += amt
            elif 'USD' in curr: usd_bal += amt

    if not df_d.empty and '실지급_날짜' in df_d.columns:
        df_d_paid = df_d[(df_d['실지급_날짜'] > base_date) & (df_d['진행단계'].astype(str).str.contains('완료', na=False))]
        for _, row in df_d_paid.iterrows():
            gubun = str(row.get('구분', ''))
            curr = str(row.get('화폐단위', '')).upper()
            amt_paid = clean_currency(row.get('실지급_금액', 0))
            amt_usd_actual = clean_currency(row.get('실제출금(USD)', 0))

            if 'USD' in gubun and 'CNY' in curr and amt_usd_actual > 0:
                usd_bal -= amt_usd_actual
            else:
                if 'USD' in curr: usd_bal -= amt_paid
                elif 'CNY' in curr: cny_bal -= amt_paid

    if not df_l.empty and '날짜' in df_l.columns:
        df_l_paid = df_l[(df_l['날짜'] > base_date) & (df_l['구분'].astype(str).str.contains('송금', na=False))]
        for _, row in df_l_paid.iterrows():
            amt_usd = clean_currency(row.get('입금액(USD)', 0))
            usd_bal -= amt_usd

    return cny_bal, usd_bal

def get_date_range(today):
    start_week = today - timedelta(days=today.weekday())
    end_next_month = (today.replace(day=1) + timedelta(days=65)).replace(day=1) - timedelta(days=1)
    return {
        "this_week": (start_week, start_week + timedelta(days=6)),
        "next_week": (start_week + timedelta(days=7), start_week + timedelta(days=13)),
        "this_month": (today.replace(day=1), (today.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)),
        "next_month": ((today.replace(day=1) + timedelta(days=32)).replace(day=1), end_next_month),
        "this_plus_next_month": (today.replace(day=1), end_next_month)
    }

def split_direct_data(df):
    mask_usd = (df['구분'].astype(str).str.contains('USD|결제', case=False) | (df['화폐단위'] == 'USD'))
    return df[~mask_usd].copy(), df[mask_usd].copy()

# -----------------------------------------------------------
# 4. 사이드바 
# -----------------------------------------------------------
with st.sidebar:
    st.title("⚙️ 자금 설정")
    menu = st.radio("화면 이동", ["전체 자금 현황", "다이렉트 (CNY)", "다이렉트 (USD)", "이우 (YIWU)", "환전 내역"])
    st.markdown("---")
    
    st.subheader("🔗 엑셀 연동 (택 1)")
    gdrive_url = st.text_input("구글 드라이브 공유 링크 붙여넣기", placeholder="https://drive.google.com/...")
    
    # ⚡ 최적화: 수동 새로고침 버튼
    if st.button("🔄 최신 데이터 불러오기"):
        st.cache_data.clear()
        st.success("데이터 캐시를 초기화했습니다. 최신 엑셀을 불러옵니다!")
        
    uploaded_file = st.file_uploader("또는 수동 업로드", type=['xlsx'])
    
    st.markdown("---")
    col_r1, col_r2 = st.columns(2)
    with col_r1: rate_cny = st.number_input("1 CNY (원)", value=195.0, format="%.2f")
    with col_r2: rate_usd = st.number_input("1 USD (원)", value=1400.0, format="%.2f")
    cny_to_usd_rate = rate_cny / rate_usd if rate_usd > 0 else 0

    st.markdown("---")
    with st.expander("🛠️ 초기 잔고 기준점 세팅 (2/12 기준)"):
        base_date = st.date_input("기준 날짜", value=pd.to_datetime("2026-02-12"))
        base_cny = st.number_input("초기 CNY", value=436013.34)
        base_usd = st.number_input("초기 USD", value=62785.86)
        
    today = pd.Timestamp.now().normalize()

# -----------------------------------------------------------
# 5. 화면 로직
# -----------------------------------------------------------
data_tuple = None

if gdrive_url:
    data_tuple = get_drive_data(gdrive_url)
    if not data_tuple: st.warning("구글 드라이브 링크를 확인할 수 없거나 로딩에 실패했습니다.")
elif uploaded_file:
    data_tuple = get_local_data(uploaded_file)

if data_tuple:
    df_d, df_y, yiwu_balance, df_l, df_ex = data_tuple
    
    base_date_ts = pd.to_datetime(base_date)
    my_cny, my_usd = calculate_realtime_balances(df_d, df_ex, df_l, base_date_ts, base_cny, base_usd)
    
    with st.sidebar:
        st.subheader("💼 실시간 환전보유액 (자동계산)")
        st.metric("CNY 보유액", fmt_num(my_cny))
        st.metric("USD 보유액", fmt_num(my_usd))
    
    df_d_active = df_d[~df_d['진행단계'].astype(str).str.contains('완료', na=False)].copy() if '진행단계' in df_d.columns else df_d.copy()
    df_y_active = df_y[~df_y['진행단계'].astype(str).str.contains('완료', na=False)].copy() if '진행단계' in df_y.columns else df_y.copy()

    dates = get_date_range(today)
    periods = [
        ("0. 전체 예정", None, None),
        ("1. 이번주", dates['this_week'][0], dates['this_week'][1]),
        ("2. 다음주", dates['next_week'][0], dates['next_week'][1]),
        ("3. 이번주+다음주", dates['this_week'][0], dates['next_week'][1]),
        ("4. 이번달", dates['this_month'][0], dates['this_month'][1]),
        ("5. 다음달", dates['next_month'][0], dates['next_month'][1]),
        ("6. 이번달+다음달", dates['this_plus_next_month'][0], dates['this_plus_next_month'][1]),
    ]

    # =======================================================
    # PAGE 1: 전체 자금 현황
    # =======================================================
    if menu == "전체 자금 현황":
        st.header("📊 전체 자금 현황 대시보드")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("CNY 환전보유액 (자동)", fmt_num(my_cny))
        c2.metric("USD 환전보유액 (자동)", fmt_num(my_usd))
        c3.metric("허사장님 물품대", fmt_num(yiwu_balance))
        
        st.markdown("---")

        c_h, c_b = st.columns([5, 2])
        with c_h: st.subheader("1️⃣ 다이렉트 (CNY) 현황")
        with c_b: st.markdown(f"**💰 CNY 보유액:** :green[{fmt_num(my_cny)}]")
        
        df_cny_only, _ = split_direct_data(df_d_active)
        rows_cny = []
        for label, s, e in periods:
            sub = df_cny_only[(df_cny_only['잔금_날짜'] >= s) & (df_cny_only['잔금_날짜'] <= e)] if s and e else df_cny_only
            exp_cny = sub['잔금_금액'].sum()
            need_cny = max(exp_cny - my_cny, 0)
            rows_cny.append({
                "기간": label, "지출예정액(CNY)": fmt_num(exp_cny), "지출예정액(KRW)": fmt_krw(exp_cny * rate_cny),
                "송금필요액(CNY)": fmt_num(need_cny), "송금필요액(KRW)": fmt_krw(need_cny * rate_cny)
            })
        st.dataframe(pd.DataFrame(rows_cny), hide_index=True, use_container_width=True)

        st.markdown("---")

        c_h, c_b = st.columns([5, 2])
        with c_h: st.subheader("2️⃣ 다이렉트 (USD) 현황")
        with c_b: st.markdown(f"**💰 USD 보유액:** :green[{fmt_num(my_usd)}]")

        _, df_usd_only = split_direct_data(df_d_active)
        rows_usd = []
        for label, s, e in periods:
            sub = df_usd_only[(df_usd_only['잔금_날짜'] >= s) & (df_usd_only['잔금_날짜'] <= e)] if s and e else df_usd_only
            val_pure = sub[sub['화폐단위'] == 'USD']['잔금_금액'].sum()
            val_conv = sub[sub['화폐단위'] == 'CNY']['잔금_금액'].sum() * cny_to_usd_rate
            exp_usd = val_pure + val_conv
            need_usd = max(exp_usd - my_usd, 0)
            rows_usd.append({
                "기간": label, "지출예정액(USD)": fmt_num(exp_usd), "지출예정액(KRW)": fmt_krw(exp_usd * rate_usd),
                "송금필요액(USD)": fmt_num(need_usd), "송금필요액(KRW)": fmt_krw(need_usd * rate_usd)
            })
        st.dataframe(pd.DataFrame(rows_usd), hide_index=True, use_container_width=True)

        st.markdown("---")

        c_h, c_b1, c_b2 = st.columns([4, 2, 2])
        with c_h: st.subheader("3️⃣ 이우 (YIWU) 현황")
        with c_b1: st.markdown(f"**📒 허사장님 물품대:** :blue[{fmt_num(yiwu_balance)}]")
        with c_b2: st.markdown(f"**💰 USD 보유액:** :green[{fmt_num(my_usd)}]")
        
        rows_yiwu = []
        for label, s, e in periods:
            sub = df_y_active[(df_y_active['잔금_날짜'] >= s) & (df_y_active['잔금_날짜'] <= e)] if s and e else df_y_active
            exp_cny = sub['잔금_금액'].sum()
            short_cny = max(exp_cny - yiwu_balance, 0)
            short_usd = short_cny * cny_to_usd_rate
            remit_usd = max(short_usd - my_usd, 0)
            rows_yiwu.append({
                "기간": label, "지출예정액(CNY)": fmt_num(exp_cny), "지출예정액(KRW)": fmt_krw(exp_cny * rate_cny),
                "물품대 부족액(CNY)": fmt_num(short_cny), "물품대 부족액(USD)": fmt_num(short_usd),
                "물품대 부족액(KRW)": fmt_krw(short_cny * rate_cny),
                "송금필요액(USD)": fmt_num(remit_usd), "송금필요액(KRW)": fmt_krw(remit_usd * rate_usd)
            })
        st.dataframe(pd.DataFrame(rows_yiwu), hide_index=True, use_container_width=True)

    # =======================================================
    # PAGE: 환전 내역
    # =======================================================
    elif menu == "환전 내역":
        st.header("💱 환전 내역 관리")
        if df_ex.empty: st.warning("데이터가 없습니다.")
        else:
            df_cny_ex = df_ex[df_ex['화폐'].astype(str).str.upper() == 'CNY'].copy()
            df_usd_ex = df_ex[df_ex['화폐'].astype(str).str.upper() == 'USD'].copy()
            tab1, tab2 = st.tabs(["🇨🇳 CNY 환전", "🇺🇸 USD 환전"])
            
            with tab1:
                st.subheader("🇨🇳 CNY 환전 내역")
                st.dataframe(df_cny_ex, hide_index=True, use_container_width=True)
            with tab2:
                st.subheader("🇺🇸 USD 환전 내역")
                st.dataframe(df_usd_ex, hide_index=True, use_container_width=True)

    # =======================================================
    # PAGE: 다이렉트 CNY
    # =======================================================
    elif menu == "다이렉트 (CNY)":
        st.header("다이렉트 관리 (CNY)")
        c1, c2 = st.columns(2)
        c1.metric("CNY 보유액 (자동)", fmt_num(my_cny), f"≈ {fmt_krw(my_cny * rate_cny)} 원")
        
        df_cny_only, _ = split_direct_data(df_d_active)
        df_view = df_cny_only.copy()
        
        rows = []
        for label, s, e in periods:
            sub = df_view[(df_view['잔금_날짜'] >= s) & (df_view['잔금_날짜'] <= e)] if s and e else df_view
            exp_cny = sub['잔금_금액'].sum()
            need_cny = max(exp_cny - my_cny, 0)
            rows.append({
                "기간": label, "지출예정액(CNY)": fmt_num(exp_cny), "지출예정액(KRW)": fmt_krw(exp_cny * rate_cny),
                "송금필요액(CNY)": fmt_num(need_cny), "송금필요액(KRW)": fmt_krw(need_cny * rate_cny)
            })
        st.subheader("📅 기간별 CNY 자금 계획")
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📋 상세 내역")
        df_view['잔금 금액(KRW)'] = df_view['잔금_금액'] * rate_cny
        df_disp = df_view[['잔금_날짜', '품목', '거래처', '잔금_금액', '잔금 금액(KRW)', '진행단계']].copy()
        df_disp.columns = ['잔금 날짜', '상품명', '거래처', '잔금 금액(CNY)', '잔금 금액(KRW)', '진행단계']
        if '잔금 날짜' in df_disp.columns: df_disp['잔금 날짜'] = df_disp['잔금 날짜'].dt.strftime('%Y-%m-%d')
        if '잔금 금액(CNY)' in df_disp.columns: df_disp['잔금 금액(CNY)'] = df_disp['잔금 금액(CNY)'].apply(fmt_num)
        if '잔금 금액(KRW)' in df_disp.columns: df_disp['잔금 금액(KRW)'] = df_disp['잔금 금액(KRW)'].apply(fmt_krw)
        st.dataframe(df_disp.sort_values('잔금 날짜'), hide_index=True, use_container_width=True)

    # =======================================================
    # PAGE: 다이렉트 USD
    # =======================================================
    elif menu == "다이렉트 (USD)":
        st.header("다이렉트 관리 (USD)")
        c1, c2 = st.columns(2)
        c1.metric("USD 보유액 (자동)", fmt_num(my_usd), f"≈ {fmt_krw(my_usd * rate_usd)} 원")
        
        _, df_usd_only = split_direct_data(df_d_active)
        df_view = df_usd_only.copy()
        
        rows = []
        for label, s, e in periods:
            sub = df_view[(df_view['잔금_날짜'] >= s) & (df_view['잔금_날짜'] <= e)] if s and e else df_view
            val_pure = sub[sub['화폐단위'] == 'USD']['잔금_금액'].sum()
            val_conv = sub[sub['화폐단위'] == 'CNY']['잔금_금액'].sum() * cny_to_usd_rate
            exp_usd = val_pure + val_conv
            need_usd = max(exp_usd - my_usd, 0)
            rows.append({
                "기간": label, "지출예정액(USD)": fmt_num(exp_usd), "지출예정액(KRW)": fmt_krw(exp_usd * rate_usd),
                "송금필요액(USD)": fmt_num(need_usd), "송금필요액(KRW)": fmt_krw(need_usd * rate_usd)
            })
        st.subheader("📅 기간별 USD 자금 계획")
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📋 상세 내역")
        df_view['잔금 금액(CNY)'] = df_view.apply(lambda r: r['잔금_금액'] if r['화폐단위'] == 'CNY' else 0, axis=1)
        df_view['잔금 금액(USD)'] = df_view.apply(lambda r: r['잔금_금액'] if r['화폐단위'] == 'USD' else r['잔금_금액'] * cny_to_usd_rate, axis=1)
        df_view['잔금 금액(KRW)'] = df_view['잔금 금액(USD)'] * rate_usd
        
        df_disp = df_view[['잔금_날짜', '품목', '거래처', '잔금 금액(CNY)', '잔금 금액(USD)', '잔금 금액(KRW)', '진행단계']].copy()
        df_disp.columns = ['잔금 날짜', '상품명', '거래처', '잔금 금액(CNY)', '잔금 금액(USD)', '잔금 금액(KRW)', '진행단계']
        if '잔금 날짜' in df_disp.columns: df_disp['잔금 날짜'] = df_disp['잔금 날짜'].dt.strftime('%Y-%m-%d')
        if '잔금 금액(CNY)' in df_disp.columns: df_disp['잔금 금액(CNY)'] = df_disp['잔금 금액(CNY)'].apply(lambda x: fmt_num(x) if x > 0 else "")
        if '잔금 금액(USD)' in df_disp.columns: df_disp['잔금 금액(USD)'] = df_disp['잔금 금액(USD)'].apply(fmt_num)
        if '잔금 금액(KRW)' in df_disp.columns: df_disp['잔금 금액(KRW)'] = df_disp['잔금 금액(KRW)'].apply(fmt_krw)
        st.dataframe(df_disp.sort_values('잔금 날짜'), hide_index=True, use_container_width=True)

    # =======================================================
    # PAGE: 이우 (YIWU)
    # =======================================================
    elif menu == "이우 (YIWU)":
        st.header("이우(YIWU) 자금 관리")
        c1, c2 = st.columns(2)
        c1.metric("허사장님 물품대", fmt_num(yiwu_balance), f"≈ {fmt_krw(yiwu_balance * rate_cny)} 원")
        c2.metric("USD 보유액 (자동)", fmt_num(my_usd), f"≈ {fmt_krw(my_usd * rate_usd)} 원")
        
        rows = []
        for label, s, e in periods:
            sub = df_y_active[(df_y_active['잔금_날짜'] >= s) & (df_y_active['잔금_날짜'] <= e)] if s and e else df_y_active
            exp_cny = sub['잔금_금액'].sum()
            short_cny = max(exp_cny - yiwu_balance, 0)
            short_usd = short_cny * cny_to_usd_rate
            rows.append({
                "기간": label, "지출예정액(CNY)": fmt_num(exp_cny), "지출예정액(USD)": fmt_num(exp_cny * cny_to_usd_rate),
                "지출예정액(KRW)": fmt_krw(exp_cny * rate_cny), "물품대 부족액(CNY)": fmt_num(short_cny),
                "물품대 부족액(USD)": fmt_num(short_usd), "물품대 부족액(KRW)": fmt_krw(short_cny * rate_cny)
            })
        st.subheader("📅 기간별 이우(YIWU) 자금 계획")
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📋 상세 내역")
        df_disp = df_y_active.sort_values('잔금_날짜').copy()
        df_disp['잔금 금액(CNY)'] = df_disp['잔금_금액']
        df_disp['잔금 금액(USD)'] = df_disp['잔금_금액'] * cny_to_usd_rate
        df_disp['잔금 금액(KRW)'] = df_disp['잔금_금액'] * rate_cny
        
        cols_to_show = ['잔금_날짜', '품목', '잔금 금액(CNY)', '잔금 금액(USD)', '잔금 금액(KRW)', '진행단계']
        disp_final = df_disp[[c for c in cols_to_show if c in df_disp.columns]].copy()
        disp_final.rename(columns={'잔금_날짜': '잔금 날짜', '품목': '상품명'}, inplace=True)
        
        if '잔금 날짜' in disp_final.columns: disp_final['잔금 날짜'] = disp_final['잔금 날짜'].dt.strftime('%Y-%m-%d')
        if '잔금 금액(CNY)' in disp_final.columns: disp_final['잔금 금액(CNY)'] = disp_final['잔금 금액(CNY)'].apply(fmt_num)
        if '잔금 금액(USD)' in disp_final.columns: disp_final['잔금 금액(USD)'] = disp_final['잔금 금액(USD)'].apply(fmt_num)
        if '잔금 금액(KRW)' in disp_final.columns: disp_final['잔금 금액(KRW)'] = disp_final['잔금 금액(KRW)'].apply(fmt_krw)
        st.dataframe(disp_final, hide_index=True, use_container_width=True)

else:
    st.info("👈 왼쪽 사이드바에 구글 드라이브 공유 링크를 붙여넣어주세요! (최초 1회만 하면 끝납니다)")