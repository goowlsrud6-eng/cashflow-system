import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# -----------------------------------------------------------
# 1. 페이지 설정
# -----------------------------------------------------------
st.set_page_config(page_title="자금 관리 대시보드 Pro", layout="wide")

def fmt_num(x):
    return f"{x:,.2f}"

def fmt_krw(x):
    return f"{x:,.0f}"

# -----------------------------------------------------------
# 2. 데이터 로딩 및 전처리
# -----------------------------------------------------------
def clean_currency(x):
    if isinstance(x, str):
        clean_str = re.sub(r'[^\d.-]', '', x)
        try: return float(clean_str) if clean_str else 0.0
        except: return 0.0
    return float(x) if pd.notnull(x) else 0.0

@st.cache_data(ttl=60)
def load_all_data(file):
    try:
        xls = pd.ExcelFile(file)
        
        # 1. 다이렉트
        if '다이렉트' in xls.sheet_names:
            df_d = pd.read_excel(xls, sheet_name='다이렉트')
            df_d.columns = df_d.columns.str.strip()
            if '잔금_금액' in df_d.columns:
                df_d['잔금_금액'] = df_d['잔금_금액'].apply(clean_currency).fillna(0)
            if '잔금_날짜' in df_d.columns:
                df_d['잔금_날짜'] = pd.to_datetime(df_d['잔금_날짜'], errors='coerce')
            if '화폐단위' in df_d.columns:
                df_d['화폐단위'] = df_d['화폐단위'].astype(str).str.upper().str.strip()
            if '구분' not in df_d.columns:
                df_d['구분'] = 'Direct'
        else:
            df_d = pd.DataFrame()

        # 2. YIWU (수수료 1.1배 로직)
        if 'YIWU' in xls.sheet_names:
            df_y = pd.read_excel(xls, sheet_name='YIWU')
            df_y.columns = df_y.columns.str.strip()
            
            if '잔금' in df_y.columns and '잔금_금액' not in df_y.columns:
                df_y.rename(columns={'잔금': '잔금_금액'}, inplace=True)
            
            if '잔금_금액' in df_y.columns:
                df_y['잔금_금액'] = df_y['잔금_금액'].apply(clean_currency).fillna(0)
                
                comm_col = next((c for c in df_y.columns if '수수료' in str(c)), None)
                if comm_col:
                    def apply_fee(row):
                        val = row['잔금_금액']
                        status = str(row[comm_col]).strip()
                        if '별도' in status:
                            return val * 1.1
                        return val
                    df_y['잔금_금액'] = df_y.apply(apply_fee, axis=1)

            if '잔금_날짜' in df_y.columns:
                df_y['잔금_날짜'] = pd.to_datetime(df_y['잔금_날짜'], errors='coerce')
        else:
            df_y = pd.DataFrame()

        # 3. 송금내역
        yiwu_balance = 0.0
        df_l = pd.DataFrame()
        target_sheet = '송금내역 (YIWU)'
        
        if target_sheet not in xls.sheet_names:
             for sheet in xls.sheet_names:
                 if '송금' in sheet and 'YIWU' in sheet:
                     target_sheet = sheet
                     break
        
        if target_sheet in xls.sheet_names:
            df_l = pd.read_excel(xls, sheet_name=target_sheet)
            col_str = str(list(df_l.columns))
            if '잔고' not in col_str:
                 df_l = pd.read_excel(xls, sheet_name=target_sheet, header=1)
            
            df_l.columns = df_l.columns.str.strip()
            bal_col = None
            for c in df_l.columns:
                if '잔고' in str(c) and 'CNY' in str(c):
                    bal_col = c
                    break
            
            if bal_col:
                balances = df_l[bal_col].apply(clean_currency)
                if not balances.dropna().empty:
                    yiwu_balance = balances.dropna().iloc[-1]
            
            if '날짜' in df_l.columns:
                df_l['날짜'] = pd.to_datetime(df_l['날짜'], errors='coerce')
                
        return df_d, df_y, yiwu_balance, df_l

    except Exception as e:
        st.error(f"데이터 로드 에러: {e}")
        return pd.DataFrame(), pd.DataFrame(), 0.0, pd.DataFrame()

def get_date_range(today):
    start_week = today - timedelta(days=today.weekday())
    end_week = start_week + timedelta(days=6)
    start_next_week = end_week + timedelta(days=1)
    end_next_week = start_next_week + timedelta(days=6)
    
    start_month = today.replace(day=1)
    if today.month == 12:
        start_next_month = today.replace(year=today.year+1, month=1, day=1)
    else:
        start_next_month = today.replace(month=today.month+1, day=1)
    end_month = start_next_month - timedelta(days=1)
    
    if start_next_month.month == 12:
        start_month_after_next = start_next_month.replace(year=start_next_month.year+1, month=1, day=1)
    else:
        start_month_after_next = start_next_month.replace(month=start_next_month.month+1, day=1)
    end_next_month = start_month_after_next - timedelta(days=1)
    
    return {
        "this_week": (start_week, end_week),
        "next_week": (start_next_week, end_next_week),
        "this_month": (start_month, end_month),
        "next_month": (start_next_month, end_next_month),
        "this_plus_next_month": (start_month, end_next_month)
    }

# -----------------------------------------------------------
# 3. 사이드바
# -----------------------------------------------------------
with st.sidebar:
    st.title("⚙️ 자금 설정")
    st.success("✅ Ver.36 적용완료 (표 옆 잔고표시)")
    
    menu = st.radio("화면 이동", 
        ["📊 전체 자금 현황", "다이렉트 (CNY)", "다이렉트 (USD)", "이우 (YIWU)"])
    
    st.markdown("---")
    uploaded_file = st.file_uploader("📂 엑셀 파일 업로드", type=['xlsx'])
    
    st.markdown("---")
    col_r1, col_r2 = st.columns(2)
    with col_r1: rate_cny = st.number_input("1 CNY (원)", value=195.0, format="%.2f")
    with col_r2: rate_usd = st.number_input("1 USD (원)", value=1400.0, format="%.2f")
    cny_to_usd_rate = rate_cny / rate_usd if rate_usd > 0 else 0

    st.markdown("---")
    st.subheader("💼 내 통장 보유액 (입력)")
    my_cny = st.number_input("CNY 보유액", value=0.0, step=100.0)
    my_usd = st.number_input("USD 보유액", value=0.0, step=100.0)
    
    st.markdown("---")
    today = pd.Timestamp.now().normalize()
    custom_date = st.date_input("📅 사용자 지정 기간", (today, today + timedelta(days=14)))

# -----------------------------------------------------------
# 4. 화면 로직
# -----------------------------------------------------------
if uploaded_file:
    df_d, df_y, yiwu_balance, df_l = load_all_data(uploaded_file)
    
    # 필터링
    if '진행단계' in df_d.columns:
        df_d_active = df_d[~df_d['진행단계'].astype(str).str.contains('완료')].copy()
    else: df_d_active = df_d.copy()
    
    if '진행단계' in df_y.columns:
        df_y_active = df_y[~df_y['진행단계'].astype(str).str.contains('완료')].copy()
    else: df_y_active = df_y.copy()

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

    def split_direct_data(df):
        mask_usd = (
            df['구분'].astype(str).str.contains('USD|결제', case=False) | 
            (df['화폐단위'] == 'USD')
        )
        return df[~mask_usd].copy(), df[mask_usd].copy()

    # =======================================================
    # PAGE 1: 전체 자금 현황 (3단 분리 대시보드)
    # =======================================================
    if menu == "📊 전체 자금 현황":
        st.header("📊 전체 자금 현황 대시보드")
        
        # 상단 큰 글씨도 유지 (원하시면 삭제 가능)
        c1, c2, c3 = st.columns(3)
        c1.metric("내 CNY 잔고", fmt_num(my_cny))
        c2.metric("내 USD 잔고", fmt_num(my_usd))
        c3.metric("이우 장부 잔고", fmt_num(yiwu_balance))
        
        st.markdown("---")

        # 1. 다이렉트 CNY 표
        # 제목 옆에 잔고 표시 (columns 사용)
        c_h, c_b = st.columns([5, 2])
        with c_h: st.subheader("1️⃣ 다이렉트 (CNY) 현황")
        with c_b: st.markdown(f"**💰 내 CNY 잔고:** :green[{fmt_num(my_cny)}]")
        
        df_cny_only, _ = split_direct_data(df_d_active)
        rows_cny = []
        for label, s, e in periods:
            if s and e: sub = df_cny_only[(df_cny_only['잔금_날짜'] >= s) & (df_cny_only['잔금_날짜'] <= e)]
            else: sub = df_cny_only
            
            expense = sub['잔금_금액'].sum()
            shortage = max(expense - my_cny, 0)
            rows_cny.append({
                "기간": label,
                "지출 예정(CNY)": fmt_num(expense),
                "내 잔고": fmt_num(my_cny),
                "🛑 송금 필요(CNY)": fmt_num(shortage),
                "🇰🇷 필요 원화": fmt_krw(shortage * rate_cny)
            })
        st.dataframe(pd.DataFrame(rows_cny), hide_index=True, use_container_width=True)

        st.markdown("---")

        # 2. 다이렉트 USD 표
        c_h, c_b = st.columns([5, 2])
        with c_h: st.subheader("2️⃣ 다이렉트 (USD) 현황")
        with c_b: st.markdown(f"**💰 내 USD 잔고:** :green[{fmt_num(my_usd)}]")

        _, df_usd_only = split_direct_data(df_d_active)
        rows_usd = []
        for label, s, e in periods:
            if s and e: sub = df_usd_only[(df_usd_only['잔금_날짜'] >= s) & (df_usd_only['잔금_날짜'] <= e)]
            else: sub = df_usd_only
            
            val_pure = sub[sub['화폐단위'] == 'USD']['잔금_금액'].sum()
            val_conv = sub[sub['화폐단위'] == 'CNY']['잔금_금액'].sum() * cny_to_usd_rate
            expense = val_pure + val_conv
            shortage = max(expense - my_usd, 0)
            
            rows_usd.append({
                "기간": label,
                "지출 예정(USD)": fmt_num(expense),
                "내 잔고": fmt_num(my_usd),
                "🛑 송금 필요(USD)": fmt_num(shortage),
                "🇰🇷 필요 원화": fmt_krw(shortage * rate_usd)
            })
        st.dataframe(pd.DataFrame(rows_usd), hide_index=True, use_container_width=True)

        st.markdown("---")

        # 3. 이우 표
        c_h, c_b1, c_b2 = st.columns([4, 2, 2])
        with c_h: st.subheader("3️⃣ 이우 (YIWU) 현황")
        with c_b1: st.markdown(f"**📒 장부 잔고:** :blue[{fmt_num(yiwu_balance)}]")
        with c_b2: st.markdown(f"**💰 내 USD 잔고:** :green[{fmt_num(my_usd)}]")
        
        rows_yiwu = []
        for label, s, e in periods:
            if s and e: sub = df_y_active[(df_y_active['잔금_날짜'] >= s) & (df_y_active['잔금_날짜'] <= e)]
            else: sub = df_y_active
            
            expense_cny = sub['잔금_금액'].sum()
            needed_cny = max(expense_cny - yiwu_balance, 0)
            needed_usd = needed_cny * cny_to_usd_rate
            final_shortage = max(needed_usd - my_usd, 0)
            
            rows_yiwu.append({
                "기간": label,
                "지출 예정(CNY)": fmt_num(expense_cny),
                "장부 부족분(CNY)": fmt_num(needed_cny),
                "환산(USD)": fmt_num(needed_usd),
                "🛑 최종 송금필요(USD)": fmt_num(final_shortage),
                "🇰🇷 필요 원화": fmt_krw(final_shortage * rate_usd)
            })
        st.dataframe(pd.DataFrame(rows_yiwu), hide_index=True, use_container_width=True)

    # =======================================================
    # PAGE 2, 3, 4 (개별 메뉴)
    # =======================================================
    elif menu == "다이렉트 (CNY)":
        st.header("다이렉트 관리 (CNY 건만)")
        st.metric("내 CNY 잔고 (차감 기준)", fmt_num(my_cny))
        
        df_cny_only, _ = split_direct_data(df_d_active)
        df_view = df_cny_only.copy()
        
        rows = []
        for label, s, e in periods:
            if s and e: sub = df_view[(df_view['잔금_날짜'] >= s) & (df_view['잔금_날짜'] <= e)]
            else: sub = df_view
            expense = sub['잔금_금액'].sum()
            shortage = max(expense - my_cny, 0)
            rows.append({
                "기간": label, "지출 예정(CNY)": fmt_num(expense), "내 잔고": fmt_num(my_cny),
                "🛑 송금 필요(CNY)": fmt_num(shortage), "🇰🇷 필요 원화": fmt_krw(shortage * rate_cny)
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        st.markdown("---")
        st.subheader("📋 상세 내역")
        st.dataframe(df_view[['잔금_날짜','품목','거래처','잔금_금액','진행단계']].sort_values('잔금_날짜'), use_container_width=True)

    elif menu == "다이렉트 (USD)":
        st.header("다이렉트 관리 (USD 건만)")
        st.metric("내 USD 잔고 (차감 기준)", fmt_num(my_usd))
        
        _, df_usd_only = split_direct_data(df_d_active)
        df_view = df_usd_only.copy()
        
        rows = []
        for label, s, e in periods:
            if s and e: sub = df_view[(df_view['잔금_날짜'] >= s) & (df_view['잔금_날짜'] <= e)]
            else: sub = df_view
            val_pure = sub[sub['화폐단위'] == 'USD']['잔금_금액'].sum()
            val_conv = sub[sub['화폐단위'] == 'CNY']['잔금_금액'].sum() * cny_to_usd_rate
            expense = val_pure + val_conv
            shortage = max(expense - my_usd, 0)
            rows.append({
                "기간": label, "지출 예정(USD)": fmt_num(expense), "내 잔고": fmt_num(my_usd),
                "🛑 송금 필요(USD)": fmt_num(shortage), "🇰🇷 필요 원화": fmt_krw(shortage * rate_usd)
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        st.markdown("---")
        st.subheader("📋 상세 내역")
        st.dataframe(df_view[['잔금_날짜','품목','거래처','잔금_금액','화폐단위','진행단계']].sort_values('잔금_날짜'), use_container_width=True)

    elif menu == "이우 (YIWU)":
        st.header("이우(YIWU) 자금 관리")
        c1, c2, c3 = st.columns(3)
        c1.metric("이우 장부 잔고 (CNY)", fmt_num(yiwu_balance))
        c2.metric("내 USD 잔고 (송금용)", fmt_num(my_usd))
        
        rows = []
        for label, s, e in periods:
            if s and e: sub = df_y_active[(df_y_active['잔금_날짜'] >= s) & (df_y_active['잔금_날짜'] <= e)]
            else: sub = df_y_active
            expense_cny = sub['잔금_금액'].sum()
            needed_cny = max(expense_cny - yiwu_balance, 0)
            needed_usd = needed_cny * cny_to_usd_rate
            final_shortage = max(needed_usd - my_usd, 0)
            rows.append({
                "기간": label, "지출 예정(CNY)": fmt_num(expense_cny), "부족분(CNY)": fmt_num(needed_cny),
                "환산(USD)": fmt_num(needed_usd), "🛑 최종 송금필요(USD)": fmt_num(final_shortage)
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        st.markdown("---")
        st.subheader("📋 상세 내역 (수수료 1.1배 적용됨)")
        df_disp = df_y_active.sort_values('잔금_날짜').copy()
        cols = ['잔금_날짜', '품목', '총_발주금액', '잔금_금액', '진행단계']
        comm_col = next((c for c in df_disp.columns if '수수료' in str(c)), None)
        if comm_col: cols.insert(2, comm_col)
        valid = [c for c in cols if c in df_disp.columns]
        disp = df_disp[valid].copy()
        for c in ['총_발주금액', '잔금_금액']: 
            if c in disp.columns: disp[c] = disp[c].apply(fmt_num)
        if '잔금_날짜' in disp.columns: disp['잔금_날짜'] = disp['잔금_날짜'].dt.strftime('%Y-%m-%d')
        st.dataframe(disp, hide_index=True, use_container_width=True)

else:
    st.info("👈 엑셀 파일을 업로드해주세요.")