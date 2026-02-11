import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# -----------------------------------------------------------
# 1. 페이지 및 스타일 설정
# -----------------------------------------------------------
st.set_page_config(page_title="자금 관리 대시보드 Pro", layout="wide")

# 숫자 포맷팅 함수
def fmt_num(x):
    return f"{x:,.2f}"

def fmt_krw(x):
    return f"{x:,.0f}"

# -----------------------------------------------------------
# 2. 데이터 로딩 및 전처리 함수
# -----------------------------------------------------------
def clean_currency(x):
    if isinstance(x, str):
        clean_str = re.sub(r'[^\d.-]', '', x)
        try: return float(clean_str) if clean_str else 0.0
        except: return 0.0
    return float(x) if pd.notnull(x) else 0.0

@st.cache_data(ttl=3600)
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

        # 2. YIWU
        if 'YIWU' in xls.sheet_names:
            df_y = pd.read_excel(xls, sheet_name='YIWU')
            df_y.columns = df_y.columns.str.strip()
            if '잔금' in df_y.columns and '잔금_금액' not in df_y.columns:
                df_y.rename(columns={'잔금': '잔금_금액'}, inplace=True)
            if '잔금_금액' in df_y.columns:
                df_y['잔금_금액'] = df_y['잔금_금액'].apply(clean_currency).fillna(0)
            if '잔금_날짜' in df_y.columns:
                df_y['잔금_날짜'] = pd.to_datetime(df_y['잔금_날짜'], errors='coerce')
        else:
            df_y = pd.DataFrame()

        # 3. 송금내역 (YIWU)
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
                valid_bals = balances[balances.notna()]
                if not valid_bals.empty:
                    yiwu_balance = valid_bals.iloc[-1]
            
            if '날짜' in df_l.columns:
                df_l['날짜'] = pd.to_datetime(df_l['날짜'], errors='coerce')
                
        return df_d, df_y, yiwu_balance, df_l

    except Exception as e:
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
        "next_month": (start_next_month, end_next_month)
    }

# -----------------------------------------------------------
# 3. 사이드바 (설정)
# -----------------------------------------------------------
with st.sidebar:
    st.title("⚙️ 자금 설정")
    menu = st.radio("화면 이동", ["📊 전체 자금 현황", "🟩 다이렉트 관리", "🟦 이우(YIWU) 관리"])
    
    st.markdown("---")
    uploaded_file = st.file_uploader("📂 통합 엑셀 파일", type=['xlsx'])
    
    st.markdown("---")
    st.subheader("💱 환율")
    col_r1, col_r2 = st.columns(2)
    with col_r1: rate_cny = st.number_input("1 CNY (원)", value=195.0, format="%.2f")
    with col_r2: rate_usd = st.number_input("1 USD (원)", value=1400.0, format="%.2f")
    cny_to_usd_rate = rate_cny / rate_usd if rate_usd > 0 else 0
    st.caption(f"1 CNY ≈ {cny_to_usd_rate:.4f} USD")

    st.markdown("---")
    st.subheader("💼 내 통장 보유액 (차감용)")
    my_cny = st.number_input("CNY 보유액", value=0.0, step=100.0, help="전체 현황의 최종 필요액 계산 시 차감됩니다.")
    my_usd = st.number_input("USD 보유액", value=0.0, step=100.0, help="전체 현황의 최종 필요액 계산 시 차감됩니다.")
    
    st.markdown("---")
    today = pd.Timestamp.now().normalize()
    custom_date = st.date_input("📅 기간 지정", (today, today + timedelta(days=14)))

# -----------------------------------------------------------
# 4. 데이터 처리 및 화면 표시
# -----------------------------------------------------------
if uploaded_file:
    df_d, df_y, yiwu_balance, df_l = load_all_data(uploaded_file)
    
    # 필터링
    if '진행단계' in df_d.columns:
        df_d_active = df_d[~df_d['진행단계'].astype(str).str.contains('완료')].copy()
    else:
        df_d_active = df_d.copy()
        
    if '진행단계' in df_y.columns:
        df_y_active = df_y[~df_y['진행단계'].astype(str).str.contains('완료')].copy()
    else:
        df_y_active = df_y.copy()

    dates = get_date_range(today)
    
    periods = [
        ("0. 전체 예정", None, None),
        ("1. 이번주", dates['this_week'][0], dates['this_week'][1]),
        ("2. 다음주", dates['next_week'][0], dates['next_week'][1]),
        ("3. 이번주+다음주", dates['this_week'][0], dates['next_week'][1]),
        ("4. 이번달", dates['this_month'][0], dates['this_month'][1]),
        ("5. 다음달", dates['next_month'][0], dates['next_month'][1]),
    ]

    # =======================================================
    # PAGE 1: 전체 자금 현황
    # =======================================================
    if menu == "📊 전체 자금 현황":
        st.header("📊 기간별 자금 흐름 요약")
        st.info(f"ℹ️ **계산 기준:** (지출 예정액) - (이우 잔고 {fmt_num(yiwu_balance)}) - (내 통장 보유액)")

        def calculate_needs(df_direct_sub, df_yiwu_sub):
            # Direct
            mask_usd = df_direct_sub['구분'].astype(str).str.contains('USD|결제', case=False)
            d_usd_rows = df_direct_sub[mask_usd]
            d_direct_rows = df_direct_sub[~mask_usd]

            # Direct USD (순수 USD + CNY의 USD환산분)
            gross_d_usd = d_usd_rows[d_usd_rows['화폐단위']=='USD']['잔금_금액'].sum() + \
                          (d_usd_rows[d_usd_rows['화폐단위']=='CNY']['잔금_금액'].sum() * cny_to_usd_rate) + \
                          d_direct_rows[d_direct_rows['화폐단위']=='USD']['잔금_금액'].sum()
            
            # Direct CNY (순수 CNY)
            gross_d_cny = d_direct_rows[d_direct_rows['화폐단위']=='CNY']['잔금_금액'].sum()
            
            # YIWU
            gross_y_expense_cny = df_yiwu_sub['잔금_금액'].sum()
            
            return gross_d_cny, gross_d_usd, gross_y_expense_cny

        # 1-1. 상세 기간별 테이블
        result_rows = []
        for label, start, end in periods:
            if start and end:
                sub_d = df_d_active[(df_d_active['잔금_날짜'] >= start) & (df_d_active['잔금_날짜'] <= end)]
                sub_y = df_y_active[(df_y_active['잔금_날짜'] >= start) & (df_y_active['잔금_날짜'] <= end)]
            else:
                sub_d = df_d_active
                sub_y = df_y_active
            
            g_d_cny, g_d_usd, g_y_cny = calculate_needs(sub_d, sub_y)
            
            # 환산
            krw_d_cny = g_d_cny * rate_cny
            krw_d_usd = g_d_usd * rate_usd
            
            usd_y = g_y_cny * cny_to_usd_rate
            krw_y = g_y_cny * rate_cny
            
            total_krw = krw_d_cny + krw_d_usd + krw_y
            
            result_rows.append({
                "기간": label,
                "🟩 다이렉트(CNY)": fmt_num(g_d_cny),
                "🟩 (KRW환산)": fmt_krw(krw_d_cny), # 다이렉트 CNY의 KRW
                "🟩 다이렉트(USD)": fmt_num(g_d_usd),
                "🟩 (KRW환산) ": fmt_krw(krw_d_usd), # 다이렉트 USD의 KRW (공백으로 구분)
                "🟦 이우(CNY)": fmt_num(g_y_cny),
                "🟦 (USD환산)": fmt_num(usd_y),
                "🟦 (KRW환산)": fmt_krw(krw_y),
                "🧾 총 청구액(KRW)": fmt_krw(total_krw)
            })
            
        st.dataframe(pd.DataFrame(result_rows), hide_index=True, use_container_width=True)
        st.markdown("---")

        # 1-2. 지정 기간 & 잔고 차감
        if len(custom_date) == 2:
            start_d, end_d = custom_date
            st.subheader(f"💰 최종 자금 분석 ({start_d} ~ {end_d})")
            
            m_d = (df_d_active['잔금_날짜'] >= pd.Timestamp(start_d)) & (df_d_active['잔금_날짜'] <= pd.Timestamp(end_d))
            m_y = (df_y_active['잔금_날짜'] >= pd.Timestamp(start_d)) & (df_y_active['잔금_날짜'] <= pd.Timestamp(end_d))
            
            bill_d_cny, bill_d_usd, bill_y_cny = calculate_needs(df_d_active[m_d], df_y_active[m_y])
            
            yiwu_shortage_cny = max(bill_y_cny - yiwu_balance, 0)
            yiwu_shortage_usd = yiwu_shortage_cny * cny_to_usd_rate
            
            final_cny_need = max(bill_d_cny - my_cny, 0)
            total_usd_need = bill_d_usd + yiwu_shortage_usd
            final_usd_need = max(total_usd_need - my_usd, 0)
            final_krw = (final_cny_need * rate_cny) + (final_usd_need * rate_usd)
            
            c1, c2, c3 = st.columns(3)
            c1.metric("🟩 CNY 최종 송금", fmt_num(final_cny_need), f"보유 {fmt_num(my_cny)} 차감")
            c2.metric("🇺🇸 USD 최종 송금", fmt_num(final_usd_need), f"보유 {fmt_num(my_usd)} 차감")
            c3.metric("🇰🇷 총 필요 (KRW)", fmt_krw(final_krw))

    # =======================================================
    # PAGE 2: 다이렉트 관리
    # =======================================================
    elif menu == "🟩 다이렉트 관리":
        st.header("🟩 다이렉트 지출 상세")

        # 2-1. 기간별 다이렉트 요약
        st.subheader("📅 기간별 다이렉트 요약")
        d_summary_rows = []
        for label, start, end in periods:
            if start and end:
                sub = df_d_active[(df_d_active['잔금_날짜'] >= start) & (df_d_active['잔금_날짜'] <= end)]
            else:
                sub = df_d_active
            
            mask_u = sub['구분'].astype(str).str.contains('USD|결제', case=False)
            d_u = sub[mask_u]
            d_c = sub[~mask_u]
            
            sum_usd = d_u[d_u['화폐단위']=='USD']['잔금_금액'].sum() + \
                      (d_u[d_u['화폐단위']=='CNY']['잔금_금액'].sum() * cny_to_usd_rate) + \
                      d_c[d_c['화폐단위']=='USD']['잔금_금액'].sum()
            sum_cny = d_c[d_c['화폐단위']=='CNY']['잔금_금액'].sum()
            sum_krw = (sum_cny * rate_cny) + (sum_usd * rate_usd)
            
            d_summary_rows.append({
                "기간": label,
                "다이렉트(CNY)": fmt_num(sum_cny),
                "다이렉트(USD)": fmt_num(sum_usd),
                "예상 KRW": fmt_num(sum_krw)
            })
        st.dataframe(pd.DataFrame(d_summary_rows), hide_index=True, use_container_width=True)
        st.markdown("---")

        # 2-2. 상세 리스트
        st.subheader("📋 상세 내역 검색")
        keyword = st.text_input("🔍 품목 검색")
        df_view = df_d_active.copy()
        if keyword:
            df_view = df_view[df_view['품목'].astype(str).str.contains(keyword, case=False)]
        
        if len(custom_date) == 2:
            start_d, end_d = custom_date
            df_view = df_view[(df_view['잔금_날짜'] >= pd.Timestamp(start_d)) & (df_view['잔금_날짜'] <= pd.Timestamp(end_d))]
        
        df_view = df_view.sort_values('잔금_날짜')
        df_view['🇰🇷 예상 KRW'] = df_view.apply(lambda r: r['잔금_금액'] * (rate_usd if r['화폐단위']=='USD' else rate_cny), axis=1)

        cols = ['잔금_날짜', '구분', '거래처', '품목', '화폐단위', '총_발주금액', '선급금_금액', '잔금_금액', '🇰🇷 예상 KRW', '진행단계']
        valid_cols = [c for c in cols if c in df_view.columns]
        df_disp = df_view[valid_cols].copy()
        
        for c in ['총_발주금액', '선급금_금액', '잔금_금액']:
            if c in df_disp.columns: df_disp[c] = df_disp[c].apply(fmt_num)
        if '🇰🇷 예상 KRW' in df_disp.columns: df_disp['🇰🇷 예상 KRW'] = df_disp['🇰🇷 예상 KRW'].apply(fmt_krw)
        if '잔금_날짜' in df_disp.columns: df_disp['잔금_날짜'] = df_disp['잔금_날짜'].dt.strftime('%Y-%m-%d')
        
        st.dataframe(df_disp, hide_index=True, use_container_width=True)
        
        # 합계
        t_cny = df_view[df_view['화폐단위']=='CNY']['잔금_금액'].sum()
        t_usd = df_view[df_view['화폐단위']=='USD']['잔금_금액'].sum()
        t_krw = df_view['🇰🇷 예상 KRW'].sum()
        
        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("CNY 합계", fmt_num(t_cny))
        c2.metric("USD 합계", fmt_num(t_usd))
        c3.metric("KRW 총 환산", fmt_krw(t_krw))

    # =======================================================
    # PAGE 3: 이우 관리
    # =======================================================
    elif menu == "🟦 이우(YIWU) 관리":
        st.header("🟦 이우 지출 & 장부")
        
        # 3-1. 기간별 이우 요약
        st.subheader("📅 기간별 이우 지출 계획")
        y_summary_rows = []
        for label, start, end in periods:
            if start and end:
                sub = df_y_active[(df_y_active['잔금_날짜'] >= start) & (df_y_active['잔금_날짜'] <= end)]
            else:
                sub = df_y_active
            
            s_cny = sub['잔금_금액'].sum()
            s_krw = s_cny * rate_cny
            
            y_summary_rows.append({
                "기간": label,
                "이우 지출(CNY)": fmt_num(s_cny),
                "예상 KRW": fmt_num(s_krw)
            })
        st.dataframe(pd.DataFrame(y_summary_rows), hide_index=True, use_container_width=True)
        st.markdown("---")
        
        # 3-2. 상세 리스트 (위)
        st.subheader("📋 지출 예정 상세")
        k_y = st.text_input("🔍 품목 검색", key='sy')
        df_y_view = df_y_active.copy()
        if k_y: df_y_view = df_y_view[df_y_view['품목'].astype(str).str.contains(k_y, case=False)]
        if len(custom_date) == 2:
            s_d, e_d = custom_date
            df_y_view = df_y_view[(df_y_view['잔금_날짜'] >= pd.Timestamp(s_d)) & (df_y_view['잔금_날짜'] <= pd.Timestamp(e_d))]
        
        df_y_view = df_y_view.sort_values('잔금_날짜')
        df_y_view['🇰🇷 예상 KRW'] = df_y_view['잔금_금액'] * rate_cny
        
        cols = ['잔금_날짜', '품목', '총_발주금액', '잔금_금액', '🇰🇷 예상 KRW', '진행단계']
        valid_cols = [c for c in cols if c in df_y_view.columns]
        df_disp = df_y_view[valid_cols].copy()
        
        for c in ['총_발주금액', '잔금_금액']: 
            if c in df_disp.columns: df_disp[c] = df_disp[c].apply(fmt_num)
        if '🇰🇷 예상 KRW' in df_disp.columns: df_disp['🇰🇷 예상 KRW'] = df_disp['🇰🇷 예상 KRW'].apply(fmt_krw)
        if '잔금_날짜' in df_disp.columns: df_disp['잔금_날짜'] = df_disp['잔금_날짜'].dt.strftime('%Y-%m-%d')
        
        st.dataframe(df_disp, hide_index=True, use_container_width=True)
        
        ty_cny = df_y_view['잔금_금액'].sum()
        ty_krw = df_y_view['🇰🇷 예상 KRW'].sum()
        st.info(f"합계: **{fmt_num(ty_cny)} CNY** (≈ {fmt_krw(ty_krw)} KRW)")

        st.markdown("---")

        # 3-3. 장부 잔고 (아래)
        st.subheader("📒 장부 잔고 (최근 입출금)")
        st.write(f"현재 누적 잔고: **{fmt_num(yiwu_balance)} CNY**")
        if not df_l.empty:
            df_l_sort = df_l.sort_values('날짜', ascending=False).head(20) if '날짜' in df_l.columns else df_l.head(20)
            if '날짜' in df_l_sort.columns: df_l_sort['날짜'] = df_l_sort['날짜'].dt.strftime('%Y-%m-%d')
            for nc in ['입금액(CNY)', '사용금액(CNY)', '누적잔고(CNY)']:
                if nc in df_l_sort.columns: df_l_sort[nc] = df_l_sort[nc].apply(fmt_num)
            st.dataframe(df_l_sort, hide_index=True, use_container_width=True)

else:
    st.info("👈 엑셀 파일을 업로드해주세요.")