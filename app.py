import streamlit as st
import pandas as pd
import os
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="재고 수불 검색 프로그램", layout="wide")

# ==========================================
# 🔑 비밀번호 설정 (원하는 비밀번호로 변경하세요)
# ==========================================
PASSWORD = "1234"

def check_password():
    """비밀번호 검증 함수"""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("🔒 재고 데이터 검색 프로그램 로그인")
        st.caption("팀원 전용 프로그램입니다. 비밀번호를 입력해주세요.")
        
        user_password = st.text_input("비밀번호 입력", type="password")
        
        if st.button("로그인", use_container_width=True):
            if user_password == PASSWORD:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ 비밀번호가 올바르지 않습니다.")
        return False
    return True

# 비밀번호 통과 시 실행
if check_password():
    
    st.title("📦 재고 데이터 검색 프로그램")
    st.caption("키워드 입력, LOT 및 기호 정밀 필터링, 엑셀 파일 즉시 업로드 기능을 제공합니다.")

    # 1. 엑셀 파일 직접 업로드 기능
    st.sidebar.header("📁 데이터 파일 관리")
    uploaded_file = st.sidebar.file_uploader("최신 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])
    
    FILE_NAME = "재고수불26.04.02.xlsx"

    # 엑셀 데이터 로드 함수
    def load_data(source):
        try:
            df = pd.read_excel(source)
            unnamed_cols = [col for col in df.columns if str(col).startswith('Unnamed')]
            if unnamed_cols:
                df = df.drop(columns=unnamed_cols)
            return df
        except Exception as e:
            st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
            return None

    # 업로드된 파일이 있으면 우선 사용, 없으면 저장된 기본 파일 사용
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        st.sidebar.success("✅ 새로운 엑셀 파일이 적용되었습니다!")
    elif os.path.exists(FILE_NAME):
        df = load_data(FILE_NAME)
    else:
        df = None

    # 상단 우측 버튼 영역
    col1, col2, col3 = st.columns([6, 2, 2])
    with col2:
        if st.button("🔄 화면 새로고침", use_container_width=True):
            st.rerun()
    with col3:
        if st.button("🔒 잠금 (로그아웃)", use_container_width=True):
            st.session_state["password_correct"] = False
            st.rerun()

    def highlight_columns(s):
        styles = [''] * len(s)
        for i, col in enumerate(s.index):
            if col == '입고수량':
                styles[i] = 'color: #E53E3E; font-weight: bold;'
            elif col == '출고수량':
                styles[i] = 'color: #3182CE; font-weight: bold;'
        return styles

    def clean_format(val):
        if pd.isna(val) or val == "" or str(val).strip() in ["None", "nan"]:
            return ""
        try:
            num = float(val)
            if num.is_integer():
                return f"{int(num)}"
            return f"{num:g}"
        except (ValueError, TypeError):
            return str(val)

    def convert_df_to_excel(df_to_export):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_to_export.to_excel(writer, index=False, sheet_name='검색결과')
        return output.getvalue()

    if df is not None:
        # 검색창, LOT 필터, 기호 필터를 3개 컬럼으로 배치
        col_search, col_lot, col_symbol = st.columns([5, 4, 3])
        
        with col_search:
            search_term = st.text_input("🔍 품목명/키워드를 입력하세요 (예: 신만복, 시금치 등)", "")

        filtered_df = df.copy()

        # 키워드 필터링
        if search_term:
            mask = filtered_df.astype(str).apply(lambda row: row.str.contains(search_term, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]

        # LOT 번호 필터링
        with col_lot:
            if 'LOT번호' in filtered_df.columns and not filtered_df.empty:
                available_lots = filtered_df['LOT번호'].dropna().astype(str).unique().tolist()
                available_lots.sort()
                lot_options = ["전체 LOT 보기"] + available_lots
                selected_lot = st.selectbox("🏷️ 특정 LOT 번호 필터", lot_options)
                
                if selected_lot != "전체 LOT 보기":
                    filtered_df = filtered_df[filtered_df['LOT번호'].astype(str) == selected_lot]

        # 기호 필터링 (신규 추가!)
        with col_symbol:
            if '기호' in filtered_df.columns and not filtered_df.empty:
                # 기호 데이터를 정수나 깔끔한 문자열 형태로 정렬
                raw_symbols = filtered_df['기호'].dropna().tolist()
                clean_symbols = []
                for s in raw_symbols:
                    try:
                        clean_symbols.append(str(int(float(s))))
                    except:
                        clean_symbols.append(str(s).strip())
                
                clean_symbols = sorted(list(set(clean_symbols)))
                symbol_options = ["전체 기호 보기"] + clean_symbols
                selected_symbol = st.selectbox("📌 특정 기호 필터", symbol_options)
                
                if selected_symbol != "전체 기호 보기":
                    # 선택한 기호와 일치하는 행만 선택
                    filtered_df = filtered_df[filtered_df['기호'].astype(str).apply(clean_format) == selected_symbol]

        if search_term or (selected_lot != "전체 LOT 보기" if 'selected_lot' in locals() else False) or (selected_symbol != "전체 기호 보기" if 'selected_symbol' in locals() else False):
            st.divider()
            col_res_info, col_download = st.columns([7, 5])
            
            with col_res_info:
                st.success(f"검색 결과: 총 **{len(filtered_df)}**건이 조회되었습니다.")
                
            with col_download:
                if not filtered_df.empty:
                    excel_data = convert_df_to_excel(filtered_df)
                    st.download_button(
                        label="📥 검색 결과 엑셀 파일 다운로드",
                        data=excel_data,
                        file_name=f"검색결과_{search_term if search_term else '필터'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
            styled_df = filtered_df.style.format(clean_format).apply(highlight_columns, axis=1)
            st.dataframe(styled_df, use_container_width=True)

        else:
            st.info("검색어 입력 또는 LOT/기호 필터를 선택하시면 해당하는 데이터만 즉시 필터링됩니다.")
            styled_df = df.style.format(clean_format).apply(highlight_columns, axis=1)
            st.dataframe(styled_df, use_container_width=True)

    else:
        st.error("엑셀 데이터를 찾을 수 없습니다. 왼쪽 사이드바에서 최신 엑셀 파일을 업로드해 주세요!")