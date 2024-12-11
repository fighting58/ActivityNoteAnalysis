import streamlit as st
import pandas as pd
import plotly.express as px
import io

# 파일 처리를 위한 캐싱 데코레이터
@st.cache_data
def process_dataframe(df):
    new_rows = []
    for index, row in df.iterrows(): 
        employees = row["활동직원"].strip(", ").split(",") 
        for employee in employees: 
            new_row = row.copy() 
            new_row["활동직원"] = employee.strip() 
            new_row["활동직원수"] = 1 
            new_rows.append(new_row)
    return pd.DataFrame(new_rows).reset_index(drop=True)

# 색상 선택 유틸리티 함수
def get_color_sequence(data, top_n=3, base_column="활동건수", top_color='orange', default_color='dodgerblue'):
    return [top_color if i in data.nlargest(top_n, base_column).index else default_color 
            for i in data.index]

# 금액관련 컬럼 정제 함수
def clean_price_column(df, column_names:list):

    for column_name in column_names:
        # 1. Null 값을 "0"으로 채우기
        df[column_name] = df[column_name].fillna("0")        
        # 2. 천 단위 구분 기호 ",", " "를 모두 제거
        df[column_name] = df[column_name].str.replace(",", "").str.replace(" ", "")        
        # 3. 숫자가 아닌 값을 "0"으로 교체
        df[column_name] = df[column_name].apply(lambda x: x if x.isnumeric() else "0")        
        # 4. "Price" 컬럼을 int 형으로 변환
        df[column_name] = df[column_name].astype(int)
    
    return df

# 그래프 슬라이더 생성 함수
def create_graph_sliders(show=["threshold", "width", "height"], threshold_label="활동건수 기준값", threshold_max=300, 
                          width_label="그래프 너비", height_label="그래프 높이",
                          threshold_key='activity_threshold', 
                          width_key='width', 
                          height_key='height'):
    col_sub1, col_sub2, col_sub3 = st.columns(3)
    threshold, width, height = 0, 800, 600
    if "threshold" in show:
        with col_sub1:
            threshold = st.slider(threshold_label, 0, threshold_max, 0, key=threshold_key)        
    if "width" in show:
        with col_sub2:
            width = st.slider(width_label, 400, 1200, 800, key=width_key)
    if "height" in show:
        with col_sub3:
            height = st.slider(height_label, 300, 900, 600, key=height_key)
        return threshold, width, height
    
# streamlit 에서 excel 파일 다운로드 위한 함수
def to_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    processed = output.getvalue()
    return processed

# Streamlit 앱 설정
st.set_page_config(page_title="활동 로그 분석", layout="wide")
st.title("📊 활동일지 분석")

# 파일 업로드 위젯
uploaded_file = st.file_uploader("파일 업로드 (xlsx 또는 csv 형식)", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        # 파일 형식 확인 및 데이터프레임 변환
        try:
            if uploaded_file.name.endswith('.xlsx'):
                df = pd.read_excel(uploaded_file, dtype=str, engine='openpyxl')
            elif uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, dtype=str)
        except pd.errors.EmptyDataError:
            st.error("업로드된 파일이 비어 있습니다.")
            st.stop()
        except UnicodeDecodeError:
            st.error("파일 인코딩을 확인해주세요.")
            st.stop()

        # 데이터 유효성 검사
        required_columns = ["활동직원", "활동직원수", "소속", "사업분류"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"다음 필수 컬럼이 누락되었습니다: {', '.join(missing_columns)}")
            st.stop()

        # 데이터 타입 변환
        df["활동직원수"] = df["활동직원수"].astype(int)
        
        st.success("📂 데이터 로딩 완료!")
        st.write("### 원본 데이터 미리보기")
        st.dataframe(df.head())

        if st.button("🔄 활동직원별 데이터 재생성"):
            # 데이터프레임 처리
            new_df = process_dataframe(df)
            new_df = clean_price_column(new_df, ["COS 연계정보(완료금액)"]) 
            st.session_state['new_df'] = new_df

        if 'new_df' in st.session_state:
            st.write("### 직원별 데이터")
            st.markdown("""
                        💡 간단한 편집 가능합니다.  
                        전체 행/열 추가/삭제는 불가하며, 다운로드 후 수정 권장
                        """)
            
            edited_df = st.data_editor(st.session_state['new_df'])

            col1, col2, col3, _ = st.columns([0.15, 0.1, 0.1, 0.65])

            with col1:
                if st.button("✅ 변경데이터 적용"):
                    st.session_state['new_df'] = edited_df
                    st.success("데이터프레임 업데이트 완료!")
            
            with col2:
                timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                csv = st.session_state['new_df'].to_csv(index=False)
                st.download_button(
                    label="💾 다운로드(csv)", 
                    data=csv, 
                    file_name=f"activityLog_{timestamp}.csv", 
                    mime="text/csv"
                )
            with col3:
                xlsx = st.session_state['new_df']
                st.download_button(
                    label="💾 다운로드(xlsx)", 
                    data=to_excel(xlsx), 
                    file_name=f"activityLog_{timestamp}.xlsx", 
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            st.divider()
            jisa_order = st.session_state['new_df']["소속"].unique().tolist()
            member_grade = st.session_state['new_df']["활동직급"].unique().tolist()
            member_grade.sort()
            member_grade.insert(0, "전체")


            # 직원별 활동건수 그래프
            st.markdown("<h3>1) 직원별 활동건수 그래프</h3>", unsafe_allow_html=True)
            member_activity_stat = st.session_state['new_df'][["활동직급", "활동직원", "활동직원수"]].groupby(["활동직급", "활동직원"]).sum().reset_index()
            with st.expander("직원별 활동건수 그래프"):
                activity_threshold, width, height = create_graph_sliders(
                    threshold_label="활동건수 기준값", 
                    threshold_max=50,
                    threshold_key='activity_threshold',
                    width_key='width',
                    height_key='height'
                )

                member_activity_stat.rename(columns={"활동직원수": "활동건수"}, inplace=True)
                filtered_member_activity_stat = member_activity_stat[member_activity_stat["활동건수"] >= activity_threshold]

                grade = st.selectbox("직급 선택", member_grade, key="member_grade")
                if grade == "전체":
                    filtered_grade_member_activity_stat = filtered_member_activity_stat
                else:
                    filtered_grade_member_activity_stat = filtered_member_activity_stat[filtered_member_activity_stat["활동직급"] == grade]
                
                colors = get_color_sequence(filtered_grade_member_activity_stat)

                fig1 = px.bar(filtered_grade_member_activity_stat, x="활동직원", y="활동건수", 
                            title=None, template="gridon", 
                            color="활동직원",
                            color_discrete_sequence=colors)
                
                fig1.update_layout(
                    width=width, height=height,
                    xaxis=dict(tickfont=dict(size=9), tickangle=90),
                    showlegend=False
                )
                fig1.update_traces(hovertemplate='활동직원: %{x}<br>활동건수: %{y}<extra></extra>')
                st.plotly_chart(fig1, use_container_width=False)

  
            # 지사별 활동건수 그래프
            st.markdown("<h3>2) 지사별 활동건수 그래프</h3>", unsafe_allow_html=True)
            jisa_activity_stat = st.session_state['new_df'][["소속", "활동직원수"]].groupby("소속").sum().reset_index()
            with st.expander("지사별 활동건수 그래프"):
                activity_threshold2, width2, height2 = create_graph_sliders(
                    show=["width", "height"],
                    width_key='width2',
                    height_key='height2'
                )

                jisa_activity_stat.rename(columns={"소속": "지사", "활동직원수":"활동건수"}, inplace=True)
                
                jisa_list = jisa_activity_stat["지사"].tolist()
                colors2 = get_color_sequence(jisa_activity_stat)
                colors2 = [colors2[jisa_list.index(jisaname)] for jisaname in jisa_order]

                fig2 = px.bar(jisa_activity_stat, x="지사", y="활동건수", 
                            title=None, template="gridon", 
                            color="지사",
                            color_discrete_sequence=colors2,
                            category_orders={"지사": jisa_order})
                fig2.update_layout(
                    width=width2, height=height2,
                    xaxis=dict(tickfont=dict(size=9), tickangle=90),
                    showlegend=False
                )
                fig2.update_traces(hovertemplate='지사: %{x}<br>활동건수: %{y}<extra></extra>')
                st.plotly_chart(fig2, use_container_width=False)


            # 사업분류별 활동건수 그래프
            st.markdown("<h3>3) 사업분류별 활동건수 그래프</h3>", unsafe_allow_html=True)
            business_class_stat = st.session_state['new_df'][["사업분류", "활동직원수"]].groupby("사업분류").sum().reset_index()
            with st.expander("사업분류별 활동건수 그래프"):
                _, width3, height3 = create_graph_sliders(
                    show=["width", "height"],
                    width_key='width3',
                    height_key='height3'
                )

                business_class_stat.rename(columns={"활동직원수": "활동건수"}, inplace=True)

                fig3 = px.pie(business_class_stat, 
                            names="사업분류", 
                            values="활동건수", 
                            title=None, 
                            color=business_class_stat.index, 
                            hole=0.3)
                fig3.update_layout(
                    width=width3, 
                    height=height3,
                    showlegend=False
                )
                fig3.update_traces(
                    textinfo='percent+label', 
                    hovertemplate='사업분류: %{label}<br>활동건수: %{value}<extra></extra>'
                )
                st.plotly_chart(fig3, use_container_width=False)


            # 지사별 활동사업 그래프
            st.markdown("<h3>4) 지사별 활동사업 그래프</h3>", unsafe_allow_html=True)
            jisa_business_stat = st.session_state['new_df'][["소속",'사업분류', "활동직원수"]].groupby(["소속", "사업분류"]).sum().reset_index()
            with st.expander("지사별 활동사업 그래프"):
                _, width4, height4 = create_graph_sliders(
                    show=["width", "height"],
                    width_key='width4',
                    height_key='height4'
                )

                jisa_business_stat.rename(columns={"소속": "지사", "활동직원수":"활동구분"}, inplace=True)   

                fig4 = px.bar(jisa_business_stat, 
                            x="지사", 
                            y="활동구분", 
                            title=None, 
                            template="gridon", 
                            color='사업분류',
                            category_orders={"지사": jisa_order})
                fig4.update_layout(
                    width=width4, 
                    height=height4,
                    xaxis=dict(tickfont=dict(size=9), tickangle=90),
                    showlegend=True
                )
                st.plotly_chart(fig4, use_container_width=False)

            # 직원별 완료금액 그래프
            st.markdown("<h3>4) 직원별 완료금액 그래프</h3>", unsafe_allow_html=True)
            income_df = st.session_state['new_df'][st.session_state['new_df']["COS 연계정보(완료금액)"]>0]
            member_income_dup = income_df[["활동직급","활동직원", "COS 연계정보(완료금액)", "사업명"]].groupby(["활동직급","활동직원"]).agg(list)
            member_income_uniq = member_income_dup.applymap(lambda x: x[0])
            member_income_stat = member_income_uniq.groupby(["활동직급", "활동직원"]).sum().reset_index()
            st.dataframe(member_income_stat)

            with st.expander("직원별 완료금액 그래프"):
                _, width5, height5 = create_graph_sliders(
                    show=["width", "height"],
                    width_key='width5',
                    height_key='height5'
                )

                fig5 = px.bar(member_income_stat, 
                            x="활동직원", 
                            y="COS 연계정보(완료금액)", 
                            title=None, 
                            template="gridon", 
                            color='활동직원')
                
                fig5.update_layout(
                    width=width4, 
                    height=height4,
                    xaxis=dict(tickfont=dict(size=9), tickangle=90),
                    showlegend=True
                )
                st.plotly_chart(fig5, use_container_width=False)


    except Exception as e:
        st.error(f"파일 처리 중 오류 발생: {e}")
