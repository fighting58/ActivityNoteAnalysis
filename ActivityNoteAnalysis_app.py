import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Activity Log", layout="wide")
st.title("활동일지 분석")

# 파일 업로드 위젯
uploaded_file = st.file_uploader("파일 업로드 (xlsx 또는 csv 형식)", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        # 파일 형식 확인 및 데이터프레임 변환
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file, dtype=str, engine='openpyxl')
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, dtype=str, encoding='euc-kr')

        # 특정 컬럼의 데이터 타입을 변경
        df["활동직원수"] = df["활동직원수"].astype(int)
        st.write("원본 데이터프레임:")
        st.dataframe(df)

        if st.button("활동직원별 재생성"):
            df['활동직원'] = df['활동직원'].str.strip(', ').str.split(',').explode().str.strip()
            df['활동직원수'] = 1
            st.session_state['new_df'] = df.reset_index(drop=True)

        if 'new_df' in st.session_state:
            st.write("직원별 데이터프레임:")
            st.markdown('간단한 편집은 가능합니다.  \n다만, 전체 행 또는 열의 추가/삭제는 불가하며, 편집된 파일을 다운로드하여 수정 후 재업로드하거나 원본을 수정하여야 합니다.')
            edited_df = st.data_editor(st.session_state['new_df'])

            col1, col2, _ = st.columns([0.1, 0.05, 1])

            with col1:
                if st.button("수정 데이터 저장"):
                    st.session_state['new_df'] = edited_df
                    st.success("데이터프레임이 업데이트되었습니다.")
            with col2:
                csv = st.session_state['new_df'].to_csv(index=False).encode('euc-kr')
                st.download_button(label=":floppy_disk:", data=csv, file_name="activityLog_by_member.csv", mime="text/csv")

            st.divider()
            col3, col4 = st.columns([0.5, 0.5])
            member_activity_stat = st.session_state['new_df'][["활동직원", "활동직원수"]].groupby("활동직원").sum().reset_index()

            with col3:
                col3_sub1, col3_sub2, col3_sub3 = st.columns(3)
                with col3_sub1:
                    activity_threshold = st.slider("활동건수 기준값", 0, 100, 0)
                with col3_sub2:
                    width = st.slider("그래프 너비", 400, 1200, 800)
                with col3_sub3:
                    height = st.slider("그래프 높이", 300, 900, 600)

                member_activity_stat.rename(columns={"활동직원수": "활동건수"}, inplace=True)
                filtered_member_activity_stat = member_activity_stat[member_activity_stat["활동건수"] >= activity_threshold].reset_index()
                top3 = filtered_member_activity_stat.nlargest(3, "활동건수")
                colors = ['red' if row in top3.index else 'blue' for row in filtered_member_activity_stat.index]

                fig1 = px.bar(filtered_member_activity_stat, x="활동직원", y="활동건수", title="직원별 활동건수", template="gridon", color=colors)
                fig1.update_layout(
                    title={'text': "직원별 활동건수", 'x': 0.5, 'xanchor': 'center', 'yanchor': 'top'},
                    width=width,
                    height=height,
                    xaxis=dict(tickfont=dict(size=9), tickangle=90),
                    showlegend=False
                )
                fig1.update_traces(hovertemplate='활동직원: %{x}<br>활동건수: %{y}<extra></extra>')

                st.plotly_chart(fig1, use_container_width=False)

    except Exception as e:
        st.error(f"Error processing file: {e}")
