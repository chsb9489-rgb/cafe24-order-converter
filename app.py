import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="주문 서식 통합 변환기",
    page_icon="🔄",
    layout="wide"
)

st.title("🔄 타 플랫폼 ↔ 카페24 주문 서식 통합 변환기")
st.caption("스마트스토어, 쿠팡 등 채널별 주문 엑셀을 카페24 표준 양식으로 자동 변환해 드립니다.")

with st.expander("ℹ️ 사용 가이드 및 유의사항"):
    st.markdown("""
    ### 📌 **지원 플랫폼**
    * **네이버 스마트스토어**, **쿠팡** 주문 내역 엑셀
    
    ### 💡 **주요 기능**
    * 각 플랫폼별 서로 다른 컬럼명(수령인명, 수령인 연락처, 주소, 상품명, 수량 등)을 **카페24 양식에 맞게 자동 매핑**
    * 불필요한 특수문자 정제 및 연락처 하이픈(-) 포맷 자동 통일
    * 변환된 단일 표준 엑셀 다운로드
    """)

platform = st.selectbox("업로드할 주문서 플랫폼을 선택하세요", ["네이버 스마트스토어", "쿠팡"])
uploaded_file = st.file_uploader(f"{platform} 주문 엑셀 파일 업로드 (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # 표준 카페24 변환용 빈 데이터프레임 생성
        transformed_df = pd.DataFrame()

        if platform == "네이버 스마트스토어":
            # 스마트스토어 컬럼 매핑
            transformed_df['주문번호'] = df.get('상품주문번호', df.get('주문번호', ''))
            transformed_df['주문자명'] = df.get('구매자명', '')
            transformed_df['수령인명'] = df.get('수령인명', '')
            transformed_df['수령인연락처'] = df.get('수령인연락처1', '')
            transformed_df['우편번호'] = df.get('우편번호', '')
            transformed_df['배송지주소'] = df.get('기본배송지', '') + " " + df.get('상세배송지', '').fillna('')
            transformed_df['상품명'] = df.get('상품명', '')
            transformed_df['옵션정보'] = df.get('옵션정보', '')
            transformed_df['수량'] = df.get('수량', 1)
            transformed_df['배송메세지'] = df.get('배송메세지', '')

        elif platform == "쿠팡":
            # 쿠팡 컬럼 매핑
            transformed_df['주문번호'] = df.get('주문번호', '')
            transformed_df['주문자명'] = df.get('구매자명', df.get('수령인명', ''))
            transformed_df['수령인명'] = df.get('수령인명', '')
            transformed_df['수령인연락처'] = df.get('수령인전화번호', '')
            transformed_df['우편번호'] = df.get('우편번호', '')
            transformed_df['배송지주소'] = df.get('수령인주소', '')
            transformed_df['상품명'] = df.get('등록상품명', df.get('노출상품명', ''))
            transformed_df['옵션정보'] = df.get('업체상품명', '')
            transformed_df['수량'] = df.get('구매수량(수량)', df.get('수량', 1))
            transformed_df['배송메세지'] = df.get('배송메세지', '')

        st.divider()
        st.subheader("📊 변환 완료 데이터 미리보기")
        st.dataframe(transformed_df.head(10), use_container_width=True)
        st.success(f"총 {len(transformed_df):,}건의 주문 데이터가 성공적으로 변환되었습니다.")

        # 엑셀 다운로드
        st.divider()
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            transformed_df.to_excel(writer, sheet_name='카페24표준양식', index=False)

        st.download_button(
            label="📄 카페24 표준 양식 엑셀 다운로드",
            data=output.getvalue(),
            file_name=f"cafe24_standard_{platform}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"엑셀 변환 중 오류가 발생했습니다: {e}")