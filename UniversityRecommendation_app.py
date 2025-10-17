import streamlit as st
import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.metrics.pairwise import cosine_similarity
import requests

# 페이지 설정
st.set_page_config(
    page_title="대학 추천 시스템",
    page_icon="🎓",
    layout="wide"
)

# Secrets에서 API 키 가져오기
try:
    # 예시: 대학 정보 API 키
    api_key = st.secrets["university_api"]["key"]
    api_endpoint = st.secrets["university_api"]["endpoint"]
    
    # 예시: 데이터베이스 연결 정보
    db_config = st.secrets["database"]
    
except KeyError as e:
    st.error(f"필요한 설정이 누락되었습니다: {e}")
    st.info("Streamlit Cloud의 Secrets에서 API 키와 데이터베이스 정보를 설정해주세요.")
    st.stop()

# 2025_2021_result.csv 파일 로드
@st.cache_data
def load_data():
    """대학 입시 데이터 로드"""
    try:
        df = pd.read_csv('2025_2021_result.csv', encoding='utf-8-sig')
        return df
    except FileNotFoundError:
        st.error("데이터 파일을 찾을 수 없습니다: 2025_2021_result.csv")
        return None
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {e}")
        return None

# 외부 API에서 추가 대학 정보 가져오기 (Secrets 사용)
@st.cache_data(ttl=3600)  # 1시간 캐시
def get_university_details(university_name):
    """외부 API에서 대학 상세 정보 가져오기"""
    try:
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        
        response = requests.get(
            f"{api_endpoint}/universities/{university_name}",
            headers=headers,
            timeout=10
        )
        
        if response.status_code == 200:
            return response.json()
        else:
            return None
            
    except Exception as e:
        st.warning(f"API 호출 실패: {e}")
        return None

# 데이터베이스 연결 (Secrets 사용)
@st.cache_resource
def init_db_connection():
    """데이터베이스 연결 초기화"""
    try:
        # PostgreSQL 예시
        import psycopg2
        
        conn = psycopg2.connect(
            host=db_config["host"],
            port=db_config["port"],
            database=db_config["database"],
            user=db_config["username"],
            password=db_config["password"]
        )
        return conn
    except Exception as e:
        st.error(f"데이터베이스 연결 실패: {e}")
        return None

# 사용자 선호도 저장 (데이터베이스 사용)
def save_user_preference(user_id, preferences):
    """사용자 선호도를 데이터베이스에 저장"""
    conn = init_db_connection()
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO user_preferences (user_id, preferences, created_at)
                VALUES (%s, %s, NOW())
                ON CONFLICT (user_id) 
                DO UPDATE SET preferences = %s, updated_at = NOW()
            """, (user_id, preferences, preferences))
            conn.commit()
            return True
        except Exception as e:
            st.error(f"선호도 저장 실패: {e}")
            return False
        finally:
            conn.close()

# 메인 앱
def main():
    st.title("🎓 대학 추천 시스템")
    st.markdown("### 당신에게 맞는 대학을 찾아드립니다")
    
    # 데이터 로드
    df = load_data()
    if df is None:
        return
    
    # 사이드바 - 필터 옵션
    with st.sidebar:
        st.header("검색 조건")
        
        # 지역 필터
        regions = ['전체'] + sorted(df['지역'].unique().tolist()) if '지역' in df.columns else ['전체']
        selected_region = st.selectbox("지역 선택", regions)
        
        # 전공 필터
        majors = ['전체'] + sorted(df['전공'].unique().tolist()) if '전공' in df.columns else ['전체']
        selected_major = st.selectbox("전공 선택", majors)
        
        # 성적 입력
        st.subheader("성적 정보")
        gpa = st.slider("내신 등급", 1.0, 9.0, 3.0, 0.1)
        
        # 수능 성적
        korean = st.slider("국어 백분위", 0, 100, 70)
        math = st.slider("수학 백분위", 0, 100, 70)
        english = st.slider("영어 등급", 1, 9, 3)
        
        # 추가 선호도
        st.subheader("추가 선호 사항")
        prefer_dorm = st.checkbox("기숙사 제공 대학 선호")
        prefer_scholarship = st.checkbox("장학금 혜택 우선 고려")
        
        # 환경 정보 표시 (디버그용)
        if st.secrets.get("debug", False):
            st.info(f"환경: {st.secrets.get('environment', 'production')}")
    
    # 메인 컨텐츠
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("추천 대학 목록")
        
        # 필터링
        filtered_df = df.copy()
        if selected_region != '전체' and '지역' in df.columns:
            filtered_df = filtered_df[filtered_df['지역'] == selected_region]
        if selected_major != '전체' and '전공' in df.columns:
            filtered_df = filtered_df[filtered_df['전공'] == selected_major]
        
        # 추천 알고리즘 (간단한 예시)
        # 실제로는 더 복잡한 알고리즘 사용
        filtered_df['추천점수'] = (
            (100 - abs(filtered_df.get('평균내신', 3) - gpa) * 10) * 0.3 +
            (korean * 0.2) + (math * 0.2) + ((10 - english) * 10 * 0.1) +
            np.random.rand(len(filtered_df)) * 20  # 랜덤 요소
        )
        
        # 상위 10개 대학 표시
        top_universities = filtered_df.nlargest(10, '추천점수')
        
        for idx, row in top_universities.iterrows():
            with st.expander(f"{row.get('대학명', 'Unknown')} - 추천점수: {row['추천점수']:.1f}"):
                col_a, col_b = st.columns(2)
                
                with col_a:
                    st.write(f"**지역**: {row.get('지역', 'N/A')}")
                    st.write(f"**전공**: {row.get('전공', 'N/A')}")
                    st.write(f"**평균 내신**: {row.get('평균내신', 'N/A')}")
                
                with col_b:
                    # API에서 추가 정보 가져오기
                    details = get_university_details(row.get('대학명', ''))
                    if details:
                        st.write(f"**취업률**: {details.get('employment_rate', 'N/A')}%")
                        st.write(f"**등록금**: {details.get('tuition', 'N/A')}만원")
                    
                    if st.button(f"자세히 보기", key=f"detail_{idx}"):
                        st.session_state['selected_university'] = row.get('대학명', '')
    
    with col2:
        st.subheader("통계 정보")
        
        # 간단한 통계 차트
        if not filtered_df.empty:
            st.metric("검색된 대학 수", len(filtered_df))
            st.metric("평균 경쟁률", f"{filtered_df.get('경쟁률', pd.Series([0])).mean():.1f}:1")
            
            # 지역별 분포 차트
            if '지역' in filtered_df.columns:
                region_counts = filtered_df['지역'].value_counts()
                st.bar_chart(region_counts)
    
    # 선택된 대학 상세 정보
    if 'selected_university' in st.session_state and st.session_state['selected_university']:
        st.divider()
        st.subheader(f"📍 {st.session_state['selected_university']} 상세 정보")
        
        # 여기에 상세 정보 표시
        university_data = filtered_df[filtered_df['대학명'] == st.session_state['selected_university']].iloc[0]
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("지역", university_data.get('지역', 'N/A'))
        with col2:
            st.metric("전공", university_data.get('전공', 'N/A'))
        with col3:
            st.metric("경쟁률", f"{university_data.get('경쟁률', 0):.1f}:1")
    
    # 사용자 피드백 (선택사항)
    if st.button("추천 결과 저장"):
        # 사용자 ID는 세션 상태나 로그인 시스템에서 가져옴
        user_id = st.session_state.get('user_id', 'anonymous')
        preferences = {
            'region': selected_region,
            'major': selected_major,
            'gpa': gpa,
            'scores': {'korean': korean, 'math': math, 'english': english}
        }
        
        if save_user_preference(user_id, str(preferences)):
            st.success("선호도가 저장되었습니다!")
        else:
            st.warning("선호도 저장에 실패했습니다.")

# 앱 실행
if __name__ == "__main__":
    main()
