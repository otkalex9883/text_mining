import io
import re
from collections import Counter

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from kiwipiepy import Kiwi


# =========================================================
# 기본 설정
# =========================================================

st.set_page_config(
    page_title="고객지원 상담내용 텍스트마이닝",
    page_icon="📊",
    layout="wide"
)

st.title("📊 고객지원 상담내용 텍스트마이닝")
st.caption(
    "제품명을 검색하면 해당 제품 상담내용의 주요 키워드와 워드클라우드를 보여줍니다."
)

kiwi = Kiwi()


# =========================================================
# 기본 불용어
# =========================================================

DEFAULT_STOPWORDS = {
    "고객", "고객님", "상담", "상담원",
    "문의", "문의사항", "확인", "요청",
    "답변", "안내", "내용", "제품",
    "말씀", "연락", "전화", "통화",
    "관련", "안녕하세요", "감사합니다",
    "감사", "부탁", "부탁드립니다",
    "드리겠습니다", "드렸습니다",
    "있습니다", "합니다", "같습니다",
    "싶습니다", "해주세요", "해당",
    "저희", "당사", "이번", "오늘",
    "어제", "그리고", "그런데",
    "정도", "부분", "때문",
    "것", "수", "점", "거", "때",
    "좀", "더", "또", "및", "등",
    "내", "외", "중", "대한",
    "대해", "위해", "통해",
    "경우", "사항", "처리",
    "진행", "예정", "가능", "여부"
}


# =========================================================
# 텍스트 정리
# =========================================================

def clean_text(text):
    if pd.isna(text):
        return ""

    text = str(text)

    text = re.sub(
        r"\b(?:01[016789]|02|0[3-9]\d)[-\s]?\d{3,4}[-\s]?\d{4}\b",
        " ",
        text
    )

    text = re.sub(
        r"Caller\s*:\s*\S+",
        " ",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
        " ",
        text
    )

    text = re.sub(
        r"https?://\S+|www\.\S+",
        " ",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"[^가-힣A-Za-z0-9\s]",
        " ",
        text
    )

    text = re.sub(
        r"\b\d+\b",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    ).strip()

    return text


# =========================================================
# 명사 추출
# =========================================================

def extract_nouns(text, stopwords, min_length=2):
    text = clean_text(text)

    if not text:
        return []

    try:
        tokens = kiwi.tokenize(text)
    except Exception:
        return []

    words = []

    for token in tokens:
        word = token.form
        tag = token.tag

        if tag not in {"NNG", "NNP", "NNB"}:
            continue

        if len(word) < min_length:
            continue

        if word in stopwords:
            continue

        if word.isdigit():
            continue

        words.append(word)

    return words


# =========================================================
# 파일 읽기
# =========================================================

def read_uploaded_file(uploaded_file):
    file_name = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if file_name.endswith(".csv"):

        for encoding in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
            try:
                return pd.read_csv(
                    io.BytesIO(file_bytes),
                    encoding=encoding
                )
            except Exception:
                continue

        raise ValueError("CSV 인코딩을 읽을 수 없습니다.")

    if file_name.endswith(".xlsx"):

        try:
            return pd.read_excel(
                io.BytesIO(file_bytes),
                engine="openpyxl"
            )

        except Exception as xlsx_error:

            try:
                tables = pd.read_html(io.BytesIO(file_bytes))
                if tables:
                    return tables[0]
            except Exception:
                pass

            raise ValueError(f"XLSX 파일을 읽을 수 없습니다: {xlsx_error}")

    if file_name.endswith(".xls"):

        try:
            return pd.read_excel(
                io.BytesIO(file_bytes),
                engine="xlrd"
            )

        except Exception:

            try:
                tables = pd.read_html(io.BytesIO(file_bytes))
                if tables:
                    return tables[0]
            except Exception:
                pass

            raise ValueError(
                "XLS 파일을 읽을 수 없습니다. "
                "Excel에서 XLSX 형식으로 다시 저장해 주세요."
            )

    raise ValueError("xlsx, xls, csv 파일만 지원합니다.")


# =========================================================
# 열 이름 찾기
# =========================================================

def find_column(columns, candidates):
    normalized_columns = {
        str(column).strip().replace(" ", "").lower(): column
        for column in columns
    }

    for candidate in candidates:
        normalized_candidate = candidate.strip().replace(" ", "").lower()
        if normalized_candidate in normalized_columns:
            return normalized_columns[normalized_candidate]

    return None


# =========================================================
# 한글 폰트 찾기
# =========================================================

def get_korean_font():
    font_candidates = [
        "fonts/NanumGothic.ttf",
        "C:/Windows/Fonts/malgun.ttf",
        "C:/Windows/Fonts/NanumGothic.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    ]

    for font_path in font_candidates:
        try:
            with open(font_path, "rb"):
                return font_path
        except Exception:
            pass

    return None


# =========================================================
# 파일 업로드
# =========================================================

uploaded_file = st.file_uploader(
    "상담 데이터 파일을 업로드하세요.",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is None:
    st.info("엑셀 또는 CSV 파일을 업로드하면 분석이 시작됩니다.")
    st.markdown(
        """
        ### 필요한 열

        필수 열:
        - 제품 또는 제품명
        - 상담내용

        선택 열:
        - 제목
        - 시작 일자/시간
        - 생산팀 (있으면 생산팀별 제품 구분에 활용됩니다)
        """
    )
    st.stop()


# =========================================================
# 파일 읽기
# =========================================================

try:
    df = read_uploaded_file(uploaded_file)
except Exception as error:
    st.error(f"파일을 읽는 중 오류가 발생했습니다: {error}")
    st.stop()

if df.empty:
    st.error("파일에 데이터가 없습니다.")
    st.stop()

df.columns = [str(column).strip() for column in df.columns]


# =========================================================
# 주요 열 찾기
# =========================================================

product_col = find_column(
    df.columns,
    ["제품명_정리", "제품", "제품명", "상품", "상품명"]
)

content_col = find_column(
    df.columns,
    ["상담내용", "상담 내용", "문의내용", "문의 내용", "내용"]
)

title_col = find_column(
    df.columns,
    ["제목", "상담제목", "문의제목"]
)

date_col = find_column(
    df.columns,
    [
        "시작 일자/시간", "시작일자/시간", "상담일시",
        "상담 일시", "상담일자", "상담 일자", "일자"
    ]
)

# ▼▼▼ [추가] 생산팀 열 찾기 ▼▼▼
team_col = find_column(
    df.columns,
    [
        "생산팀", "생산 팀", "생산팀명", "담당생산팀",
        "담당팀", "부서", "생산공장", "공장",
        "생산라인", "생산부서"
    ]
)
# ▲▲▲ [추가] 끝 ▲▲▲


# =========================================================
# 필수 열 확인
# =========================================================

if product_col is None:
    st.error("제품 또는 제품명 열을 찾지 못했습니다.")
    st.write("현재 파일의 열 이름:", list(df.columns))
    st.stop()

if content_col is None:
    st.error("상담내용 열을 찾지 못했습니다.")
    st.write("현재 파일의 열 이름:", list(df.columns))
    st.stop()


# =========================================================
# 데이터 정리
# =========================================================

df[product_col] = df[product_col].fillna("").astype(str).str.strip()
df[content_col] = df[content_col].fillna("").astype(str)

# ▼▼▼ [추가] 생산팀 열 정리 ▼▼▼
if team_col is not None:
    df[team_col] = df[team_col].fillna("미분류").astype(str).str.strip()
    df.loc[df[team_col] == "", team_col] = "미분류"
# ▲▲▲ [추가] 끝 ▲▲▲

df = df[df[product_col] != ""].copy()

if df.empty:
    st.error("제품명이 입력된 데이터가 없습니다.")
    st.stop()


# ▼▼▼ [추가] 제품 ↔ 생산팀 매핑 생성 ▼▼▼
# 한 제품이 여러 생산팀 레코드에 걸쳐 있을 수 있으므로,
# 제품별로 등장한 생산팀을 모두 모아 "/"로 묶어서 표시용 라벨로 사용합니다.
product_team_map = {}

if team_col is not None:
    grouped_teams = (
        df.groupby(product_col)[team_col]
        .apply(lambda series: sorted(set(series)))
    )

    for product, teams in grouped_teams.items():
        product_team_map[product] = "/".join(teams) if teams else "미분류"
# ▲▲▲ [추가] 끝 ▲▲▲


# =========================================================
# 사이드바 제품 검색 (생산팀별 구분 적용)
# =========================================================

with st.sidebar:

    st.header("🔎 제품 검색")

    all_products = sorted(
        df[product_col].dropna().astype(str).unique().tolist()
    )

    # ▼▼▼ [추가] 생산팀 필터 ▼▼▼
    selected_teams = None

    if team_col is not None:
        all_teams = sorted(df[team_col].dropna().astype(str).unique().tolist())

        selected_teams = st.multiselect(
            "생산팀 필터",
            options=all_teams,
            default=[],
            placeholder="전체 생산팀 (선택 안 하면 전체 조회)"
        )

        if selected_teams:
            team_filtered_products = sorted(
                df.loc[df[team_col].isin(selected_teams), product_col]
                .dropna().astype(str).unique().tolist()
            )
        else:
            team_filtered_products = all_products

        st.caption(
            f"생산팀 필터 적용 대상 제품: {len(team_filtered_products):,}개"
        )
    else:
        team_filtered_products = all_products
        st.caption("생산팀 열이 없어 전체 제품을 대상으로 검색합니다.")
    # ▲▲▲ [추가] 끝 ▲▲▲

    st.caption("제품명의 일부만 입력해도 검색됩니다.")

    product_search = st.text_input("제품명 검색", placeholder="예: 아일")
    search_keyword = product_search.strip()

    if search_keyword:

        searched_products = [
            product for product in team_filtered_products
            if search_keyword.casefold() in product.casefold()
        ]

        if searched_products:

            # ▼▼▼ [추가] 생산팀별로 정렬 + [생산팀] 제품명 라벨 구성 ▼▼▼
            if team_col is not None:
                searched_products = sorted(
                    searched_products,
                    key=lambda product: (
                        product_team_map.get(product, "미분류"),
                        product
                    )
                )

                label_to_product = {
                    f"[{product_team_map.get(product, '미분류')}] {product}": product
                    for product in searched_products
                }

                selected_label = st.selectbox(
                    "제품을 선택하세요 (생산팀별 구분)",
                    list(label_to_product.keys()),
                    index=0,
                    label_visibility="collapsed"
                )

                selected_product = label_to_product[selected_label]
            else:
                selected_product = st.selectbox(
                    "제품을 선택하세요",
                    searched_products,
                    index=0,
                    label_visibility="collapsed"
                )
            # ▲▲▲ [추가] 끝 ▲▲▲

        else:
            st.warning("검색 결과가 없습니다.")
            selected_product = None

    else:
        st.info("제품명을 입력하면 자동완성 목록이 표시됩니다.")
        selected_product = None


if selected_product is None:
    st.info("제품명 검색창에 제품명을 입력해 주세요.")
    st.stop()


# =========================================================
# 제품 필터링
# =========================================================

filtered_df = df[df[product_col] == selected_product].copy()


# =========================================================
# 분석 설정
# =========================================================

with st.sidebar:

    st.markdown("---")
    st.subheader("분석 설정")

    min_word_length = st.slider("최소 단어 길이", min_value=1, max_value=5, value=2)
    top_n = st.slider("상위 키워드 개수", min_value=10, max_value=100, value=30)

    custom_stopwords_text = st.text_area(
        "추가 불용어",
        placeholder="예: 양념, 제품명, 문의"
    )


stopwords = set(DEFAULT_STOPWORDS)

if custom_stopwords_text.strip():
    extra_words = custom_stopwords_text.replace("\n", ",").split(",")
    stopwords.update(word.strip() for word in extra_words if word.strip())


# =========================================================
# 요약 정보
# =========================================================

st.subheader(f"선택 제품: {selected_product}")

# ▼▼▼ [추가] 생산팀 열 존재 여부에 따라 메트릭 컬럼 수 조정 ▼▼▼
if team_col is not None:
    metric1, metric2, metric3, metric4 = st.columns(4)
else:
    metric1, metric2, metric3 = st.columns(3)
# ▲▲▲ [추가] 끝 ▲▲▲

with metric1:
    st.metric("선택 제품 상담 건수", f"{len(filtered_df):,}건")

with metric2:
    st.metric("전체 상담 건수", f"{len(df):,}건")

with metric3:
    st.metric("전체 제품 종류", f"{df[product_col].nunique():,}개")

# ▼▼▼ [추가] 선택 제품의 생산팀 표시 ▼▼▼
if team_col is not None:
    with metric4:
        st.metric("소속 생산팀", product_team_map.get(selected_product, "미분류"))
# ▲▲▲ [추가] 끝 ▲▲▲


if filtered_df.empty:
    st.warning("선택한 제품의 상담내용이 없습니다.")
    st.stop()


# =========================================================
# 키워드 추출
# =========================================================

all_words = []

for content in filtered_df[content_col]:
    words = extract_nouns(content, stopwords, min_length=min_word_length)
    all_words.extend(words)

word_counts = Counter(all_words)

if not word_counts:
    st.warning("분석할 키워드가 없습니다. 최소 단어 길이나 불용어를 조정해 보세요.")
    st.stop()

count_df = pd.DataFrame(
    word_counts.most_common(top_n),
    columns=["키워드", "빈도"]
)


# =========================================================
# 폰트 확인 (워드클라우드 전용)
# =========================================================

font_path = get_korean_font()

if font_path is None:
    st.error(
        "한글 폰트를 찾지 못했습니다. "
        "프로젝트 폴더에 fonts/NanumGothic.ttf 파일이 있는지 확인해주세요."
    )
    st.stop()


# =========================================================
# 워드클라우드 생성
# =========================================================

wordcloud = WordCloud(
    font_path=font_path,
    width=1200,
    height=700,
    background_color="white",
    colormap="viridis",
    max_words=100,
    collocations=False,
    prefer_horizontal=0.8
).generate_from_frequencies(word_counts)


# =========================================================
# 워드클라우드 + 상위 키워드 표시
# =========================================================

left_col, right_col = st.columns([1.5, 1])

with left_col:

    st.subheader("워드클라우드")

    fig_wc, ax_wc = plt.subplots(figsize=(12, 7))
    ax_wc.imshow(wordcloud, interpolation="bilinear")
    ax_wc.axis("off")

    st.pyplot(fig_wc, use_container_width=True)
    plt.close(fig_wc)

with right_col:

    st.subheader("상위 키워드")

    st.dataframe(
        count_df,
        use_container_width=True,
        hide_index=True,
        height=700
    )


# =========================================================
# 상담 원문
# =========================================================

st.subheader("관련 상담내용")

display_columns = []

if date_col is not None:
    display_columns.append(date_col)

if title_col is not None:
    display_columns.append(title_col)

# ▼▼▼ [추가] 생산팀 열도 원문 테이블에 함께 표시 ▼▼▼
if team_col is not None:
    display_columns.append(team_col)
# ▲▲▲ [추가] 끝 ▲▲▲

display_columns.append(product_col)
display_columns.append(content_col)

display_columns = list(dict.fromkeys(display_columns))

display_df = filtered_df[display_columns].copy()

consultation_search = st.text_input(
    "상담내용 검색",
    placeholder="예: 맛, 포장, 배송, 이물, 유통기한"
)

if consultation_search.strip():

    search_mask = display_df.astype(str).apply(
        lambda column: column.str.contains(
            consultation_search, case=False, na=False, regex=False
        )
    ).any(axis=1)

    display_df = display_df[search_mask]

st.write(f"조회 결과: {len(display_df):,}건")

st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# 다운로드
# =========================================================

st.subheader("분석 결과 다운로드")

csv_data = count_df.to_csv(index=False, encoding="utf-8-sig")

st.download_button(
    label="키워드 빈도 CSV 다운로드",
    data=csv_data,
    file_name="키워드_빈도.csv",
    mime="text/csv"
)
