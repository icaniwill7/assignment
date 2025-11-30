# app.py
import os
import io
import csv
from datetime import datetime

import streamlit as st
import numpy as np
import pandas as pd
from nltk import word_tokenize, pos_tag, download
from docx import Document
from PyPDF2 import PdfReader
from nltk import word_tokenize, pos_tag, download

# ---------- NLTK 리소스 다운로드 ---------- #
download("punkt", quiet=True)
download("punkt_tab", quiet=True)          # 🔹 이 줄 추가
download("averaged_perceptron_tagger", quiet=True)

# ---------- 상수 설정 ---------- #
POS_CATEGORIES = {
    "Verb": "VB",
    "Noun": "NN",
    "Adjective": "JJ",
    "Adverb": "RB",
}

ACADEMIC_WORDS = {
    "analyze", "approach", "area", "assess", "assume", "authority", "concept",
    "consistent", "constitute", "context", "contract", "create", "data",
    "definition", "derive", "distribute", "economy", "environment",
    "establish", "estimate", "evidence", "export", "factor", "formula",
    "function", "identify", "income", "indicate", "interpret", "involve",
    "issue", "legal", "major", "method", "occur", "percent", "policy",
    "principle", "process", "require", "research", "response", "role",
    "section", "sector", "significant", "similar", "source", "specific",
    "structure", "theory", "vary",
}

# ---------- 분석용 함수 ---------- #
def extract_pos(text: str, prefix: str):
    """특정 품사(POS prefix)에 해당하는 단어 리스트 추출"""
    tokens = word_tokenize(text)
    tagged = pos_tag(tokens)
    return [w.lower() for w, tag in tagged if tag.startswith(prefix)]


def calculate_mattr(words, window_size=50):
    """전체 단어 리스트에 대한 MATTR 계산"""
    if not words:
        return 0.0
    n = len(words)
    if n < window_size:
        return len(set(words)) / n
    ratios = []
    for i in range(n - window_size + 1):
        window = words[i : i + window_size]
        ratios.append(len(set(window)) / window_size)
    return float(np.mean(ratios))


def calculate_category_mattr(category_words, all_words, window_size=11):
    """전체 단어 시퀀스 안에서 특정 category 단어들의 MATTR 비슷하게 계산"""
    if not all_words:
        return 0.0
    n = len(all_words)
    if n < window_size:
        return len(set(category_words)) / len(category_words) if category_words else 0.0

    # 윈도우마다 category 단어만 필터링해서 type/token 비율을 계산
    ratios = []
    cat_set = set(category_words)
    for i in range(n - window_size + 1):
        window = all_words[i : i + window_size]
        hits = [w for w in window if w in cat_set]
        if hits:
            ratios.append(len(set(hits)) / window_size)
    return float(np.mean(ratios)) if ratios else 0.0


def calc_lexical_soph(all_words):
    """AWL 비율, bigram type 비율, trigram type 비율"""
    if not all_words:
        return 0.0, 0.0, 0.0

    total = len(all_words)
    awl_ratio = sum(1 for w in all_words if w in ACADEMIC_WORDS) / total

    bigrams = ["_".join(all_words[i : i + 2]) for i in range(len(all_words) - 1)]
    trigrams = ["_".join(all_words[i : i + 3]) for i in range(len(all_words) - 2)]

    bigram_ratio = len(set(bigrams)) / len(bigrams) if bigrams else 0.0
    trigram_ratio = len(set(trigrams)) / len(trigrams) if trigrams else 0.0

    return round(awl_ratio, 4), round(bigram_ratio, 4), round(trigram_ratio, 4)


# ---------- PDF → DOCX + 텍스트 추출 ---------- #
def pdf_to_docx_and_text(uploaded_file):
    """Streamlit 업로드된 PDF 파일 1개를 DOCX + 전체 텍스트로 변환"""
    reader = PdfReader(uploaded_file)
    doc = Document()

    all_text_parts = []

    for page in reader.pages:
        text = page.extract_text() or ""
        all_text_parts.append(text)
        for line in text.splitlines():
            doc.add_paragraph(line)
        doc.add_page_break()

    full_text = "\n".join(all_text_parts)

    # DOCX를 메모리 버퍼에 저장
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    return buf, full_text


# ---------- 텍스트 분석 ---------- #
def analyze_text(filename, text, win_all, win_pos):
    tokens = word_tokenize(text)
    all_words = [w.lower() for w in tokens if w.isalpha()]

    row = {"Filename": filename}
    row["All_words_MATTR"] = round(calculate_mattr(all_words, win_all), 4)

    for label, prefix in POS_CATEGORIES.items():
        cat_words = extract_pos(text, prefix)
        row[f"{label}_MATTR"] = round(
            calculate_category_mattr(cat_words, all_words, win_pos), 4
        )

    awl, big, tri = calc_lexical_soph(all_words)
    row["LexSoph_AWLratio"] = awl
    row["LexSoph_BigramRatio"] = big
    row["LexSoph_TrigramRatio"] = tri

    return row


# ---------- Streamlit 메인 ---------- #
def main():
    st.set_page_config(
        page_title="디지털말뭉치 분석기 (PDF → DOCX + MATTR/LexSoph)",
        page_icon="📚",
        layout="wide",
    )

    # ---- 상단 제목 영역 ---- #
    st.title("디지털말뭉치 분석과 언어교육 분석기 (Streamlit 버전)")
    st.caption("PDF → DOCX 변환 + MATTR, Lexical sophistication 분석 자동화 도구")

    st.markdown(
        """
        **사용 방법**
        1. 왼쪽 사이드바에서 *window size*를 설정합니다.  
        2. 아래 영역에 분석할 **PDF 파일**들을 업로드합니다.  
        3. `PDF 변환 + 분석 시작` 버튼을 누르면,  
           - 각 PDF가 Word(DOCX)로 변환되고  
           - 텍스트를 이용해 MATTR + Lexical sophistication 지표가 계산됩니다.
        """,
    )

    st.divider()

    # ---- 사이드바: window size ---- #
    st.sidebar.header("⚙️ Window size 설정")
    win_all = st.sidebar.number_input(
        "All words window size", min_value=5, max_value=500, value=50, step=1
    )
    win_pos = st.sidebar.number_input(
        "POS window size", min_value=5, max_value=200, value=11, step=1
    )
    st.sidebar.info("※ 수업 예제 기준: All words=50, POS=11")

    # ---- 파일 업로드 영역 ---- #
    st.subheader("📂 분석할 PDF 파일 업로드")
    uploaded_files = st.file_uploader(
        "여러 개의 PDF를 한 번에 업로드할 수 있습니다.",
        type=["pdf"],
        accept_multiple_files=True,
    )

    col_btn, col_msg = st.columns([1, 3])
    with col_btn:
        start = st.button("🚀 PDF 변환 + 분석 시작", use_container_width=True)
    with col_msg:
        if uploaded_files:
            st.success(f"{len(uploaded_files)}개의 PDF가 업로드되었습니다.")
        else:
            st.info("현재 업로드된 파일이 없습니다.")

    if not start:
        return

    if not uploaded_files:
        st.warning("먼저 PDF 파일을 하나 이상 업로드하세요.")
        return

    # ---- 실제 처리 ---- #
    results = []
    docx_downloads = []

    progress = st.progress(0)
    status = st.empty()
    total = len(uploaded_files)

    for i, up in enumerate(uploaded_files, start=1):
        status.text(f"{i}/{total} 처리 중…  ({up.name})")
        try:
            docx_bytes, text = pdf_to_docx_and_text(up)
            row = analyze_text(up.name, text, int(win_all), int(win_pos))
            results.append(row)

            docx_downloads.append((up.name, docx_bytes))
        except Exception as e:
            st.error(f"❌ {up.name} 처리 중 오류: {e}")

        progress.progress(i / total)

    status.text("모든 파일 처리 완료!")

    if not results:
        st.error("유효한 결과가 없습니다.")
        return

    # ---- 결과 테이블 ---- #
    st.subheader("📊 분석 결과 (MATTR + Lexical sophistication)")
    df = pd.DataFrame(results)
    st.dataframe(df, use_container_width=True)

    # CSV 다운로드
    csv_buf = io.StringIO()
    df.to_csv(csv_buf, index=False, encoding="utf-8-sig")
    csv_bytes = csv_buf.getvalue().encode("utf-8-sig")

    st.download_button(
        "📥 결과 CSV 다운로드 (results.csv)",
        data=csv_bytes,
        file_name="results.csv",
        mime="text/csv",
    )

    # ---- DOCX 다운로드 ---- #
    st.divider()
    st.subheader("📄 변환된 DOCX 파일 다운로드")
    for name, buf in docx_downloads:
        base = os.path.splitext(os.path.basename(name))[0]
        st.download_button(
            label=f"📄 {base}.docx 다운로드",
            data=buf,
            file_name=f"{base}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
        )


if __name__ == "__main__":
    main()
