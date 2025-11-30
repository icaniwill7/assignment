# -*- coding: utf-8 -*-
import os
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from PyPDF2 import PdfReader
from docx import Document

# 🔥 NLTK는 반드시 최상단에서 import
import nltk
from nltk import word_tokenize, pos_tag

# 🔥 NLTK 리소스 다운로드 (requirements에서 nltk==3.8.1이면 이 두 개면 충분)
nltk.download("punkt", quiet=True)
nltk.download("averaged_perceptron_tagger", quiet=True)

# ---------- 기본 설정 (원본 Tk 코드와 동일) ---------- #
POS_CATEGORIES = {
    "Verb": "VB",
    "Noun": "NN",
    "Adjective": "JJ",
    "Adverb": "RB",
}

ACADEMIC_WORDS = {
    "analyze", "approach", "area", "assess", "assume", "authority", "concept",
    "consistent", "constitute", "context", "contract", "create", "data",
    "definition", "derive", "distribute", "economy", "environment", "establish",
    "estimate", "evidence", "export", "factor", "formula", "function", "identify",
    "income", "indicate", "interpret", "involve", "issue", "legal", "major",
    "method", "occur", "percent", "policy", "principle", "process", "require",
    "research", "response", "role", "section", "sector", "significant", "similar",
    "source", "specific", "structure", "theory", "vary"
}


# ---------- PDF → DOCX + 텍스트 추출 ---------- #
def pdf_to_docx_and_text(pdf_file) -> tuple[BytesIO, str]:
    """
    Streamlit 업로드 객체 (pdf_file)를 받아서
    1) 페이지별 텍스트를 추출해 DOCX로 변환하고
    2) 전체 텍스트를 하나의 문자열로 반환한다.
    (원래 pdf_to_docx_simple과 거의 동일한 로직)
    """
    reader = PdfReader(pdf_file)
    doc = Document()
    all_text_parts = []

    pages = reader.pages
    num_pages = len(pages)

    for i, page in enumerate(pages):
        text = page.extract_text()
        if text:
            all_text_parts.append(text)
            # 줄 단위로 paragraph 추가 (원본 코드와 동일한 구조)
            for line in text.splitlines():
                doc.add_paragraph(line)
        # 마지막 페이지가 아니면 page break 추가
        if i < num_pages - 1:
            doc.add_page_break()

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    full_text = "\n".join(all_text_parts)
    return buffer, full_text


# ---------- 분석 함수들 (원본 Tk 코드 알고리즘 그대로) ---------- #
def extract_pos(text: str, prefix: str):
    tagged = pos_tag(word_tokenize(text))
    return [w.lower() for w, tag in tagged if tag.startswith(prefix)]


def calculate_mattr(words, win: int = 50) -> float:
    if not words:
        return 0.0
    if len(words) < win:
        return len(set(words)) / len(words)
    return float(
        np.mean(
            [len(set(words[i:i + win])) / win for i in range(len(words) - win + 1)]
        )
    )


def calculate_category_mattr(cat, allw, win: int = 11) -> float:
    if len(allw) < win:
        return len(set(cat)) / len(cat) if cat else 0.0
    vals = []
    for i in range(len(allw) - win + 1):
        window = allw[i:i + win]
        hits = [w for w in window if w in cat]
        if hits:
            vals.append(len(set(hits)) / win)
    return float(np.mean(vals)) if vals else 0.0


def calc_lexical_soph(allw):
    """
    AWL 비율 + bigram/trigram type-token ratio
    (원본 코드 calc_lexical_soph와 동일한 아이디어)
    """
    if not allw:
        return 0.0, 0.0, 0.0

    total = len(allw)
    awl = sum(1 for w in allw if w in ACADEMIC_WORDS) / total

    bigr = ["_".join(allw[i:i + 2]) for i in range(len(allw) - 1)]
    trigr = ["_".join(allw[i:i + 3]) for i in range(len(allw) - 2)]

    big = len(set(bigr)) / len(bigr) if bigr else 0.0
    tri = len(set(trigr)) / len(trigr) if trigr else 0.0

    return round(awl, 4), round(big, 4), round(tri, 4)


def analyze_text(filename: str, text: str, win_all: int, win_pos: int) -> dict:
    """
    한 파일(텍스트)에 대해:
    - All_words_MATTR
    - POS별 MATTR
    - Lexical Sophistication
    계산해서 dict로 반환
    """
    tokens = word_tokenize(text)
    allw = [w.lower() for w in tokens if w.isalpha()]

    row = {
        "Filename": filename,
        "All_words_MATTR": round(calculate_mattr(allw, win_all), 4),
    }

    for lbl, pref in POS_CATEGORIES.items():
        cat = extract_pos(text, pref)
        row[f"{lbl}_MATTR"] = round(
            calculate_category_mattr(cat, allw, win_pos), 4
        )

    awl, big, tri = calc_lexical_soph(allw)
    row["LexSoph_AWLratio"] = awl
    row["LexSoph_BigramRatio"] = big
    row["LexSoph_TrigramRatio"] = tri

    return row


# ---------- Streamlit 메인 앱 ---------- #
def main():
    st.set_page_config(
        page_title="디지털말뭉치 분석기 (PDF → DOCX + MATTR/LexSoph)",
        page_icon="📚",
        layout="wide",
    )

    st.title("📚 디지털말뭉치 분석과 언어교육 분석기 (Streamlit 버전)")
    st.markdown(
        """
        1. **PDF 파일들을 업로드**하면, 각 파일을 Word(DOCX)로 변환합니다.  
        2. 동시에 PDF에서 추출한 텍스트로 **MATTR + Lexical sophistication** 분석을 수행합니다.  
        3. 원래 Tkinter + ttkbootstrap 버전과 **동일한 계산 로직(window size, POS, AWL, n-gram 비율)**을 사용합니다.
        """
    )

    # 사이드바: window size 설정 (원래 기본값 50 / 11 그대로)
    st.sidebar.header("Window size 설정")
    win_all = st.sidebar.number_input(
        "All words window size", min_value=5, max_value=500, value=50, step=1
    )
    win_pos = st.sidebar.number_input(
        "POS window size", min_value=5, max_value=200, value=11, step=1
    )

    st.sidebar.markdown("---")
    st.sidebar.info("여러 개의 PDF를 한 번에 업로드할 수 있습니다.")

    uploaded_files = st.file_uploader(
        "분석할 PDF 파일을 선택하세요 (복수 선택 가능)",
        type=["pdf"],
        accept_multiple_files=True,
    )

    if st.button("PDF 변환 + 분석 시작"):
        if not uploaded_files:
            st.warning("먼저 PDF 파일을 하나 이상 업로드하세요.")
            st.stop()

        results = []
        docx_downloads = []

        progress = st.progress(0)
        status_text = st.empty()

        total = len(uploaded_files)

        for idx, up in enumerate(uploaded_files, start=1):
            status_text.text(f"{idx}/{total} 처리 중: {up.name}")
            try:
                # PDF → DOCX + 텍스트 추출
                docx_bytes, text = pdf_to_docx_and_text(up)

                # 분석
                row = analyze_text(up.name, text, int(win_all), int(win_pos))
                results.append(row)

                # DOCX 다운로드용 데이터 저장
                docx_downloads.append((up.name, docx_bytes))

            except Exception as e:
                st.error(f"❌ {up.name} 처리 중 오류: {e}")

            progress.progress(idx / total)

        status_text.text("모든 파일 처리 완료!")

        if not results:
            st.error("유효한 결과가 없습니다.")
            st.stop()

        # 결과 테이블 표시
        df_results = pd.DataFrame(results)
        st.subheader("📊 분석 결과 (MATTR + Lexical Sophistication)")
        st.dataframe(df_results, use_container_width=True)

        # CSV 다운로드
        csv_buf = BytesIO()
        df_results.to_csv(csv_buf, index=False, encoding="utf-8-sig")
        csv_buf.seek(0)
        st.download_button(
            label="결과 CSV 다운로드 (results.csv)",
            data=csv_buf,
            file_name="results.csv",
            mime="text/csv",
        )

        st.markdown("---")
        st.subheader("📄 변환된 DOCX 파일 다운로드")

        for name, b in docx_downloads:
            base = os.path.splitext(os.path.basename(name))[0]
            st.download_button(
                label=f"{base}.docx 다운로드",
                data=b,
                file_name=f"{base}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
            )


if __name__ == "__main__":
    main()
