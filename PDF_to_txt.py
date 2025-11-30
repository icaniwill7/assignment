def main():
    st.set_page_config(
        page_title="디지털말뭉치 분석기 (PDF → DOCX + MATTR/LexSoph)",
        page_icon="📚",
        layout="wide",
    )

    # 🔹 커스텀 스타일 적용
    set_custom_page_style()

    # ---------- 상단 헤더 (로고 + 제목) ---------- #
    header_col1, header_col2 = st.columns([1, 4])

    with header_col1:
        try:
            st.image("yonsei_logo.png", width=90)
        except Exception:
            st.markdown("### 🏫")

    with header_col2:
        st.markdown(
            """
            <h1>디지털말뭉치 분석과 언어교육 분석기</h1>
            <p style="font-size:16px; color:#4b5563; margin-top:0;">
            Yonsei University · English Language & Literature · Digital Corpus Linguistics
            </p>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ---------- 사이드바: window size 설정 ---------- #
    st.sidebar.header("⚙️ Window size 설정")
    win_all = st.sidebar.number_input(
        "All words window size", min_value=5, max_value=500, value=50, step=1
    )
    win_pos = st.sidebar.number_input(
        "POS window size", min_value=5, max_value=200, value=11, step=1
    )

    st.sidebar.markdown("---")
    st.sidebar.info("여러 개의 PDF를 한 번에 업로드할 수 있습니다.")

    # ---------- 메인: 설명 + 업로더 + 버튼을 카드 안에 ---------- #
    with st.container():
        st.markdown(
            """
            <div class="yonsei-card">
                <h3>📑 분석 개요</h3>
                <ul>
                    <li><b>PDF 파일</b>을 업로드하면, 각 파일을 Word(DOCX)로 변환합니다.</li>
                    <li>PDF에서 추출한 텍스트를 이용해 <b>MATTR + Lexical sophistication</b>을 계산합니다.</li>
                    <li>윈도우 크기, POS, AWL, n-gram 비율은 <b>원래 Tkinter 버전과 동일한 로직</b>을 사용합니다.</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.write("")  # 살짝 여백

    with st.container():
        st.markdown('<div class="yonsei-card">', unsafe_allow_html=True)

        st.markdown("#### 📂 분석할 PDF 파일 업로드")
        uploaded_files = st.file_uploader(
            "Drag & Drop 또는 [Browse files] 버튼으로 업로드하세요.",
            type=["pdf"],
            accept_multiple_files=True,
        )

        st.write("")
        start_col1, start_col2 = st.columns([1, 3])
        with start_col1:
            start_button = st.button("🚀 PDF 변환 + 분석 시작", use_container_width=True)
        with start_col2:
            if uploaded_files:
                st.success(f"{len(uploaded_files)}개의 파일이 업로드되었습니다.")
            else:
                st.info("현재 업로드된 파일이 없습니다.")

        st.markdown("</div>", unsafe_allow_html=True)

    # ---------- 버튼이 눌렸을 때만 분석 ---------- #
    if start_button:
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

        # ---------- 결과 카드 ---------- #
        st.markdown(
            """
            <div class="yonsei-card">
                <h3>📊 분석 결과 (MATTR + Lexical sophistication)</h3>
            """,
            unsafe_allow_html=True,
        )

        df_results = pd.DataFrame(results)
        st.dataframe(df_results, use_container_width=True)

        csv_buf = BytesIO()
        df_results.to_csv(csv_buf, index=False, encoding="utf-8-sig")
        csv_buf.seek(0)
        st.download_button(
            label="📥 결과 CSV 다운로드 (results.csv)",
            data=csv_buf,
            file_name="results.csv",
            mime="text/csv",
        )

        st.markdown("</div>", unsafe_allow_html=True)

        # ---------- DOCX 다운로드 카드 ---------- #
        st.write("")
        st.markdown(
            """
            <div class="yonsei-card">
                <h3>📄 변환된 DOCX 파일 다운로드</h3>
            """,
            unsafe_allow_html=True,
        )

        for name, b in docx_downloads:
            base = os.path.splitext(os.path.basename(name))[0]
            st.download_button(
                label=f"📄 {base}.docx 다운로드",
                data=b,
                file_name=f"{base}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
            )

        st.markdown("</div>", unsafe_allow_html=True)
