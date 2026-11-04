# PDF 문서 검수 워크플로우

Google Drive에 업로드된 PDF를 검수하는 Mistral Workflow입니다.

## 동작 방식
1. Drive에서 PDF 검색
2. 텍스트 추출 도구(`read_file_content`)로 본문 읽기 — PDF 바이너리를
   다운로드하지 않아 커넥터의 base64 2MB 전송 한도보다 큰 파일도 처리 가능
3. 페이지 표식(`- 8 -` 등)으로 실제 페이지 번호를 파악하고 `page_start`/`page_end`
   범위의 문단만 슬라이싱
4. LLM 교정(오타/문법/띄어쓰기) + 한문 구절 번역
5. 결과물 DOCX 업로드: `{원본}_p{start}-{end}_교정본.docx`, `{원본}_p{start}-{end}_검수리포트.docx`

## 입력
- `file_name`: Drive에서 검색할 PDF 파일명
- `page_start`, `page_end`: (선택) 검수할 페이지 범위
- `output_parent_id`: (선택) 결과물을 저장할 Drive 폴더 ID

## 구조
- `src/worker/workflows/pdf_document_review.py` — 워크플로우 + LLM/Drive 액티비티
- `src/worker/workflows/hwp_review_docx.py` — DOCX 파싱/편집/리포트 헬퍼 라이브러리
- `src/worker/workflows/hwp_review_models.py` — 입출력/편집 모델
- `tests/` — 단위 + E2E 테스트
