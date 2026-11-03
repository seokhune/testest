HWP 문서 검수 워크플로우
========================

Google Drive에 업로드된 한컴(HWP/HWPX) 문서를 검수하는 Mistral Workflow입니다.

동작 방식
---------

1. Drive에서 원본 파일을 찾아 **사본을 생성**합니다 (원본은 수정하지 않음)
2. HWP/HWPX → DOCX 변환 (LibreOffice + H2Orestart)
3. 본문 오탈자 수정 → 주석 윤문(책 제목 〈 〉 표기 포함) → 주석 출처 표기 → 한문 주석 번역("번역:" 라벨, 원문 아래 배치) → 본문-주석 차이 비교
4. 결과물 업로드: `{원본}_교정본.docx`, `{원본}_교정리포트.docx`

입력
----

- `file_name`: Drive에서 검색할 파일명 (.hwp / .hwpx / .docx; .docx는 변환 없이 그대로 사용)
- `output_parent_id`: (선택) 결과물을 저장할 Drive 폴더 ID
- `page_start`, `page_end`: (선택) 검수 범위를 페이지 단위로 제한 (예: 8~15페이지만 처리). 지정하면 결과 파일명에 `_p{start}-{end}` 붙습니다

구조
----

- `src/worker/workflows/hwp_document_review.py` — 워크플로우 + LLM/Drive 액티비티
- `src/worker/workflows/hwp_review_docx.py` — DOCX 추출/편집/리포트 생성 + HWP 변환
- `src/worker/workflows/hwp_review_models.py` — 입출력/편집 모델
- `tests/` — 단위 + E2E 테스트
