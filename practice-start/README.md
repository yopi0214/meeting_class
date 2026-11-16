# 연구 관리 워크스페이스 (Word·Excel·PPT 버전)

대학원생이 자기 연구의 자료, 주간 기록, 지도교수 미팅을 Claude Code로 관리하는 워크스페이스입니다. 자료는 Word, Excel, PPT로 넣고, 결과물도 Word, PPT로 받습니다. 폴더 안의 자료는 모두 가상입니다(첫 줄에 `예시`).

## 지금 어떤 상태인가

이 폴더는 **3주차가 시작된 때(2026-11-16)** 입니다.

- **1주차(11-02 ~ 11-08):** 첫 주간 기록을 썼고, 프로젝트 2개(p53-ko-cell, rnaseq-drug-response)의 상태를 정리했습니다.
- **2주차(11-09 ~ 11-15):** 주간 기록, 지도교수 미팅(11-10, 메모·정리·발표 PPT), qPCR·DEG 결과 자료가 쌓였습니다.
- **3주차 시작:** `inbox/`에 정리하지 않은 새 자료 3개(qPCR 결과 Excel, 업체 메일, 세미나 공지 Word)가 들어와 있습니다.

## 파일 형식

| 무엇 | 형식 | 누가 만드나 |
|---|---|---|
| 원본 자료(데이터, 메일, 공지) | `.xlsx`, `.eml`, `.docx` | 학생 (`inbox/`에 넣음) |
| 주간 기록 `weekly/날짜.docx` | Word | `/weekly` |
| 미팅 준비 `meetings/…/brief.docx` | Word | `/meeting-prep` |
| 미팅 메모 `meetings/…/notes.docx` | Word | 학생이 직접 씀 |
| 미팅 정리 `meetings/…/summary.docx` | Word | `/meeting-close` |
| 발표 `meetings/…/slides.pptx` | PowerPoint | `/progress-ppt` |
| 프로젝트 상태·할 일·일정 (`status.md`, `todo.md` 등) | 텍스트 | Claude가 승인을 받고 갱신 |

Claude Code는 `tools/`의 스크립트로 Word, Excel, PPT를 읽고 씁니다. 설치할 것은 없고 Windows PowerShell만 있으면 됩니다.

## 명령

| 명령 | 언제 | 하는 일 |
|---|---|---|
| `/intake` | 새 자료를 `inbox/`에 넣은 뒤 | 자료를 읽고 프로젝트·접수일별로 `sources/`에 옮기며 출처를 기록 |
| `/weekly [날짜]` | 매주 한 번 | 지난 기록과 새 자료를 비교해 주간 기록(Word, 2쪽 이내) 작성, 프로젝트 상태·할 일 갱신 제안 |
| `/meeting-prep <종류> [날짜]` | 미팅 전날 | 지난 미팅 이후 달라진 점, 피드백 처리 상태, 여쭤볼 질문을 Word(2쪽 이내)로 정리 |
| `/progress-ppt [종류] [날짜]` | 발표 전날 | 지난 발표 PPT를 복사해 새 데이터와 주간 기록으로 고친 이번 진도점검 PPT 작성 |
| `/meeting-close [폴더]` | 미팅 직후 | 내 메모(Word)를 결정·피드백·할 일로 정리(Word)하고 할 일 갱신 제안 |
| `/paper-note <PDF 또는 URL>` | 논문을 읽을 때 | 읽기 노트(Word) 작성, 내 프로젝트와의 연결 정리 |

## 3주차 실습

순서대로 해 보고, 끝날 때마다 폴더를 열어 무엇이 바뀌었는지 봅니다.

| 단계 | 입력할 것 | 결과 |
|---|---|---|
| 1. 자료 정리 | `/intake` | `inbox/`의 3개가 `sources/`로 이동하고 `sources.md`에 기록됨. 세미나 공지는 `sources/_기타/`로 |
| 2. 주간 기록 | `/weekly 2026-11-16` | `weekly/2026-11-16.docx` 생성. 승인하면 `status.md`와 `todo.md` 갱신 |
| 3. 미팅 준비 | `/meeting-prep 교수미팅 2026-11-17` | `meetings/2026-11-17_교수미팅/brief.docx`와 빈 `notes.docx` 생성 |
| 4. 진도점검 PPT | `/progress-ppt 교수미팅 2026-11-17` | 지난 발표를 복사해 고친 `slides.pptx` 생성 (지난 발표는 그대로) |
| 5. 미팅 후 정리 | `notes.docx`에 미팅 메모를 적은 뒤 `/meeting-close 2026-11-17_교수미팅` | `summary.docx` 생성. 승인하면 `todo.md` 갱신 |
| 6. 논문 노트 | `/paper-note <오픈 액세스 논문 URL 또는 PDF>` | `reading/`에 Word 노트 |

날짜가 시나리오(2026-11)보다 이르거나 늦어도 오류가 아닙니다.

## 주의

- Claude가 만든 Word, PPT는 반드시 직접 열어 확인하세요. 숫자와 날짜가 근거 파일과 맞는지, 이미지 자리가 비어 있지 않은지 봅니다.
- 그림, 블롯 이미지, 차트는 Claude가 내용을 읽지 못합니다. 내용을 아는 사람이 직접 적어 줘야 합니다.
- `.hwp`(한글) 파일은 읽지 못합니다. Word나 PDF로 저장해서 넣으세요.
