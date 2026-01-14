# Outlook Mail Auto Backup (MSG)

Outlook에서 수신/발신 메일을 자동 및 수동으로 .msg 파일로 백업하는 VBA 매크로입니다.
받은편지함 하위 폴더 이동까지 감지하며,
메일을 하위 폴더로 옮기면 Inbox 루트에 생성된 기존 백업 파일은 자동 삭제되어
항상 최종 정리된 위치의 백업만 남도록 설계되었습니다.

---

## 주요 기능

- Inbox(받은편지함) 루트 + 하위 폴더 전체 모니터링
- Sent(보낸편지함) 자동 백업
- 메일 수동 이동 / 규칙 이동 자동 감지
- 선택한 메일 수동 백업 기능 제공
- 하위 폴더로 이동 시 기존 루트 백업 자동 삭제
- 연도(YYYY) 기준 자동 분리 저장
- 파일명 길이 / 특수문자 안전 처리
- 성공 / 에러 로그 자동 기록
- Outlook 시작 시 자동 활성화

---

## 백업 파일 형식

- 형식: .msg
- 저장 방식:

  mail.SaveAs fullPath, olMSG

MSG 형식을 사용하는 이유:
- Outlook에서 원본 그대로 열림
- 첨부파일, HTML 서식, 메타데이터 100% 보존
- Exchange / Outlook 환경에서 가장 안정적

---

## 저장 경로 구조

D:\backup\mail\
 └─ 2026\
    ├─ 프로젝트A\
    │  └─ 고객사1\
    │     └─ 20260114_123045_홍길동_메일제목.msg
    └─ Sent\
       └─ 20260114_091530_수신자_메일제목.msg

- 연도(YYYY) 기준 자동 분리
- Inbox 하위 폴더 구조 그대로 미러링
- Sent는 항상 YYYY\Sent 아래 저장

---

## 동작 방식 요약

- 메일 수신(Inbox 루트): 백업 1회
- 수신 후 하위 폴더로 이동: 하위 폴더 백업 + 루트 백업 자동 삭제
- 규칙으로 바로 하위 폴더 이동: 하위 폴더에만 백업
- 보낸 메일: Sent 폴더에 백업
- 수동 백업: 현재 메일 위치 기준으로 백업

---

## 파일 구성 (루트 폴더)

/
 ├─ README.md
 ├─ ThisOutlookSession.cls
 ├─ clsFolderItemsHandler.cls
 └─ modMailBackup.bas

---

## 설치 방법

1. Outlook 실행
2. Alt + F11 → VBA 편집기

### 표준 모듈
- 삽입 > 모듈
- 이름: modMailBackup
- modMailBackup.bas 내용 전체 복사 후 붙여넣기

### 클래스 모듈
- 삽입 > 클래스 모듈
- 이름: clsFolderItemsHandler
- clsFolderItemsHandler.cls 내용 전체 복사 후 붙여넣기
- Instancing: Private (기본값 유지)

### ThisOutlookSession
- Microsoft Outlook Objects > ThisOutlookSession
- ThisOutlookSession.cls 내용 전체 복사 후 기존 코드 전부 교체

---

## 매크로 활성화

- Outlook 재시작
또는
- Alt + F8 → InitializeEventHandler 실행

---

## 수동 백업 사용 방법

1. Outlook에서 메일 하나 또는 여러 개 선택
2. Alt + F8
3. SaveSelectedMailsAsMSG 실행

---

## 로그 파일

D:\backup\mail\logs\
 ├─ 2026-01_success.log
 └─ 2026-01_error.log

---

## UserProperty

BackupFilePath
- Outlook 내부 사용자 속성
- 마지막으로 성공 저장된 백업 파일 경로 기록

---

## 백업 경로 변경

Public Const BACKUP_BASE_PATH As String = "D:\backup\mail\"

---

## 라이선스

MIT License
