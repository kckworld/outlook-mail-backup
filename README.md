# Outlook Mail Auto Backup (MSG)

Outlook에서 **수신/발신 메일을 자동으로 `.msg` 파일로 백업**하는 VBA 매크로입니다.  
받은편지함 **하위 폴더 이동까지 감지**하며,  
메일을 하위 폴더로 옮기면 **Inbox 루트에 생성된 백업 파일은 자동 삭제**되어  
최종 정리된 위치의 백업만 남도록 설계되었습니다.

---

## 주요 기능

- Inbox(받은편지함) 하위 폴더 전체 모니터링
- Sent(보낸편지함) 자동 백업
- 메일 수동 이동 / 규칙 이동 자동 감지
- 하위 폴더로 이동 시 기존 루트 백업 자동 삭제
- 연도(YYYY) 기준 자동 분리 저장
- 파일명 길이 / 특수문자 안전 처리
- 성공 / 에러 로그 자동 기록
- Outlook 시작 시 자동 활성화

---

## 백업 파일 형식

- 형식: `.msg`
- 저장 방식:
  
  mail.SaveAs fullPath, olMSG

### MSG 형식을 사용하는 이유

- Outlook에서 원본 그대로 열림
- 첨부파일, HTML 서식, 메타데이터 100% 보존
- Exchange / Outlook 환경에서 가장 안정적

※ `.eml` 형식은 사용하지 않습니다.

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
- 수신 후 하위 폴더로 이동: 하위 폴더에 백업 + 루트 백업 자동 삭제
- 규칙으로 바로 하위 폴더 이동: 하위 폴더에만 백업
- 보낸 메일: Sent 폴더에 백업

---

## 설치 방법

### 1. 코드 다운로드

git clone https://github.com/YOUR_ID/outlook-mail-backup.git

---

### 2. Outlook VBA에 코드 등록

1. Outlook 실행
2. Alt + F11 → VBA 편집기

(1) 표준 모듈  
- 삽입 > 모듈  
- 이름: modMailBackup  
- modMailBackup.bas 내용 붙여넣기  

(2) 클래스 모듈  
- 삽입 > 클래스 모듈  
- 이름: clsFolderItemsHandler  
- Instancing = Private (기본값 유지)  
- clsFolderItemsHandler.cls 내용 붙여넣기  

(3) ThisOutlookSession  
- Microsoft Outlook Objects > ThisOutlookSession  
- ThisOutlookSession.cls 내용 전체 교체  

---

### 3. 매크로 활성화

- Outlook 재시작  
또는  
- Alt + F8 → InitializeEventHandler 실행  

---

## 로그 파일

D:\backup\mail\logs\
 ├─ 2026-01_success.log  
 └─ 2026-01_error.log  

- 성공 로그: 저장 경로, 파일 크기, 발신자, 제목
- 에러 로그: 오류 메시지 + 메일 정보

---

## 중요 참고 사항

### 새 하위 폴더 생성 시
받은편지함 하위에 새 폴더를 추가한 경우  
InitializeEventHandler를 다시 실행해야 모니터링됩니다.

---

### UserProperty 사용
중복 백업 정리를 위해 메일에 다음 UserProperty를 사용합니다.

BackupFilePath

- Outlook 내부 사용자 속성
- 메일 본문/헤더에는 영향 없음

---

### 매크로 보안
- Outlook 매크로 보안 정책에 따라 실행이 차단될 수 있습니다.
- 회사 환경에서는 관리자 정책 확인이 필요합니다.

---

## 커스터마이징

- 백업 경로 변경:

Public Const BACKUP_BASE_PATH As String = "D:\backup\mail\"

- Inbox 루트 백업 비활성화
- EML 형식 저장
- 중복 백업 완전 차단(EntryID 기준)
- 월 단위(YYYY\MM) 분리

---

## 라이선스

MIT License  
자유롭게 사용 / 수정 가능하며, 사용에 따른 책임은 사용자에게 있습니다.

---

## 기여

Issue / Pull Request 환영합니다.
