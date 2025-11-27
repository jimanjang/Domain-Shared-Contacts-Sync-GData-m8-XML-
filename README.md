# Domain Shared Contacts Sync (GData m8 XML) README

이 스크립트는 Google Apps Script로 **도메인 공유 연락처(Directory / Shared Contacts)** 를
레거시 Contacts API(GData m8 XML) 방식으로 가져와 스프레드시트에 저장하고,
시트에서 수정/삭제한 내용을 다시 API로 반영하기 위한 예제입니다.

> **주의:** 이 코드는 Google Contacts의 레거시 GData API를 사용합니다.
> 조직 정책/Google API 변경에 따라 동작이 제한될 수 있습니다.

---

## 주요 기능

### 1) 공유 연락처 전체 조회 → 시트 저장

* 함수: `getAllSharedContacts()`
* 도메인(`@domain`) 기준 공유 연락처 피드 조회
* 페이지네이션(`rel="next"`) 처리로 전체 연락처 수집
* 각 연락처의 상세 필드(이름/이메일/전화/조직/주소/생일/웹사이트/노트 등) 파싱
* 결과를 시트에 헤더 포함 저장

### 2) 시트 기반 삭제 반영

* 함수: `deleteSharedContactsFromSheet()`
* 시트의 `삭제`(또는 `delete`) 컬럼에 `y / yes` 표기된 행을 찾음
* `EditLink`(edit URL)를 사용해 해당 연락처 DELETE 요청
* 삭제 과정 Response Code / Body 로깅 강화

### 3) 시트 기반 업데이트 반영(현재 스켈레톤)

* 함수: `updateSharedContactsFromSheet()`
* 현재는 **로그만 있는 상태**이며 실제 업데이트 로직은 생략됨
* 필요 시 기존 로직을 추가/유지해서 사용

---

## 사전 준비

### 1. Google Cloud Console에서 Contacts API 활성화

* 프로젝트에서 **Contacts API(contacts.googleapis.com)** 를 활성화해야 합니다.

### 2. `appsscript.json`에 OAuth Scope 추가

프로젝트 설정 파일(`appsscript.json`)에 아래 스코프가 포함되어야 합니다.

```json
"oauthScopes": [
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.google.com/m8/feeds"
]
```

### 3. 권한 재승인

스코프 수정 후에는 스크립트 권한을 다시 승인해야 합니다.

---

## 설정 값

코드 상단에서 스프레드시트와 시트를 지정합니다.

```js
var SHEET_ID   = 'YOUR_SHEET_ID';
var SHEET_NAME = 'list';
```

* `SHEET_ID` : 결과를 기록할 스프레드시트 ID
* `SHEET_NAME` : 기록할 시트명 (없으면 자동 생성)

---

## 출력 시트 구조

`getAllSharedContacts()` 실행 시 시트는 아래 헤더 구조로 초기화됩니다.

| 컬럼명           | 설명                       |
| ------------- | ------------------------ |
| ID            | 연락처 feed id              |
| EditLink      | 수정/삭제 API endpoint       |
| Title         | 연락처 title                |
| Full Name     | 전체 이름                    |
| Given Name    | 이름                       |
| Family Name   | 성                        |
| Emails        | 이메일 목록 (`;`로 join)       |
| Phones        | 전화번호 목록                  |
| Organizations | 조직명/직책                   |
| Addresses     | 주소 (formattedAddress 우선) |
| Birthday      | 생일 (yyyy-mm-dd 형태)       |
| Websites      | 웹사이트 목록                  |
| Note          | 노트/메모                    |
| 삭제            | 삭제 마킹용 컬럼                |

---

## 사용 방법

### Step 1) 공유 연락처를 시트로 가져오기

1. Apps Script 편집기에서 `getAllSharedContacts()` 실행
2. 실행 로그에서 fetch / pagination 진행 확인
3. 지정한 시트에 연락처 전체가 기록됨

---

### Step 2) 시트에서 삭제할 연락처 마킹

1. 삭제할 행의 `삭제` 컬럼에 `y` 또는 `yes` 입력
   (대소문자 무관)

예시:

| Full Name | Emails                                      | 삭제 |
| --------- | ------------------------------------------- | -- |
| Alice Kim | [alice@domain.com](mailto:alice@domain.com) |    |
| Bob Lee   | [bob@domain.com](mailto:bob@domain.com)     | y  |

---

### Step 3) 삭제 반영 실행

1. Apps Script에서 `deleteSharedContactsFromSheet()` 실행
2. `y/yes` 표시된 행의 `EditLink`로 DELETE 요청 수행
3. Logger에서 응답 코드 확인

---

## 로깅(Logging)

모든 주요 단계에서 상세 로그를 남기도록 되어 있습니다.

* 요청 URL
* HTTP Response Code
* 일부 Response Body(최대 500자)
* 삭제 대상 행/마킹값
* DELETE 응답 코드/본문

문제 발생 시 **실행 로그를 그대로 공유하면 디버깅이 쉬워집니다.**

---

## 주의사항 / 제한

1. **max-results=1000**

   * 한 페이지에 최대 1000개를 가져오며, `next` 링크를 따라 전체 수집합니다.

2. **EditLink 의존**

   * 삭제/수정은 시트에 저장된 `EditLink`가 정확해야만 동작합니다.
   * 시트를 임의로 편집하거나 값이 깨지면 삭제가 실패할 수 있습니다.

3. **레거시 API**

   * `https://www.google.com/m8/feeds` GData 방식은 구형이며
   * Google 정책 변경에 따라 차단될 수 있습니다.

4. **업데이트 로직 미포함**

   * `updateSharedContactsFromSheet()`는 스켈레톤입니다.
   * 실제 수정/추가 반영 로직이 필요하면 구현해야 합니다.

---

## 확장/개선 아이디어

* 시트 변경 감지 후 **diff 기반 업데이트**
* 삭제 성공 시 시트 행 자동 표시/백업
* 특정 OU/라벨/조건 기반 필터링
* 신규 연락처 추가(create) 지원
* People API 기반으로 신규 마이그레이션 검토

---

원하시면 `updateSharedContactsFromSheet()` 실제 업데이트 로직(이름/이메일/전화/조직/주소 등)을
**현재 시트 구조에 맞춰 완성**해드릴게요.
