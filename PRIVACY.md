# 개인정보 처리 안내 (Privacy)

**한국어**

- 이 플러그인은 사용자가 고른 한글 파일(.hwp/.hwpx)과 명단(CSV/엑셀/JSON)을 **사용자의 컴퓨터(또는 Claude 작업 공간) 안에서만** 읽고 씁니다.
- 문서 내용, 이름·주소·전화번호 같은 개인정보를 **어떤 서버로도 보내지 않고, 따로 저장하지도 않습니다.** 결과 파일은 사용자가 정한 곳에만 만들어집니다.
- 미리 보기·PDF 기능을 처음 쓸 때 공개 npm 패키지 `@rhwp/core`(고정 판 0.8.7)를 한 번 내려받습니다. 이때 문서나 개인정보는 보내지 않습니다.
- 비밀번호, API 키, 토큰을 읽거나 쓰지 않습니다. 화면 그리기 도구(node)에는 PATH·HOME 같은 실행에 꼭 필요한 환경 값만 넘깁니다.
- "개인정보 가리기"(redact) 기능은 주민등록번호·전화번호·이메일 등을 찾아 문서 안에서 가립니다. 찾은 값은 화면에만 보여 주고 어디에도 기록하지 않습니다.

**English**

- The plugin reads and writes only the Hangul documents and rosters the user selects, locally (on the user's machine or in Claude's workspace).
- It sends no document content or personal data to any server and keeps no copy. Output files are written only where the user asks.
- On first use of preview/PDF it downloads the public npm package `@rhwp/core` (pinned 0.8.7) once; no document data is sent.
- It does not read or use passwords, API keys or tokens. Only the environment values needed to run (PATH, HOME, temp folders, etc.) are passed to the node renderer.
- The redact feature masks resident numbers, phone numbers and emails inside the document; matches are only shown to the user, never logged elsewhere.

Contact: GitHub issues at https://github.com/rlawhddja1231-hash/hangul-file/issues
