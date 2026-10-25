# 한글 파일 플러그인 (hangul-file)

Claude가 **한컴오피스 한글 문서(.hwp, .hwpx)** 를 한글 프로그램 없이 다루게 해 주는 무료 플러그인입니다.

Claude는 워드 파일은 잘 다루지만 한글 파일은 잘 못 읽습니다. 이 플러그인을 넣으면 학교·관공서·회사 서류를 Claude에게 그대로 올려서 일을 시킬 수 있습니다.

## 할 수 있는 일

| 이렇게 말하면 | 이렇게 해 줍니다 |
|---|---|
| "이 한글 파일 읽고 요약해 줘" | 글, 표, 각주까지 읽고 정리 |
| "홍길동을 김철수로 전부 바꿔 줘" | 글자 모양은 그대로 두고 바꿔서 같은 형식(.hwp/.hwpx)으로 돌려줌 |
| "이 신청서 채워 줘" | `{{이름}}` 빈칸, 표의 "성명" 옆 칸, 누름틀을 채움 |
| "이 명단대로 한 장씩 만들어 줘" | 엑셀 명단(.xlsx 그대로, CSV도 됨)으로 사람마다 한 파일씩 |
| "이 내용으로 한글 문서 만들어 줘" | 제목·표·글머리표·그림이 있는 새 .hwpx 문서 |
| "hwp로 바꿔 줘" | hwp ↔ hwpx 바꾸기 |

깨진 파일을 주지 않도록, 한글이 못 열 것 같은 결과는 **저장하지 않고 이유를 알려 줍니다.**

## 설치

Claude Code에서:

```
/plugin marketplace add rlawhddja1231-hash/hangul-file
/plugin install hangul-file@hangul-file
```

Claude 앱(claude.ai)에서는 `skills/hangul-file` 폴더를 압축해 **설정 → 기능 → 스킬**에 올리면 됩니다.

처음 쓸 때 Claude가 필요한 파이썬 패키지를 알아서 설치합니다:

```
pip install python-hwpx olefile openpyxl pillow
```

## 시험 결과 (2026-10-24 ~ 25)

- **공개 한글 파일 2,798개 대량 시험**: 파일마다 읽기 → 고치기 → 저장 → 다시 읽기. 읽은 2,703개 모두 고친 곳 말고는 글자가 하나도 안 바뀜. 못 연 것은 한글 3.0(1997년 이전)·암호·배포용·깨진 파일이고, 이때는 쉬운 말로 이유와 해결 방법을 알려 줌
- **실제 관공서 서식 149개**(국가법령정보센터 신청서·신고서·계약서): 표 칸 채우기 105개 중 101개 성공
- **한컴오피스 2024 한글 뷰어**(진짜 한컴 엔진)로 새로 만든 문서(.hwpx/.hwp)·고친 공문·채운 양식·명단 문서·그림 문서를 열어 확인
- 7MB 실제 공문(표 109개, 그림 75개)에서 154곳 바꾸기 → 본문 말고는 파일이 한 바이트도 안 바뀜

자세한 기록은 [`skills/hangul-file/references/log.md`](skills/hangul-file/references/log.md)에 있습니다.

## 못 하는 것

- 암호가 걸린 파일, 배포용(읽기 전용) 파일
- 한글 프로그램 화면과 똑같은 미리 보기 (비슷한 모양만)
- 일부 특수 필드가 든 옛 .hwp 저장 → 한글에서 HWPX로 저장한 파일을 주면 됩니다

## 직접 써 보기 (명령어)

```
python3 skills/hangul-file/scripts/hangul.py info    문서.hwp
python3 skills/hangul-file/scripts/hangul.py read    문서.hwp --out 문서.md
python3 skills/hangul-file/scripts/hangul.py replace 문서.hwpx 홍길동 김철수
python3 skills/hangul-file/scripts/hangul.py fill    양식.hwpx 값.json
python3 skills/hangul-file/scripts/hangul.py merge   양식.hwpx 명단.xlsx --name '{index:02d}_{이름}.hwpx'
python3 skills/hangul-file/scripts/hangul.py new     원고.md --out 문서.hwpx
python3 skills/hangul-file/scripts/hangul.py convert 문서.hwp 문서.hwpx
python3 skills/hangul-file/scripts/hangul.py check   문서.hwpx
python3 skills/hangul-file/scripts/hangul.py preview 문서.hwpx
```

`examples/` 폴더에 원고와 양식 예시가 있습니다.

## 고마운 곳

문서를 읽고 쓰는 핵심은 [python-hwpx](https://github.com/airmang/python-hwpx) (Apache-2.0)가 합니다. 이 플러그인은 그 위에 한국어 사용법, 양식 채우기, 한글 이름 빈칸 여러 장 만들기, 마크다운 문서 만들기, 큰 .hwp 예비 읽기, 안전 저장을 더했습니다.

"한글", "한컴오피스"는 (주)한글과컴퓨터의 상표입니다. 이 플러그인은 한글과컴퓨터와 관계없는 개인 프로젝트입니다.

## 라이선스

MIT
