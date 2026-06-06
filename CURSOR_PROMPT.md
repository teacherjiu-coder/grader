# 지우쌤 컴활 실기 자동채점기 — 빌드 지시서

너는 이 프로젝트를 **정적 웹앱**으로 완성한다. 학생이 자기가 푼 엑셀(.xlsm)을 드래그앤드롭하면 영역별 점수와 합격/불합격, 문제별 정오를 보여준다. GitHub Pages에 그대로 올라가야 하므로 **빌드 없이 동작하는 순수 HTML/JS/CSS**로 만든다(번들러 X, 외부 npm install X, 라이브러리는 CDN).

## 절대 규칙
- **`grading.js`는 채점 엔진이다. 절대 다시 작성하거나 로직을 수정하지 마라.** 그대로 `<script>`로 불러서 `window.gradeWorkbook`을 호출만 한다.
- 정답키는 `answers/{급수}/{회차}.json`에 이미 들어있다. (예: `answers/2급/3.json`, `answers/1급/1.json`) **fetch로 불러온다. 절대 수정/재생성하지 마라.**
- 채점은 100% 브라우저에서 처리(서버 없음). 업로드 파일은 외부로 전송하지 않는다.

## 폴더 구조 (이미 있는 것 + 네가 만들 것)
```
index.html        ← 네가 만든다
style.css         ← 네가 만든다 (또는 index.html에 인라인)
app.js            ← 네가 만든다
grading.js        ← 이미 있음. 손대지 마라
answers/
  1급/1.json
  2급/1.json ~ 10.json
```

## 라이브러리 (CDN, index.html에서 로드)
```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="grading.js"></script>
```
DOMParser는 브라우저 내장이다. (별도 로드 불필요)

## 채점 엔진 호출법 (이게 전부다)
```js
async function gradeFile(file, 급수, 회차) {
  const buf = new Uint8Array(await file.arrayBuffer());
  const KEY = await (await fetch(`answers/${급수}/${회차}.json`)).json();
  const result = window.gradeWorkbook(buf, { XLSX, JSZip, DOMParser }, KEY);
  return result;
}
```

### gradeWorkbook이 돌려주는 값 (이 형태로 온다)
```js
{
  total: 88,            // 총점 (0~100)
  pass: true,           // total >= 70 이면 true (합격)
  autoMax: 81,          // 자동채점 가능 만점 (참고용)
  manualPts: 19,        // 수동(=만점처리) 점수 (참고용)
  results: [
    {
      id: "계산 1번 신청과목",   // 문제 이름
      area: "계산작업",          // 영역: 기본작업/계산작업/분석작업/기타작업
      points: 6,                 // 배점
      earned: 6,                 // 획득점수 (0 또는 points)
      ok: true,                  // 정답 여부
      manual: false,             // true면 자동확인 불가 → 이미 만점처리됨
      msg: "",                   // 틀렸을 때 사유 (ok=true면 보통 빈값)
      answer: "값"               // 정답/채점기준 힌트 (틀렸을 때 보여주면 좋음)
    },
    ...
  ]
}
```
- **`manual: true`인 항목은 이미 만점(earned=points) 처리되어 있다.** (VBA 폼, 차트 색/그림자 등 자동확인 불가 항목 → 학생에게 ✗ 대신 "👀 확인 필요(만점 인정)"로 부드럽게 표시)
- 영역별 점수는 `results`를 `area`로 묶어 `earned` / `points` 합산해서 만든다. (엔진이 영역합을 따로 주지 않음)

## 화면 흐름 (3단계)
1. **급수 선택**: `2급`, `1급` 두 버튼.
2. **회차 선택**: 선택한 급수의 가능한 회차만 활성화.
   - 2급 → 1~10회 모두 활성
   - 1급 → 1회만 활성, 2회~ 는 비활성(흐리게) + "준비 중" 뱃지
   - (회차 목록은 하드코딩하지 말고, fetch 실패하면 비활성 처리해도 됨. 단순하게 2급:[1..10], 1급:[1] 로 둬도 OK)
3. **파일 업로드**: 큰 드래그앤드롭 영역 + 파일선택 버튼. `.xlsm/.xlsx` 받음. 드롭 즉시 채점.
4. **결과 화면**:
   - 상단에 **큰 총점**(예: `88점`)과 **합격/불합격 배지**(70점 기준, 합격=초록, 불합격=빨강).
   - **영역별 점수 카드 4개**(기본작업/계산작업/분석작업/기타작업): 각 `획득/배점`과 진행바.
   - **문제별 리스트**: 영역별로 그룹핑. 각 줄에 아이콘(✓ 초록 / ✗ 빨강 / 👀 수동) + 문제이름 + `획득/배점`. 틀린 항목(✗)은 아래에 `answer`(채점기준/정답 힌트)와 `msg`(사유)를 작게 보여준다.
   - "다시 채점하기" 버튼 → 1단계로.

## 디자인 (지우쌤 브랜드)
- 톤: **초록 `#0d7a35`(메인) / 크림 `#f0f1eb`(배경)** / 텍스트 진회색.
- 폰트: **Pretendard** (CDN: `https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css`), 코드/점수 강조엔 JetBrains Mono.
- 헤더 타이틀: **"지우쌤 · 컴활 실기 자동채점기"** (부제: "내 답안 올리면 바로 채점").
  - ⚠️ "퇴근 30분 줄여주는" 슬로건은 **쓰지 마라**(그건 엑셀실무용이다).
- 합격 배지엔 가벼운 축하 느낌(이모지 🎉 정도 OK), 불합격도 기죽지 않게 응원 톤("조금만 더!").
- 결제/구독/광고 요소 **없음**. 깔끔하게.
- 모바일 반응형(폰에서도 드래그 대신 파일선택으로 동작).

## 자잘한 요구
- 채점 중 로딩 스피너.
- 파일이 해당 회차 양식이 아니어서 에러나면(시트 못 찾음 등) "이 파일은 [급수] [회차] 기출 양식이 아닌 것 같아요. 회차를 확인해주세요." 안내.
- 결과의 수동항목 안내문 한 줄: "👀 표시는 프로그램이 자동으로 확인하기 어려운 항목이라 만점으로 인정했어요. 실제 시험에선 채점 대상입니다."
- 콘솔에 raw result도 찍어둬(디버깅용).

먼저 `index.html` / `style.css` / `app.js` 전체를 만들고, 로컬에서 `python3 -m http.server`로 열어 2급 3회 / 1급 1회로 동작 확인까지 해줘.
