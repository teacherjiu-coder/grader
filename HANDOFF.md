# HANDOFF

## 1. 한 줄 요약
`자동채점앱`은 컴활 실기 `.xlsm/.xlsx` 답안 파일을 브라우저에서 바로 읽어 자동 채점하는 정적 웹앱이다.

## 2. 핵심 폴더 / 파일 구조
- `index.html`: UI 진입점. 급수 선택, 회차 선택, 업로드, 결과 화면.
- `app.js`: 화면 전환, 파일 업로드, 채점 결과 렌더링.
- `grading.js`: 실제 채점 엔진. 시트 값/수식/서식/XML을 읽어 정답 JSON과 비교.
- `style.css`: 전체 스타일.
- `answers/1급/1.json`: 1급 실기 정답키.
- `answers/2급/1.json` ~ `answers/2급/10.json`: 2급 실기 정답키.
- `README.md`: 현재 기본 실행/배포 설명.
- `CURSOR_PROMPT.md`: 채점 결과 UI/동작 의도 메모.

## 3. 로컬 실행 방법
정적 서버가 필요하다. `file://`로 열면 정답 JSON `fetch`가 막힌다.

```bash
cd "/Users/jiwoo/Desktop/업무/7.cursor/자동채점앱"
python3 -m http.server 8765
```

브라우저에서 `http://localhost:8765` 접속.

## 4. 빌드 & 배포
별도 빌드 단계는 없다. 정적 파일 그대로 배포한다.

### GitHub Pages 배포
```bash
cd "/Users/jiwoo/Desktop/업무/7.cursor/자동채점앱"
git add .
git commit -m "..."
git push origin main
```

GitHub Pages는 `main` 브랜치 `root` 기준으로 배포한다.

## 5. 데이터 / 정답키 관리
- 정답키 위치: `answers/{급수}/{회차}.json`
- 새 회차 추가 시 위 경로에 JSON 추가하면 UI에서 버튼이 자동 활성화된다.
- 현재 포함 회차:
  - `1급`: `1회`
  - `2급`: `1~10회`

## 6. 현재 알려진 버그 / 미완성
- 일부 문제는 자동 확인이 어려워 `manual` 항목으로 만점 처리한다.
  - 예: VBA 폼, 차트 그림자/윤곽선/둥근 모서리 같은 시각 요소
- 결과 화면의 `👀` 표시는 실제로는 수동 확인이 필요한 항목이다.
- 실행은 정적 서버 전제다. 로컬 파일 직접 열기는 불가.
- 현재 작업 트리에 `.DS_Store`와 `README.md` 수정 흔적이 있으니, 추가 커밋 전에 상태를 한 번 확인하는 편이 안전하다.

## 7. Git 리모트 / 배포 위치
- Git remote: `https://github.com/teacherjiu-coder/grader.git`
- 배포 URL: `https://teacherjiu-coder.github.io/grader/`
