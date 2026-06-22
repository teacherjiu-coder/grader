# 컴활 실기 자동채점기

컴퓨터활용능력 실기 기출 파일(.xlsm)을 업로드하면 자동으로 채점하는 정적 웹앱입니다.

## 로컬 테스트

프로젝트 루트(`자동채점앱/`)에서 실행:

```bash
python3 -m http.server 8765
```

브라우저에서 `http://localhost:8765` 접속 (`file://`로는 정답키 fetch가 동작하지 않습니다).

## GitHub Pages 배포

저장소 Settings → Pages → Source를 **main 브랜치 / root** 로 설정하면 `https://<username>.github.io/<repo>/` 에 배포됩니다.

## 정답키 추가

`answers/{급수}/{회차}.json` 형식으로 JSON 파일을 추가하면 해당 회차 버튼이 자동으로 활성화됩니다.
