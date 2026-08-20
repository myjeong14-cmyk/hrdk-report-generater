# hrdk-report-generater
한국산업인력공단 경북지사 : 외부시험장 출장 증빙 보고서 자동생성기

<접속 url>
https://hrdk-report-generater-eqxyse95p2vw7seoa5ngku.streamlit.app/



<img width="812" height="617" alt="image" src="https://github.com/user-attachments/assets/e383a46a-40be-4482-a652-ae77321fcb64" />

## 도착지 추가 기능 - GitHub 자동 커밋 설정

"도착지" 드롭다운에서 `➕ 도착지 추가하기`로 새 도착지를 등록하면, `custom_destinations.json`과 첨부한 지도 캡처 이미지(`map/` 폴더)를 이 저장소에 **자동으로 커밋**합니다. Streamlit Community Cloud는 파일시스템이 임시적이라(재시작 시 초기화), GitHub에 커밋해 두어야 도착지 정보가 영구적으로 유지됩니다.

이 기능을 쓰려면 GitHub Personal Access Token(PAT)이 필요합니다.

1. GitHub → Settings → Developer settings → Personal access tokens에서 **이 저장소(`hrdk-report-generater`)에 대한 쓰기 권한(Contents: Read and write)**을 가진 토큰을 발급합니다.
2. Streamlit Cloud 앱 관리 화면 → **Settings → Secrets**에 아래처럼 추가합니다.

   ```toml
   GITHUB_TOKEN = "여기에_발급받은_토큰"
   ```
3. 저장 후 앱을 재시작(reboot)하면 적용됩니다.

토큰이 설정되어 있지 않으면, 도착지는 현재 세션에서는 추가되어 사용할 수 있지만 GitHub에는 커밋되지 않으며, 앱이 재시작되면 사라진다는 경고가 표시됩니다.

