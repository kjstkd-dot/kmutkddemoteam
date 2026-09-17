# K-STAR 현황판

계명대학교 태권도 시범단 역량관리(COMpass K / K-STAR) 현황판입니다.

- 단원은 학번과 비밀번호로 들어와 자기 역량 현황과 다음 단계를 확인합니다.
- 관리자는 전체 단원을 S·T·A·R 단계별로 보고, 점수를 고치거나 비밀번호를 초기화합니다.
- 휴대폰에서 홈 화면에 추가하면 앱처럼 쓸 수 있습니다.

## 파일

| 파일 | 설명 |
|---|---|
| `index.html` | 앱 본체 |
| `manifest.json` | 앱 설치 정보 |
| `sw.js` | 오프라인 동작 |
| `icon-180.png` `icon-192.png` `icon-512.png` | 홈 화면 아이콘 |
| `og-image.png` | 카카오톡 미리보기 이미지 |
| `apps-script.gs` | 자료 저장 서버 (구글 Apps Script 에 붙여넣는 코드) |

## 자료 저장 연결

`index.html` 위쪽의 아래 줄에 구글 Apps Script 웹 앱 주소를 넣으면 서버 저장으로 바뀝니다.

```js
var API_URL = '';
```

비워두면 그 기기 안에만 저장됩니다.

## 다른 앱

`convergence/` 폴더에는 같은 구조로 만든 "융합전공 이수 현황판"이 들어있습니다. 자세한 내용은 `convergence/README.md` 를 보세요.

## 카카오톡 미리보기

배포 주소가 정해지면 `index.html` 의 `og:image` 두 줄을 전체 주소로 바꿔주세요.

```html
<meta property="og:image" content="https://주소/og-image.png">
<meta property="og:image:secure_url" content="https://주소/og-image.png">
```
