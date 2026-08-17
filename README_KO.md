# FocusCell2021 1.0

Microsoft Excel 2021(Windows)에서 Microsoft 365의 Focus Cell과 비슷한 행/열 강조 기능을 제공하는 Excel-DNA 기반 XLL 애드인입니다.

워크북에 셀 색상, Shape, 조건부서식 등을 기록하지 않고 **Excel 화면 위에 클릭이 통과되는 투명 오버레이**를 표시합니다.

## 주요 기능

- 선택 셀의 행/열 강조
- 선택된 **현재 셀 내부는 완전 투명**하게 유지
- 기존 셀 서식 및 워크북 내용 변경 없음
- `보기` 탭의 Focus Cell ON/OFF 및 설정 버튼
- 강조 색상과 투명도 설정
- 행+열 / 행만 / 열만 선택
- 선택 셀 테두리 표시, 굵기, 투명도 설정
- Excel 창 이동/크기 변경 중 오버레이 자동 숨김
- Zoom 변경 대응
- 병합 셀 대응
- 틀 고정 및 분할 Pane 대응
- 파일 메뉴(Backstage) 등 셀 영역이 없는 화면에서 오버레이 숨김
- 실제 셀 표시 영역만 화면 hit-test로 계산
  - 시트 탭 영역 제외
  - 가로/세로 스크롤바 제외
  - 행/열 머리글 제외
  - Pane splitter 제외
- 하나의 소스에서 한글판(KR)과 영문판(EN) 동시 생성

## 개발 환경

- Windows
- Microsoft Excel 2021
- .NET Framework 4.8
- Excel-DNA 1.9.0
- Visual Studio 2022 권장
  - `.NET 데스크톱 개발` 워크로드
  - .NET Framework 4.8 Developer Pack

## 빌드

Excel을 완전히 종료한 후 프로젝트 루트의 아래 파일을 실행합니다.

```text
build_release.cmd
```

실제 빌드는 `build_release.ps1`이 처리합니다.

성공하면 프로젝트 최상위에 다음 폴더가 생성됩니다.

```text
INSTALL\
├─ KR\    한글판
└─ EN\    영문판
```

64비트 Excel은 파일명에 `64`가 포함된 XLL을 사용하고, 32비트 Excel은 `64`가 없는 XLL을 사용합니다.

## Excel에 설치

1. Excel 실행
2. `파일 > 옵션 > 추가 기능`
3. 하단 `관리: Excel 추가 기능` 선택 후 `이동`
4. `찾아보기` 선택
5. `INSTALL\KR` 또는 `INSTALL\EN`의 XLL 선택
6. `보기` 탭에서 Focus Cell 확인

개발 중에는 XLL 파일을 Excel에서 직접 열어 테스트할 수도 있습니다.

## 설정 저장 위치

한글판:

```text
%AppData%\FocusCell2021_KR\settings.ini
```

영문판:

```text
%AppData%\FocusCell2021_EN\settings.ini
```

두 언어판의 설정은 서로 분리됩니다.

## 화면 영역 계산 방식

FocusCell2021은 Excel 창 크기에서 스크롤바 높이나 시트 탭 높이를 고정값으로 빼지 않습니다.

Excel의 화면 hit-test를 이용해 **실제로 셀이 존재하는 화면 픽셀 영역**을 찾고, 그 범위 안에서만 오버레이를 표시합니다. 이 방식으로 Excel 확대/축소, 창 크기, 스크롤바, 시트 탭, Pane 경계에 따른 오차를 줄입니다.

## GitHub Actions

`.github/workflows/build.yml`이 포함되어 있습니다. GitHub에서 Actions를 실행하면 Windows 환경에서 KR/EN을 빌드하고 `INSTALL` 폴더를 Artifact로 업로드합니다.

## 버전

**1.0**

## 라이선스

현재 저장소에는 라이선스 파일을 포함하지 않았습니다. 공개 배포 시 다른 사용자의 재사용/수정 권한을 허용하려면 원하는 라이선스를 추가하세요.
