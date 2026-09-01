# Focus Cell 2021 1.0

Microsoft Excel 2021(Windows)에서 Microsoft 365의 Focus Cell과 비슷한 행/열 강조 기능을 제공하는 Excel-DNA 기반 XLL 애드인입니다.
워크북에 셀 색상, Shape, 조건부서식 등을 기록하지 않고 Excel 화면 위에 클릭이 통과되는 투명 오버레이를 표시합니다.

A Focus Cell-style add-in for **Microsoft Excel 2021 on Windows**.
It highlights the active cell's row and/or column using a click-through transparent overlay, without modifying workbook formatting or contents.


<img width="306" height="321" alt="스크린샷 2026-08-17 162636" src="https://github.com/user-attachments/assets/b9b2e402-d64c-443a-9495-4cd0c92e3868" />



**한국어:** [README_KO.md](README_KO.md)  
**English:** [README_EN.md](README_EN.md)

## Highlights

- 선택 셀의 행/열 강조
- 선택된 현재 셀 내부는 완전 투명하게 유지
- 기존 셀 서식 및 워크북 내용 변경 없음
- 보기 탭의 Focus Cell ON/OFF 및 설정 버튼
- 강조 색상과 투명도 설정
- 행+열 / 행만 / 열만 선택

- Transparent overlay: does not write shapes, fills, or conditional formatting into the workbook
- Active cell itself remains transparent while the row/column is highlighted
- View-tab ribbon with **Focus Cell** toggle and **Settings** button
- Configurable color, opacity, row/column mode, selected-cell border, and refresh interval

## Build

Requirements:

- Windows
- Microsoft Excel 2021
- .NET Framework 4.8 Developer Pack

## Install in Excel

1. Excel 실행
2. 파일 > 옵션 > 추가 기능
3. 하단 관리: Excel 추가 기능 선택 후 이동
4. 찾아보기 선택
5. INSTALL\KR 폴더의 XLL 파일 선택
6. 보기 탭에서 Focus Cell 확인
-------------
1. Open Excel.
2. Go to **File > Options > Add-ins**.
3. At the bottom, choose **Excel Add-ins** and click **Go**.
4. Click **Browse** and select the appropriate XLL from `INSTALL\EN`.
5. Open the **View** tab and use **Focus Cell**.

## Version

Current public version: **1.0**

