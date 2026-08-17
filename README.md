# Focus Cell 2021 1.0

A Focus Cell-style add-in for **Microsoft Excel 2021 on Windows**.
It highlights the active cell's row and/or column using a click-through transparent overlay, without modifying workbook formatting or contents.


<img width="306" height="321" alt="스크린샷 2026-08-17 162636" src="https://github.com/user-attachments/assets/b9b2e402-d64c-443a-9495-4cd0c92e3868" />



**한국어:** [README_KO.md](README_KO.md)  
**English:** [README_EN.md](README_EN.md)

## Highlights

- Transparent overlay: does not write shapes, fills, or conditional formatting into the workbook
- Active cell itself remains transparent while the row/column is highlighted
- View-tab ribbon with **Focus Cell** toggle and **Settings** button
- Configurable color, opacity, row/column mode, selected-cell border, and refresh interval

## Build

Requirements:

- Windows
- Microsoft Excel 2021
- Visual Studio 2022 with **.NET desktop development**, or a .NET SDK capable of building .NET Framework 4.8 desktop projects
- .NET Framework 4.8 Developer Pack

Close Excel, then run:

```bat
build_release.cmd
```

Build output:

```text
INSTALL\
├─ KR\
└─ EN\
```

Use the XLL containing `64` for 64-bit Excel. Use the XLL without `64` for 32-bit Excel.

## Install in Excel

1. Open Excel.
2. Go to **File > Options > Add-ins**.
3. At the bottom, choose **Excel Add-ins** and click **Go**.
4. Click **Browse** and select the appropriate XLL from `INSTALL\KR` or `INSTALL\EN`.
5. Open the **View** tab and use **Focus Cell**.

## GitHub Actions

The included workflow builds both KR and EN editions on Windows and uploads the `INSTALL` folder as a workflow artifact.

## Version

Current public version: **1.0**

## License

No license file is included. Add a license before public redistribution if you want to grant reuse or modification rights.
