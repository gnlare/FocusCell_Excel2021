# FocusCell2021 1.0

FocusCell2021 is an Excel-DNA XLL add-in for **Microsoft Excel 2021 on Windows**, designed to provide a Focus Cell-style row/column highlight.

Instead of changing workbook formatting, it draws a **click-through transparent overlay over Excel**.

## Features

- Highlights the active cell's row and/or column
- Keeps the **active cell interior fully transparent**
- Does not modify cell formatting or workbook content
- Focus Cell toggle and Settings button on the View tab
- Configurable highlight color and opacity
- Row + Column / Row only / Column only modes
- Configurable selected-cell border, thickness, and opacity
- Automatically hides while the Excel window is moving or resizing
- Responds to Excel zoom changes
- Handles merged cells
- Handles frozen/split panes
- Hides on Backstage/File screens where no worksheet cells are visible
- Detects the actual on-screen cell viewport using Excel hit-testing
  - excludes sheet tabs
  - excludes horizontal and vertical scroll bars
  - excludes row/column headers
  - excludes pane split areas
- Builds Korean and English editions from one source project

## Development requirements

- Windows
- Microsoft Excel 2021
- .NET Framework 4.8
- Excel-DNA 1.9.0
- Visual Studio 2022 recommended
  - .NET desktop development workload
  - .NET Framework 4.8 Developer Pack

## Build

Close Excel and run from the repository root:

```bat
build_release.cmd
```

`build_release.ps1` performs the actual build.

Successful builds are copied to:

```text
INSTALL\
├─ KR\
└─ EN\
```

Use an XLL containing `64` for 64-bit Excel, and an XLL without `64` for 32-bit Excel.

## Install in Excel

1. Start Excel.
2. Open **File > Options > Add-ins**.
3. At the bottom, select **Excel Add-ins** and click **Go**.
4. Click **Browse**.
5. Select an XLL from `INSTALL\KR` or `INSTALL\EN`.
6. Open the **View** tab and verify the Focus Cell group.

During development, you can also open the XLL directly from Excel for testing.

## Settings locations

Korean edition:

```text
%AppData%\FocusCell2021_KR\settings.ini
```

English edition:

```text
%AppData%\FocusCell2021_EN\settings.ini
```

The two editions keep separate settings.

## Viewport calculation

FocusCell2021 does not subtract fixed scrollbar or sheet-tab sizes from an Excel window rectangle.

It uses Excel screen hit-testing to find the pixels that actually contain worksheet cells and clips the overlay to that region. This avoids fixed-size assumptions across zoom levels, window sizes, scrollbars, sheet tabs, and pane layouts.

## GitHub Actions

`.github/workflows/build.yml` builds both KR and EN editions on Windows and uploads the generated `INSTALL` directory as a workflow artifact.

## Version

**1.0**

## License

No license file is included. Add the license you want before public redistribution if you intend to grant reuse or modification rights.
