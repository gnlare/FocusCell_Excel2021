namespace FocusCell2021.Localization
{
    /// <summary>
    /// Compile-time localization shared by the KR and EN builds.
    /// The source tree is identical for both editions; build_release.cmd sets FocusLanguage.
    /// </summary>
    internal static class AppText
    {
#if FOCUS_LANG_EN
        public const string BuildCode = "EN";
        public const string SettingsFolderName = "FocusCell2021_EN";
        public const string SettingsFileHeader = "# FocusCell2021 EN user settings";

        public const string RibbonGroupId = "FocusCell2021ENGroup";
        public const string RibbonToggleId = "FocusCellENToggle";
        public const string RibbonSettingsId = "FocusCellENSettings";
        public const string FocusScreenTip = "Turn Focus Cell on or off";
        public const string FocusSuperTip = "Highlights the selected cell row and column with a transparent overlay. The workbook formatting is not modified.";
        public const string SettingsLabel = "Settings";
        public const string SettingsScreenTip = "Focus Cell settings";
        public const string SettingsSuperTip = "Configure color, opacity, border, highlight range, and behavior options.";

        public const string WindowTitle = "Focus Cell Settings";
        public const string Defaults = "Defaults";
        public const string Apply = "Apply";
        public const string Cancel = "Cancel";
        public const string Ok = "OK";
        public const string HighlightHeading = "Highlight";
        public const string HighlightColor = "Highlight color";
        public const string HighlightOpacity = "Highlight opacity";
        public const string HighlightRange = "Highlight range";
        public const string RowAndColumn = "Row + Column";
        public const string RowOnly = "Row only";
        public const string ColumnOnly = "Column only";
        public const string BorderHeading = "Selected cell border";
        public const string ShowBorder = "Show selected cell border";
        public const string BorderThickness = "Border thickness";
        public const string BorderOpacity = "Border opacity";
        public const string BehaviorHeading = "Behavior";
        public const string HideWhileMoving = "Hide overlay while moving/resizing Excel";
        public const string RefreshInterval = "Refresh interval";
        public const string RefreshFast = "Fast (35 ms)";
        public const string RefreshDefault = "Default (55 ms)";
        public const string RefreshPowerSaving = "Power saving (100 ms)";
        public const string RefreshSlow = "Slow (150 ms)";
        public static string SettingsNote =>
            "Settings are saved to %AppData%\\" + SettingsFolderName + "\\settings.ini.\n" +
            "The overlay does not modify cell formatting or workbook content.";
#else
        public const string BuildCode = "KR";
        public const string SettingsFolderName = "FocusCell2021_KR";
        public const string SettingsFileHeader = "# FocusCell2021 KR user settings";

        public const string RibbonGroupId = "FocusCell2021KRGroup";
        public const string RibbonToggleId = "FocusCellKRToggle";
        public const string RibbonSettingsId = "FocusCellKRSettings";
        public const string FocusScreenTip = "Focus Cell 켜기/끄기";
        public const string FocusSuperTip = "선택한 셀의 행과 열을 Excel 위의 투명 오버레이로 강조합니다. 워크북 서식은 변경하지 않습니다.";
        public const string SettingsLabel = "설정";
        public const string SettingsScreenTip = "Focus Cell 설정";
        public const string SettingsSuperTip = "색상, 투명도, 테두리, 강조 범위와 동작 옵션을 설정합니다.";

        public const string WindowTitle = "Focus Cell 설정";
        public const string Defaults = "기본값";
        public const string Apply = "적용";
        public const string Cancel = "취소";
        public const string Ok = "확인";
        public const string HighlightHeading = "강조 표시";
        public const string HighlightColor = "강조 색상";
        public const string HighlightOpacity = "강조 투명도";
        public const string HighlightRange = "강조 범위";
        public const string RowAndColumn = "행 + 열";
        public const string RowOnly = "행만";
        public const string ColumnOnly = "열만";
        public const string BorderHeading = "선택 셀 테두리";
        public const string ShowBorder = "선택한 셀 테두리 표시";
        public const string BorderThickness = "테두리 굵기";
        public const string BorderOpacity = "테두리 투명도";
        public const string BehaviorHeading = "동작";
        public const string HideWhileMoving = "Excel 창 이동/크기 조절 중 오버레이 숨기기";
        public const string RefreshInterval = "화면 갱신 주기";
        public const string RefreshFast = "빠름 (35 ms)";
        public const string RefreshDefault = "기본 (55 ms)";
        public const string RefreshPowerSaving = "절전 (100 ms)";
        public const string RefreshSlow = "느림 (150 ms)";
        public static string SettingsNote =>
            "※ 설정은 %AppData%\\" + SettingsFolderName + "\\settings.ini 에 저장됩니다.\n" +
            "※ 오버레이 방식이므로 셀 서식과 워크북 내용은 변경하지 않습니다.";
#endif
        public const string FocusCellLabel = "Focus Cell";
    }
}
