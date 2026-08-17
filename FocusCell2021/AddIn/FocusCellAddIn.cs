using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using FocusCell2021.Overlay;
using FocusCell2021.Settings;

namespace FocusCell2021.AddIn
{
    public sealed class FocusCellAddIn : IExcelAddIn
    {
        private object _application;
        private FocusOverlayManager _overlayManager;

        public static FocusCellAddIn Current { get; private set; }
        public FocusSettings Settings { get; private set; }

        public bool Enabled
        {
            get => Settings?.Enabled ?? false;
            set
            {
                if (Settings == null) return;
                Settings.Enabled = value;
                Settings.Save();
                _overlayManager?.SetEnabled(value);
                FocusCellRibbon.Invalidate();
            }
        }

        public void AutoOpen()
        {
            Current = this;
            Settings = FocusSettings.Load();

            // ExcelDnaUtil.Application returns Excel's COM Application object.
            // Keep it late-bound so the add-in does not depend on Office PIA DLLs.
            _application = ExcelDnaUtil.Application;

            _overlayManager = new FocusOverlayManager(_application, Settings);
            _overlayManager.SetEnabled(Settings.Enabled);
        }

        public void AutoClose()
        {
            _overlayManager?.Dispose();
            _overlayManager = null;
            _application = null;
            Current = null;
        }

        public void ApplySettings()
        {
            Settings?.Save();
            _overlayManager?.ApplySettings();
            _overlayManager?.RefreshNow();
            FocusCellRibbon.Invalidate();
        }

        public void ShowSettingsDialog()
        {
            if (Settings == null) return;

            var ownerHwnd = GetActiveExcelWindowHwnd();
            var dialog = new FocusSettingsWindow(Settings, ownerHwnd, ApplySettings);
            dialog.ShowDialog();
        }

        private IntPtr GetActiveExcelWindowHwnd()
        {
            object window = null;
            try
            {
                dynamic app = _application;
                window = app.ActiveWindow;
                if (window == null) return IntPtr.Zero;
                dynamic dynWindow = window;
                return new IntPtr(Convert.ToInt64(dynWindow.Hwnd));
            }
            catch
            {
                return IntPtr.Zero;
            }
            finally
            {
                try
                {
                    if (window != null && Marshal.IsComObject(window))
                        Marshal.ReleaseComObject(window);
                }
                catch { }
            }
        }
    }
}
