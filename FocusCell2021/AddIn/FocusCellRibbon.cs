using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using FocusCell2021.Localization;

namespace FocusCell2021.AddIn
{
    public sealed class FocusCellRibbon : ExcelRibbon
    {
        private static IRibbonUI _ribbon;
        private static Bitmap _focusCellBitmap;
        private static object _focusCellPictureDisp;

        public override string GetCustomUI(string ribbonId)
        {
            return $@"<?xml version='1.0' encoding='UTF-8'?>
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab idMso='TabView'>
        <group id='{AppText.RibbonGroupId}' label='{AppText.FocusCellLabel}'>
          <toggleButton id='{AppText.RibbonToggleId}'
                        label='{AppText.FocusCellLabel}'
                        size='large'
                        getImage='GetFocusCellImage'
                        screentip='{AppText.FocusScreenTip}'
                        supertip='{AppText.FocusSuperTip}'
                        getPressed='GetFocusPressed'
                        onAction='OnToggleFocus'/>
          <button id='{AppText.RibbonSettingsId}'
                  label='{AppText.SettingsLabel}'
                  size='large'
                  imageMso='ApplicationOptionsDialog'
                  screentip='{AppText.SettingsScreenTip}'
                  supertip='{AppText.SettingsSuperTip}'
                  onAction='OnOpenSettings'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnLoad(IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        public object GetFocusCellImage(IRibbonControl control)
        {
            try
            {
                if (_focusCellPictureDisp != null)
                    return _focusCellPictureDisp;

                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("FocusCell2021.Assets.FocusCellBlue.png"))
                {
                    if (stream == null)
                        return null;

                    using (var source = new Bitmap(stream))
                    {
                        // Clone so the bitmap no longer depends on the resource stream lifetime.
                        _focusCellBitmap = new Bitmap(source);
                    }
                }

                _focusCellPictureDisp = PictureDispConverter.ToPictureDisp(_focusCellBitmap);
                return _focusCellPictureDisp;
            }
            catch
            {
                // A missing custom icon must never prevent the add-in itself from loading.
                return null;
            }
        }

        public bool GetFocusPressed(IRibbonControl control)
        {
            return FocusCellAddIn.Current?.Enabled ?? false;
        }

        public void OnToggleFocus(IRibbonControl control, bool pressed)
        {
            if (FocusCellAddIn.Current != null)
                FocusCellAddIn.Current.Enabled = pressed;
        }

        public void OnOpenSettings(IRibbonControl control)
        {
            FocusCellAddIn.Current?.ShowSettingsDialog();
        }

        internal static void Invalidate()
        {
            try { _ribbon?.Invalidate(); } catch { }
        }

        /// <summary>
        /// Office Ribbon image callbacks expect an OLE IPictureDisp. AxHost exposes
        /// the framework's built-in converter without requiring a direct stdole reference.
        /// </summary>
        private sealed class PictureDispConverter : AxHost
        {
            private PictureDispConverter() : base(string.Empty) { }

            internal static object ToPictureDisp(Image image)
            {
                return GetIPictureDispFromPicture(image);
            }
        }
    }
}
