using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;
using FocusCell2021.Interop;
using FocusCell2021.Settings;

namespace FocusCell2021.Overlay
{
    internal sealed class FocusOverlayWindow : Window
    {
        private readonly Canvas _canvas;
        private readonly Rectangle _rowLeftHighlight;
        private readonly Rectangle _rowRightHighlight;
        private readonly Rectangle _columnTopHighlight;
        private readonly Rectangle _columnBottomHighlight;
        private readonly Rectangle _cellBorder;
        private readonly IntPtr _ownerHwnd;
        private IntPtr _hwnd;

        public IntPtr Handle => _hwnd;

        public FocusOverlayWindow(IntPtr ownerHwnd)
        {
            _ownerHwnd = ownerHwnd;

            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            ShowInTaskbar = false;
            ShowActivated = false;
            Topmost = false;
            Focusable = false;
            IsHitTestVisible = false;
            SnapsToDevicePixels = true;
            UseLayoutRounding = true;

            _canvas = new Canvas
            {
                Background = Brushes.Transparent,
                IsHitTestVisible = false,
                ClipToBounds = true,
                UseLayoutRounding = true,
                SnapsToDevicePixels = true
            };

            // Split the row/column overlays around the active cell instead of drawing
            // through it. This leaves the current cell completely transparent while
            // still highlighting the rest of its row and column.
            _rowLeftHighlight = MakeRectangle();
            _rowRightHighlight = MakeRectangle();
            _columnTopHighlight = MakeRectangle();
            _columnBottomHighlight = MakeRectangle();
            _cellBorder = MakeRectangle();
            _cellBorder.Fill = Brushes.Transparent;

            _canvas.Children.Add(_rowLeftHighlight);
            _canvas.Children.Add(_rowRightHighlight);
            _canvas.Children.Add(_columnTopHighlight);
            _canvas.Children.Add(_columnBottomHighlight);
            _canvas.Children.Add(_cellBorder);
            Content = _canvas;

            SourceInitialized += OnSourceInitialized;
        }

        private static Rectangle MakeRectangle()
        {
            return new Rectangle
            {
                IsHitTestVisible = false,
                SnapsToDevicePixels = true,
                UseLayoutRounding = true
            };
        }

        private void OnSourceInitialized(object sender, EventArgs e)
        {
            var helper = new WindowInteropHelper(this);
            if (_ownerHwnd != IntPtr.Zero)
                helper.Owner = _ownerHwnd;

            _hwnd = helper.Handle;
            var exStyle = NativeMethods.GetWindowLong(_hwnd, NativeMethods.GWL_EXSTYLE);
            exStyle |= NativeMethods.WS_EX_TRANSPARENT |
                       NativeMethods.WS_EX_NOACTIVATE |
                       NativeMethods.WS_EX_TOOLWINDOW;
            NativeMethods.SetWindowLong(_hwnd, NativeMethods.GWL_EXSTYLE, exStyle);
        }

        public void HideImmediately()
        {
            try
            {
                if (_hwnd != IntPtr.Zero)
                    NativeMethods.ShowWindow(_hwnd, NativeMethods.SW_HIDE);
                else if (IsVisible)
                    Hide();
            }
            catch { }
        }

        private void ShowWithoutActivation()
        {
            try
            {
                if (!IsVisible)
                {
                    Show();
                    return;
                }

                if (_hwnd != IntPtr.Zero)
                    NativeMethods.ShowWindow(_hwnd, NativeMethods.SW_SHOWNOACTIVATE);
            }
            catch { }
        }

        public void Apply(OverlayGeometry g, FocusSettings settings)
        {
            var scale = g.DpiScale <= 0 ? 1.0 : g.DpiScale;

            Left = g.PaneLeft / scale;
            Top = g.PaneTop / scale;
            Width = Math.Max(1, g.PaneWidth / scale);
            Height = Math.Max(1, g.PaneHeight / scale);

            var relativeCellLeft = (g.CellLeft - g.PaneLeft) / scale;
            var relativeCellTop = (g.CellTop - g.PaneTop) / scale;
            var cellWidth = Math.Max(1, g.CellWidth / scale);
            var cellHeight = Math.Max(1, g.CellHeight / scale);
            var relativeCellRight = relativeCellLeft + cellWidth;
            var relativeCellBottom = relativeCellTop + cellHeight;

            var baseColor = (Color)ColorConverter.ConvertFromString(settings.HighlightColor);
            var fillColor = Color.FromArgb((byte)Math.Round(255 * settings.Opacity), baseColor.R, baseColor.G, baseColor.B);
            var borderColor = Color.FromArgb((byte)Math.Round(255 * settings.BorderOpacity), baseColor.R, baseColor.G, baseColor.B);

            var fillBrush = new SolidColorBrush(fillColor);
            fillBrush.Freeze();
            var borderBrush = new SolidColorBrush(borderColor);
            borderBrush.Freeze();

            _rowLeftHighlight.Fill = fillBrush;
            _rowRightHighlight.Fill = fillBrush;
            _columnTopHighlight.Fill = fillBrush;
            _columnBottomHighlight.Fill = fillBrush;
            _cellBorder.Stroke = borderBrush;
            _cellBorder.StrokeThickness = settings.BorderThickness;
            _cellBorder.Visibility = settings.ShowCellBorder ? Visibility.Visible : Visibility.Collapsed;

            var rowEnabled = settings.Mode != FocusHighlightMode.ColumnOnly;
            var columnEnabled = settings.Mode != FocusHighlightMode.RowOnly;

            // Clamp the cut-out edges to the overlay window. The active cell can be
            // partially clipped when it is at the edge of a pane or frozen region.
            var holeLeft = Clamp(relativeCellLeft, 0, Width);
            var holeRight = Clamp(relativeCellRight, 0, Width);
            var holeTop = Clamp(relativeCellTop, 0, Height);
            var holeBottom = Clamp(relativeCellBottom, 0, Height);

            // Row highlight, left and right of the active cell.
            PlaceRectangle(
                _rowLeftHighlight,
                0,
                relativeCellTop,
                holeLeft,
                cellHeight,
                rowEnabled);

            PlaceRectangle(
                _rowRightHighlight,
                holeRight,
                relativeCellTop,
                Math.Max(0, Width - holeRight),
                cellHeight,
                rowEnabled);

            // Column highlight, above and below the active cell.
            PlaceRectangle(
                _columnTopHighlight,
                relativeCellLeft,
                0,
                cellWidth,
                holeTop,
                columnEnabled);

            PlaceRectangle(
                _columnBottomHighlight,
                relativeCellLeft,
                holeBottom,
                cellWidth,
                Math.Max(0, Height - holeBottom),
                columnEnabled);

            Canvas.SetLeft(_cellBorder, relativeCellLeft);
            Canvas.SetTop(_cellBorder, relativeCellTop);
            _cellBorder.Width = cellWidth;
            _cellBorder.Height = cellHeight;

            ShowWithoutActivation();
        }

        private static void PlaceRectangle(
            Rectangle rectangle,
            double left,
            double top,
            double width,
            double height,
            bool enabled)
        {
            if (!enabled || width <= 0.01 || height <= 0.01)
            {
                rectangle.Visibility = Visibility.Collapsed;
                return;
            }

            rectangle.Visibility = Visibility.Visible;
            Canvas.SetLeft(rectangle, left);
            Canvas.SetTop(rectangle, top);
            rectangle.Width = width;
            rectangle.Height = height;
        }

        private static double Clamp(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }
    }
}
