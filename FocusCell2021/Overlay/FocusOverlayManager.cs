using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using FocusCell2021.Interop;
using FocusCell2021.Settings;

namespace FocusCell2021.Overlay
{
    internal sealed class FocusOverlayManager : IDisposable
    {
        private readonly object _application;
        private readonly FocusSettings _settings;
        private readonly Timer _timer;
        private readonly NativeMethods.WinEventDelegate _winEventProc;
        private FocusOverlayWindow _overlay;
        private IntPtr _currentOwnerHwnd;
        private IntPtr _moveSizeHook;
        private bool _enabled;
        private bool _refreshing;
        private volatile bool _ownerMovingOrSizing;

        // RangeFromPoint hit-test cache. The expensive pixel-boundary search only runs when
        // the active cell, zoom, scroll position, pane/grid rectangle or row size changes.
        private string _verticalBoundsCacheKey;
        private double _verticalBoundsCacheTopPx;
        private double _verticalBoundsCacheBottomPx;

        public FocusOverlayManager(object application, FocusSettings settings)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _winEventProc = WinEventCallback;

            // Poll selection, scrolling, zoom, row/column resizing and pane changes.
            // Office PIA event interfaces are intentionally not required.
            _timer = new Timer { Interval = Math.Max(30, settings.RefreshIntervalMs) };
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        public void SetEnabled(bool enabled)
        {
            _enabled = enabled;
            if (!enabled)
            {
                Hide();
                return;
            }

            RefreshNow();
        }

        public void ApplySettings()
        {
            _timer.Interval = Math.Max(30, _settings.RefreshIntervalMs);
            if (!_settings.HideWhileMoving)
                _ownerMovingOrSizing = false;
        }

        public void RefreshNow()
        {
            if (!_enabled || _refreshing) return;

            if (_settings.HideWhileMoving && _ownerMovingOrSizing)
            {
                HideImmediate();
                return;
            }

            try
            {
                _refreshing = true;
                if (!TryGetGeometry(out var geometry, out var ownerHwnd))
                {
                    Hide();
                    return;
                }

                EnsureOverlay(ownerHwnd);
                EnsureMoveSizeHook(ownerHwnd);

                if (_settings.HideWhileMoving && _ownerMovingOrSizing)
                {
                    HideImmediate();
                    return;
                }

                _overlay.Apply(geometry, _settings);
            }
            catch
            {
                // Overlay failures must never interfere with normal Excel use.
                Hide();
            }
            finally
            {
                _refreshing = false;
            }
        }

        public void Hide()
        {
            try
            {
                if (_overlay != null && _overlay.IsVisible)
                    _overlay.Hide();
            }
            catch { }
        }

        private void HideImmediate()
        {
            try { _overlay?.HideImmediately(); } catch { }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            RefreshNow();
        }

        private void EnsureOverlay(IntPtr ownerHwnd)
        {
            if (_overlay != null && _currentOwnerHwnd == ownerHwnd)
                return;

            if (_overlay != null)
            {
                try { _overlay.Close(); } catch { }
                _overlay = null;
            }

            if (_currentOwnerHwnd != ownerHwnd)
            {
                ReleaseMoveSizeHook();
                _ownerMovingOrSizing = false;
            }

            _currentOwnerHwnd = ownerHwnd;
            _overlay = new FocusOverlayWindow(ownerHwnd);
        }

        private void EnsureMoveSizeHook(IntPtr ownerHwnd)
        {
            if (_moveSizeHook != IntPtr.Zero || ownerHwnd == IntPtr.Zero)
                return;

            try
            {
                NativeMethods.GetWindowThreadProcessId(ownerHwnd, out var processId);
                if (processId == 0) return;

                _moveSizeHook = NativeMethods.SetWinEventHook(
                    NativeMethods.EVENT_SYSTEM_MOVESIZESTART,
                    NativeMethods.EVENT_SYSTEM_MOVESIZEEND,
                    IntPtr.Zero,
                    _winEventProc,
                    processId,
                    0,
                    NativeMethods.WINEVENT_OUTOFCONTEXT);
            }
            catch
            {
                _moveSizeHook = IntPtr.Zero;
            }
        }

        private void ReleaseMoveSizeHook()
        {
            if (_moveSizeHook == IntPtr.Zero) return;
            try { NativeMethods.UnhookWinEvent(_moveSizeHook); } catch { }
            _moveSizeHook = IntPtr.Zero;
        }

        private void WinEventCallback(
            IntPtr hWinEventHook,
            uint eventType,
            IntPtr hwnd,
            int idObject,
            int idChild,
            uint dwEventThread,
            uint dwmsEventTime)
        {
            if (hwnd == IntPtr.Zero || hwnd != _currentOwnerHwnd)
                return;

            if (eventType == NativeMethods.EVENT_SYSTEM_MOVESIZESTART)
            {
                _ownerMovingOrSizing = true;
                if (_settings.HideWhileMoving)
                    HideImmediate();
            }
            else if (eventType == NativeMethods.EVENT_SYSTEM_MOVESIZEEND)
            {
                _ownerMovingOrSizing = false;
                // Excel COM must not be touched from this callback thread.
                // The UI timer recomputes the final geometry on the next tick.
            }
        }

        private bool TryGetGeometry(out OverlayGeometry geometry, out IntPtr ownerHwnd)
        {
            geometry = default;
            ownerHwnd = IntPtr.Zero;

            object window = null;
            object activeCell = null;
            object displayCell = null;
            object panes = null;
            object selectedPane = null;
            object visibleRange = null;

            try
            {
                dynamic app = _application;
                window = app.ActiveWindow;
                if (window == null) return false;

                dynamic dynWindow = window;
                ownerHwnd = new IntPtr(Convert.ToInt64(dynWindow.Hwnd));
                if (ownerHwnd == IntPtr.Zero ||
                    !NativeMethods.IsWindowVisible(ownerHwnd) ||
                    NativeMethods.IsIconic(ownerHwnd))
                    return false;

                // Find the real Excel worksheet grid window. This fixes two cases:
                // 1) Backstage/File screens: EXCEL7 is hidden or no longer focused, so hide overlay.
                // 2) Resized Excel windows: clip the overlay to the actual grid rectangle rather than
                //    the full widths/heights of VisibleRange cells.
                if (!TryGetVisibleExcelGrid(ownerHwnd, out var gridHwnd, out var gridRect))
                    return false;

                if (!IsWorksheetGridActive(ownerHwnd, gridHwnd))
                    return false;

                // Do not subtract scrollbar/tab sizes using fixed system metrics here.
                // EXCEL7 geometry differs across Excel/Office builds and can already exclude some
                // child UI. Subtracting fixed scrollbar metrics can therefore trim the viewport twice.
                // The real worksheet cell viewport is refined later with Window.RangeFromPoint,
                // which directly distinguishes on-screen cells from scrollbars, sheet tabs and splitters.

                activeCell = app.ActiveCell;
                if (activeCell == null) return false;

                displayCell = GetDisplayCell(activeCell);
                if (displayCell == null) return false;

                panes = dynWindow.Panes;
                if (panes == null) return false;

                dynamic dynPanes = panes;
                int paneCount = Convert.ToInt32(dynPanes.Count);
                if (paneCount < 1) return false;

                for (int i = 1; i <= paneCount; i++)
                {
                    object pane = null;
                    object vr = null;
                    try
                    {
                        // Late-bound COM default Item property. No Excel PIA needed.
                        pane = dynPanes.Item[i];
                        if (pane == null) continue;

                        dynamic dynPane = pane;
                        vr = dynPane.VisibleRange;
                        if (vr != null && RangeContainsCell(vr, displayCell))
                        {
                            selectedPane = pane;
                            pane = null; // ownership transferred
                            visibleRange = vr;
                            vr = null;
                            break;
                        }
                    }
                    finally
                    {
                        ReleaseCom(vr);
                        ReleaseCom(pane);
                    }
                }

                // Active cell can be scrolled fully out of view.
                if (selectedPane == null || visibleRange == null)
                    return false;

                dynamic cell = displayCell;
                dynamic vrange = visibleRange;
                dynamic paneForCoords = selectedPane;

                double cellLeftPt = Convert.ToDouble(cell.Left);
                double cellTopPt = Convert.ToDouble(cell.Top);
                double cellWidthPt = Convert.ToDouble(cell.Width);
                double cellHeightPt = Convert.ToDouble(cell.Height);

                int cellFirstRow = Convert.ToInt32(cell.Row);
                int cellFirstCol = Convert.ToInt32(cell.Column);
                int cellLastRow = cellFirstRow + Convert.ToInt32(cell.Rows.Count) - 1;
                int cellLastCol = cellFirstCol + Convert.ToInt32(cell.Columns.Count) - 1;

                double paneLeftPt = Convert.ToDouble(vrange.Left);
                double paneTopPt = Convert.ToDouble(vrange.Top);
                double paneRightPt = paneLeftPt + Convert.ToDouble(vrange.Width);
                double paneBottomPt = paneTopPt + Convert.ToDouble(vrange.Height);

                // PointsToScreenPixelsX/Y returns integer pixels. We still use it to obtain a
                // close screen estimate, but the active ROW's final top/bottom is no longer
                // derived from row-height points. Excel rasterizes row boundaries differently at
                // different Zoom levels, so the estimate can be a pixel or more off.
                double pixelsPerPointX = MeasurePixelsPerPointX(paneForCoords, cellLeftPt);
                double pixelsPerPointY = MeasurePixelsPerPointY(paneForCoords, cellTopPt);
                if (pixelsPerPointX <= 0 || pixelsPerPointY <= 0)
                    return false;

                uint dpi = NativeMethods.GetDpiForWindow(gridHwnd);
                if (dpi == 0) dpi = NativeMethods.GetDpiForWindow(ownerHwnd);
                double dpiScale = dpi > 0 ? dpi / 96.0 : 1.0;

                double cellLeftPx = PointsToScreenPixelsExactX(paneForCoords, cellLeftPt, pixelsPerPointX);
                double cellTopPx = PointsToScreenPixelsExactY(paneForCoords, cellTopPt, pixelsPerPointY);
                double cellRightPx = cellLeftPx + (cellWidthPt * pixelsPerPointX);
                double cellBottomPx = cellTopPx + (cellHeightPt * pixelsPerPointY);

                cellLeftPx = Math.Round(cellLeftPx);
                cellTopPx = Math.Round(cellTopPx);
                cellRightPx = Math.Round(cellRightPx);
                cellBottomPx = Math.Round(cellBottomPx);

                double paneLeftPx = PointsToScreenPixelsExactX(paneForCoords, paneLeftPt, pixelsPerPointX);
                double paneTopPx = PointsToScreenPixelsExactY(paneForCoords, paneTopPt, pixelsPerPointY);
                double paneRightPx = PointsToScreenPixelsExactX(paneForCoords, paneRightPt, pixelsPerPointX);
                double paneBottomPx = PointsToScreenPixelsExactY(paneForCoords, paneBottomPt, pixelsPerPointY);

                // VisibleRange contains whole worksheet cells and may extend beyond the visible
                // viewport. Clip to the real EXCEL7 grid, excluding scroll bars.
                paneLeftPx = Math.Max(paneLeftPx, gridRect.Left);
                paneTopPx = Math.Max(paneTopPx, gridRect.Top);
                paneRightPx = Math.Min(paneRightPx, gridRect.Right);
                paneBottomPx = Math.Min(paneBottomPx, gridRect.Bottom);

                // Refine the current pane viewport using actual on-screen cell hit-testing.
                // This removes the bottom workbook-tab strip, horizontal scrollbar, right vertical
                // scrollbar, row/column headers and pane splitters without relying on their pixel sizes.
                // It also avoids double-trimming on Office builds where EXCEL7 already excludes them.
                TryRefinePaneViewportFromHitTest(
                    (object)dynWindow,
                    visibleRange,
                    cellLeftPx, cellTopPx, cellRightPx, cellBottomPx,
                    gridRect,
                    ref paneLeftPx, ref paneTopPx, ref paneRightPx, ref paneBottomPx);

                // Refine row boundaries using Excel's actual on-screen hit-test.
                // Window.RangeFromPoint returns the Range at a pair of screen coordinates. We find
                // one pixel inside the selected row, then binary-search the first/last pixel that
                // still belongs to that row. This bypasses Zoom/DPI rounding in point math.
                //
                // RangeFromPoint has had DPI-scaling issues in some Excel builds, so this is a
                // best-effort refinement. If hit-testing fails, the previous estimate remains as a
                // safe fallback and Excel is never interrupted.
                var verticalKey = BuildVerticalBoundsCacheKey(
                    ownerHwnd, gridRect, dynWindow, paneForCoords,
                    cellFirstRow, cellLastRow, cellFirstCol, cellLastCol,
                    cellLeftPt, cellTopPt, cellWidthPt, cellHeightPt);

                if (string.Equals(verticalKey, _verticalBoundsCacheKey, StringComparison.Ordinal))
                {
                    cellTopPx = _verticalBoundsCacheTopPx;
                    cellBottomPx = _verticalBoundsCacheBottomPx;
                }
                else
                {
                    bool estimatedCellFullyVisible =
                        cellTopPx >= paneTopPx + 1 && cellBottomPx <= paneBottomPx - 1;

                    double actualTopPx;
                    double actualBottomPx;

                    if (estimatedCellFullyVisible &&
                        TryFindActualVerticalBounds(
                            (object)dynWindow,
                            cellLeftPx, cellRightPx,
                            cellTopPx, cellBottomPx,
                            paneLeftPx, paneTopPx, paneRightPx, paneBottomPx,
                            cellFirstRow, cellLastRow, cellFirstCol, cellLastCol,
                            out actualTopPx, out actualBottomPx))
                    {
                        cellTopPx = actualTopPx;
                        cellBottomPx = actualBottomPx;
                        _verticalBoundsCacheKey = verticalKey;
                        _verticalBoundsCacheTopPx = actualTopPx;
                        _verticalBoundsCacheBottomPx = actualBottomPx;
                    }
                    else
                    {
                        // Do not cache a failure. A subsequent timer tick may succeed after Excel
                        // finishes a zoom/redraw transition.
                        _verticalBoundsCacheKey = null;
                    }
                }

                double paneWidthPx = paneRightPx - paneLeftPx;
                double paneHeightPx = paneBottomPx - paneTopPx;
                double cellWidthPx = cellRightPx - cellLeftPx;
                double cellHeightPx = cellBottomPx - cellTopPx;

                if (paneWidthPx <= 0 || paneHeightPx <= 0 || cellWidthPx <= 0 || cellHeightPx <= 0)
                    return false;

                // If the active cell is not actually inside the current on-screen pane, do not draw.
                if (cellRightPx <= paneLeftPx || cellLeftPx >= paneRightPx ||
                    cellBottomPx <= paneTopPx || cellTopPx >= paneBottomPx)
                    return false;

                geometry = new OverlayGeometry(
                    paneLeftPx, paneTopPx, paneWidthPx, paneHeightPx,
                    cellLeftPx, cellTopPx, cellWidthPx, cellHeightPx,
                    dpiScale);

                return true;
            }
            catch (COMException)
            {
                return false;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return false;
            }
            finally
            {
                ReleaseCom(visibleRange);
                ReleaseCom(selectedPane);
                ReleaseCom(panes);
                if (!ReferenceEquals(displayCell, activeCell)) ReleaseCom(displayCell);
                ReleaseCom(activeCell);
                ReleaseCom(window);
            }
        }

        private static string BuildVerticalBoundsCacheKey(
            IntPtr ownerHwnd,
            NativeMethods.RECT gridRect,
            dynamic window,
            dynamic pane,
            int firstRow,
            int lastRow,
            int firstCol,
            int lastCol,
            double cellLeftPt,
            double cellTopPt,
            double cellWidthPt,
            double cellHeightPt)
        {
            object zoom = null;
            object scrollRow = null;
            object scrollColumn = null;
            try { zoom = window.Zoom; } catch { }
            try { scrollRow = pane.ScrollRow; } catch { }
            try { scrollColumn = pane.ScrollColumn; } catch { }

            return string.Join("|", new[]
            {
                ownerHwnd.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture),
                gridRect.Left.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gridRect.Top.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gridRect.Right.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gridRect.Bottom.ToString(System.Globalization.CultureInfo.InvariantCulture),
                firstRow.ToString(System.Globalization.CultureInfo.InvariantCulture),
                lastRow.ToString(System.Globalization.CultureInfo.InvariantCulture),
                firstCol.ToString(System.Globalization.CultureInfo.InvariantCulture),
                lastCol.ToString(System.Globalization.CultureInfo.InvariantCulture),
                cellLeftPt.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
                cellTopPt.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
                cellWidthPt.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
                cellHeightPt.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
                Convert.ToString(zoom, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(scrollRow, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(scrollColumn, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty
            });
        }

        private static bool TryFindActualVerticalBounds(
            object window,
            double estimatedCellLeftPx,
            double estimatedCellRightPx,
            double estimatedCellTopPx,
            double estimatedCellBottomPx,
            double paneLeftPx,
            double paneTopPx,
            double paneRightPx,
            double paneBottomPx,
            int firstRow,
            int lastRow,
            int firstCol,
            int lastCol,
            out double actualTopPx,
            out double actualBottomPx)
        {
            actualTopPx = estimatedCellTopPx;
            actualBottomPx = estimatedCellBottomPx;

            int paneTop = (int)Math.Ceiling(paneTopPx);
            int paneBottomExclusive = (int)Math.Floor(paneBottomPx);
            int paneLeft = (int)Math.Ceiling(paneLeftPx);
            int paneRightExclusive = (int)Math.Floor(paneRightPx);
            if (paneBottomExclusive - paneTop < 3 || paneRightExclusive - paneLeft < 3)
                return false;

            int estimatedLeft = (int)Math.Round(estimatedCellLeftPx);
            int estimatedRight = (int)Math.Round(estimatedCellRightPx);
            int estimatedTop = (int)Math.Round(estimatedCellTopPx);
            int estimatedBottom = (int)Math.Round(estimatedCellBottomPx);

            int[] xCandidates =
            {
                ClampInt((estimatedLeft + estimatedRight) / 2, paneLeft + 1, paneRightExclusive - 2),
                ClampInt(estimatedLeft + 2, paneLeft + 1, paneRightExclusive - 2),
                ClampInt(estimatedRight - 2, paneLeft + 1, paneRightExclusive - 2),
                ClampInt((3 * estimatedLeft + estimatedRight) / 4, paneLeft + 1, paneRightExclusive - 2),
                ClampInt((estimatedLeft + 3 * estimatedRight) / 4, paneLeft + 1, paneRightExclusive - 2)
            };

            int estimatedCenterY = ClampInt((estimatedTop + estimatedBottom) / 2, paneTop + 1, paneBottomExclusive - 2);
            int chosenX = 0;
            int insideY = 0;
            bool foundInside = false;

            foreach (int x in xCandidates)
            {
                if (TryFindInsideY(window, x, estimatedCenterY, paneTop, paneBottomExclusive,
                    firstRow, lastRow, firstCol, lastCol, out insideY))
                {
                    chosenX = x;
                    foundInside = true;
                    break;
                }
            }

            if (!foundInside)
                return false;

            // First target pixel: false ... false | true ... insideY
            int topPixel;
            if (IsTargetCellAtPoint(window, chosenX, paneTop, firstRow, lastRow, firstCol, lastCol))
            {
                topPixel = paneTop;
            }
            else
            {
                int lo = paneTop;      // known non-target
                int hi = insideY;      // known target
                while (hi - lo > 1)
                {
                    int mid = lo + ((hi - lo) / 2);
                    if (IsTargetCellAtPoint(window, chosenX, mid, firstRow, lastRow, firstCol, lastCol))
                        hi = mid;
                    else
                        lo = mid;
                }
                topPixel = hi;
            }

            // First non-target pixel after the row: insideY ... true | false ... bottom
            int lastVisibleY = paneBottomExclusive - 1;
            int bottomExclusive;
            if (IsTargetCellAtPoint(window, chosenX, lastVisibleY, firstRow, lastRow, firstCol, lastCol))
            {
                bottomExclusive = paneBottomExclusive;
            }
            else
            {
                int lo = insideY;      // known target
                int hi = lastVisibleY; // known non-target
                while (hi - lo > 1)
                {
                    int mid = lo + ((hi - lo) / 2);
                    if (IsTargetCellAtPoint(window, chosenX, mid, firstRow, lastRow, firstCol, lastCol))
                        lo = mid;
                    else
                        hi = mid;
                }
                bottomExclusive = hi;
            }

            if (bottomExclusive <= topPixel)
                return false;

            // Sanity check: hit-test result should be reasonably close to the point-based estimate.
            // This also rejects bad RangeFromPoint results caused by unusual DPI virtualization.
            double measuredHeight = bottomExclusive - topPixel;
            double estimatedHeight = Math.Max(1.0, estimatedCellBottomPx - estimatedCellTopPx);
            if (measuredHeight < estimatedHeight * 0.35 || measuredHeight > estimatedHeight * 2.8)
                return false;

            actualTopPx = topPixel;
            actualBottomPx = bottomExclusive;
            return true;
        }

        private static bool TryFindInsideY(
            dynamic window,
            int x,
            int estimatedCenterY,
            int paneTop,
            int paneBottomExclusive,
            int firstRow,
            int lastRow,
            int firstCol,
            int lastCol,
            out int insideY)
        {
            insideY = 0;
            int minY = paneTop + 1;
            int maxY = paneBottomExclusive - 2;
            if (maxY < minY) return false;

            int center = ClampInt(estimatedCenterY, minY, maxY);
            if (IsTargetCellAtPoint(window, x, center, firstRow, lastRow, firstCol, lastCol))
            {
                insideY = center;
                return true;
            }

            // The point estimate should already be close. Probe a small exponential set of
            // offsets rather than scanning every pixel; RangeFromPoint is a COM call and must stay
            // lightweight enough for interactive Excel use.
            int[] offsets = { 2, 4, 8, 16, 32, 64 };
            foreach (int delta in offsets)
            {
                int up = center - delta;
                if (up >= minY && IsTargetCellAtPoint(window, x, up, firstRow, lastRow, firstCol, lastCol))
                {
                    insideY = up;
                    return true;
                }

                int down = center + delta;
                if (down <= maxY && IsTargetCellAtPoint(window, x, down, firstRow, lastRow, firstCol, lastCol))
                {
                    insideY = down;
                    return true;
                }
            }

            return false;
        }

        private static bool IsTargetCellAtPoint(
            dynamic window,
            int x,
            int y,
            int firstRow,
            int lastRow,
            int firstCol,
            int lastCol)
        {
            object hit = null;
            try
            {
                hit = window.RangeFromPoint(x, y);
                if (hit == null) return false;

                dynamic range = hit;
                int row = Convert.ToInt32(range.Row);
                int column = Convert.ToInt32(range.Column);
                return row >= firstRow && row <= lastRow && column >= firstCol && column <= lastCol;
            }
            catch
            {
                // RangeFromPoint may return a Shape or fail while Excel is repainting.
                return false;
            }
            finally
            {
                ReleaseCom(hit);
            }
        }

        private static int ClampInt(int value, int min, int max)
        {
            if (max < min) return min;
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        private static double PointsToScreenPixelsExactX(dynamic pane, double points, double pixelsPerPoint)
        {
            int wholePoints = SafeTruncateToInt(points);
            double anchorPx = Convert.ToDouble(pane.PointsToScreenPixelsX(wholePoints));
            return anchorPx + ((points - wholePoints) * pixelsPerPoint);
        }

        private static double PointsToScreenPixelsExactY(dynamic pane, double points, double pixelsPerPoint)
        {
            int wholePoints = SafeTruncateToInt(points);
            double anchorPx = Convert.ToDouble(pane.PointsToScreenPixelsY(wholePoints));
            return anchorPx + ((points - wholePoints) * pixelsPerPoint);
        }

        private static double MeasurePixelsPerPointX(dynamic pane, double nearPoints)
        {
            return MeasurePixelsPerPoint(pane, nearPoints, true);
        }

        private static double MeasurePixelsPerPointY(dynamic pane, double nearPoints)
        {
            return MeasurePixelsPerPoint(pane, nearPoints, false);
        }

        private static double MeasurePixelsPerPoint(dynamic pane, double nearPoints, bool horizontal)
        {
            // Measure over a wide point interval to average away the integer-pixel return value
            // of PointsToScreenPixels*. 256 pt gives stable precision even at low zoom levels.
            const int SpanPoints = 256;
            int center = SafeTruncateToInt(nearPoints);
            int start = center - (SpanPoints / 2);
            int end = start + SpanPoints;

            // Worksheet coordinates are non-negative in normal use. Keep the sample in a safe
            // range and avoid integer overflow at extreme worksheet positions.
            if (start < 0)
            {
                start = 0;
                end = SpanPoints;
            }
            if (end < start || end > int.MaxValue - 1)
            {
                end = Math.Min(int.MaxValue - 1, center);
                start = Math.Max(0, end - SpanPoints);
            }
            if (end <= start)
                return 0;

            double p0 = horizontal
                ? Convert.ToDouble(pane.PointsToScreenPixelsX(start))
                : Convert.ToDouble(pane.PointsToScreenPixelsY(start));
            double p1 = horizontal
                ? Convert.ToDouble(pane.PointsToScreenPixelsX(end))
                : Convert.ToDouble(pane.PointsToScreenPixelsY(end));

            return Math.Abs((p1 - p0) / (end - start));
        }

        private static int SafeTruncateToInt(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                return 0;
            if (value <= int.MinValue) return int.MinValue;
            if (value >= int.MaxValue) return int.MaxValue;
            return (int)Math.Truncate(value);
        }

        private static bool TryGetVisibleExcelGrid(
            IntPtr ownerHwnd,
            out IntPtr gridHwnd,
            out NativeMethods.RECT gridRect)
        {
            gridHwnd = IntPtr.Zero;
            gridRect = default;

            long bestArea = 0;
            IntPtr bestHwnd = IntPtr.Zero;
            NativeMethods.RECT bestRect = default;

            try
            {
                NativeMethods.EnumChildWindows(ownerHwnd, (hwnd, lParam) =>
                {
                    try
                    {
                        if (!NativeMethods.IsWindowVisible(hwnd))
                            return true;

                        var className = new StringBuilder(128);
                        if (NativeMethods.GetClassName(hwnd, className, className.Capacity) <= 0)
                            return true;

                        // EXCEL7 is the worksheet grid view used by desktop Excel.
                        if (!string.Equals(className.ToString(), "EXCEL7", StringComparison.OrdinalIgnoreCase))
                            return true;

                        if (!NativeMethods.GetWindowRect(hwnd, out var rect))
                            return true;

                        if (rect.Width <= 0 || rect.Height <= 0)
                            return true;

                        long area = (long)rect.Width * rect.Height;
                        if (area > bestArea)
                        {
                            bestArea = area;
                            bestHwnd = hwnd;
                            bestRect = rect;
                        }
                    }
                    catch { }

                    return true;
                }, IntPtr.Zero);
            }
            catch
            {
                return false;
            }

            if (bestHwnd == IntPtr.Zero || bestArea <= 0)
                return false;

            gridHwnd = bestHwnd;
            gridRect = bestRect;
            return true;
        }

        private static bool IsWorksheetGridActive(IntPtr ownerHwnd, IntPtr gridHwnd)
        {
            if (ownerHwnd == IntPtr.Zero || gridHwnd == IntPtr.Zero || !NativeMethods.IsWindowVisible(gridHwnd))
                return false;

            try
            {
                // Query Excel's own GUI thread rather than the timer thread. In Backstage/File view
                // keyboard focus leaves the EXCEL7 worksheet grid, so the overlay is hidden.
                uint threadId = NativeMethods.GetWindowThreadProcessId(ownerHwnd, out _);
                if (threadId == 0)
                    return true;

                var info = new NativeMethods.GUITHREADINFO
                {
                    cbSize = (uint)Marshal.SizeOf(typeof(NativeMethods.GUITHREADINFO))
                };

                if (!NativeMethods.GetGUIThreadInfo(threadId, ref info))
                    return true;

                IntPtr focusedHwnd = info.hwndFocus;
                if (focusedHwnd == IntPtr.Zero)
                    return true;

                return focusedHwnd == gridHwnd || NativeMethods.IsChild(gridHwnd, focusedHwnd);
            }
            catch
            {
                return true;
            }
        }

        private static bool TryRefinePaneViewportFromHitTest(
            object window,
            object visibleRange,
            double estimatedCellLeftPx,
            double estimatedCellTopPx,
            double estimatedCellRightPx,
            double estimatedCellBottomPx,
            NativeMethods.RECT gridRect,
            ref double paneLeftPx,
            ref double paneTopPx,
            ref double paneRightPx,
            ref double paneBottomPx)
        {
            if (window == null || visibleRange == null)
                return false;

            int firstRow;
            int lastRow;
            int firstCol;
            int lastCol;

            try
            {
                dynamic vr = visibleRange;
                firstRow = Convert.ToInt32(vr.Row);
                firstCol = Convert.ToInt32(vr.Column);
                lastRow = firstRow + Convert.ToInt32(vr.Rows.Count) - 1;
                lastCol = firstCol + Convert.ToInt32(vr.Columns.Count) - 1;
            }
            catch
            {
                return false;
            }

            int minX = gridRect.Left;
            int maxX = gridRect.Right - 1;
            int minY = gridRect.Top;
            int maxY = gridRect.Bottom - 1;

            if (maxX - minX < 4 || maxY - minY < 4)
                return false;

            int estimatedCenterX = ClampInt(
                (int)Math.Round((estimatedCellLeftPx + estimatedCellRightPx) / 2.0),
                minX + 1, maxX - 1);
            int estimatedCenterY = ClampInt(
                (int)Math.Round((estimatedCellTopPx + estimatedCellBottomPx) / 2.0),
                minY + 1, maxY - 1);

            int insideX;
            int insideY;
            if (!TryFindPointInsideVisibleRange(
                    window,
                    estimatedCenterX,
                    estimatedCenterY,
                    minX, minY, maxX, maxY,
                    firstRow, lastRow, firstCol, lastCol,
                    out insideX, out insideY))
            {
                return false;
            }

            int left = FindFirstInsideX(
                window, minX, insideX, insideY,
                firstRow, lastRow, firstCol, lastCol);

            int rightExclusive = FindFirstOutsideXAfter(
                window, insideX, maxX, insideY,
                firstRow, lastRow, firstCol, lastCol);

            int top = FindFirstInsideY(
                window, minY, insideY, insideX,
                firstRow, lastRow, firstCol, lastCol);

            int bottomExclusive = FindFirstOutsideYAfter(
                window, insideY, maxY, insideX,
                firstRow, lastRow, firstCol, lastCol);

            if (rightExclusive <= left || bottomExclusive <= top)
                return false;

            double refinedLeft = Math.Max(paneLeftPx, left);
            double refinedTop = Math.Max(paneTopPx, top);
            double refinedRight = Math.Min(paneRightPx, rightExclusive);
            double refinedBottom = Math.Min(paneBottomPx, bottomExclusive);

            if (refinedRight - refinedLeft < 2 || refinedBottom - refinedTop < 2)
                return false;

            paneLeftPx = refinedLeft;
            paneTopPx = refinedTop;
            paneRightPx = refinedRight;
            paneBottomPx = refinedBottom;
            return true;
        }

        private static bool TryFindPointInsideVisibleRange(
            object window,
            int centerX,
            int centerY,
            int minX,
            int minY,
            int maxX,
            int maxY,
            int firstRow,
            int lastRow,
            int firstCol,
            int lastCol,
            out int insideX,
            out int insideY)
        {
            insideX = 0;
            insideY = 0;
            int[] offsets = { 0, 2, -2, 4, -4, 8, -8, 16, -16, 32, -32 };

            foreach (int dy in offsets)
            {
                int y = ClampInt(centerY + dy, minY + 1, maxY - 1);
                foreach (int dx in offsets)
                {
                    int x = ClampInt(centerX + dx, minX + 1, maxX - 1);
                    if (IsVisibleRangeAtPoint(window, x, y, firstRow, lastRow, firstCol, lastCol))
                    {
                        insideX = x;
                        insideY = y;
                        return true;
                    }
                }
            }
            return false;
        }

        private static int FindFirstInsideX(object window, int edgeX, int insideX, int y,
            int firstRow, int lastRow, int firstCol, int lastCol)
        {
            if (IsVisibleRangeAtPoint(window, edgeX, y, firstRow, lastRow, firstCol, lastCol))
                return edgeX;
            int lo = edgeX;
            int hi = insideX;
            while (hi - lo > 1)
            {
                int mid = lo + ((hi - lo) / 2);
                if (IsVisibleRangeAtPoint(window, mid, y, firstRow, lastRow, firstCol, lastCol)) hi = mid;
                else lo = mid;
            }
            return hi;
        }

        private static int FindFirstOutsideXAfter(object window, int insideX, int edgeX, int y,
            int firstRow, int lastRow, int firstCol, int lastCol)
        {
            if (IsVisibleRangeAtPoint(window, edgeX, y, firstRow, lastRow, firstCol, lastCol))
                return edgeX + 1;
            int lo = insideX;
            int hi = edgeX;
            while (hi - lo > 1)
            {
                int mid = lo + ((hi - lo) / 2);
                if (IsVisibleRangeAtPoint(window, mid, y, firstRow, lastRow, firstCol, lastCol)) lo = mid;
                else hi = mid;
            }
            return hi;
        }

        private static int FindFirstInsideY(object window, int edgeY, int insideY, int x,
            int firstRow, int lastRow, int firstCol, int lastCol)
        {
            if (IsVisibleRangeAtPoint(window, x, edgeY, firstRow, lastRow, firstCol, lastCol))
                return edgeY;
            int lo = edgeY;
            int hi = insideY;
            while (hi - lo > 1)
            {
                int mid = lo + ((hi - lo) / 2);
                if (IsVisibleRangeAtPoint(window, x, mid, firstRow, lastRow, firstCol, lastCol)) hi = mid;
                else lo = mid;
            }
            return hi;
        }

        private static int FindFirstOutsideYAfter(object window, int insideY, int edgeY, int x,
            int firstRow, int lastRow, int firstCol, int lastCol)
        {
            if (IsVisibleRangeAtPoint(window, x, edgeY, firstRow, lastRow, firstCol, lastCol))
                return edgeY + 1;
            int lo = insideY;
            int hi = edgeY;
            while (hi - lo > 1)
            {
                int mid = lo + ((hi - lo) / 2);
                if (IsVisibleRangeAtPoint(window, x, mid, firstRow, lastRow, firstCol, lastCol)) lo = mid;
                else hi = mid;
            }
            return hi;
        }

        private static bool IsVisibleRangeAtPoint(object window, int x, int y,
            int firstRow, int lastRow, int firstCol, int lastCol)
        {
            object hit = null;
            object topLeftCell = null;
            try
            {
                dynamic dynWindow = window;
                hit = dynWindow.RangeFromPoint(x, y);
                if (hit == null) return false;

                int row;
                int column;
                try
                {
                    dynamic range = hit;
                    row = Convert.ToInt32(range.Row);
                    column = Convert.ToInt32(range.Column);
                }
                catch
                {
                    try
                    {
                        dynamic shape = hit;
                        topLeftCell = shape.TopLeftCell;
                        if (topLeftCell == null) return false;
                        dynamic cell = topLeftCell;
                        row = Convert.ToInt32(cell.Row);
                        column = Convert.ToInt32(cell.Column);
                    }
                    catch
                    {
                        return false;
                    }
                }

                return row >= firstRow && row <= lastRow && column >= firstCol && column <= lastCol;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(topLeftCell);
                ReleaseCom(hit);
            }
        }

        private static object GetDisplayCell(object activeCell)
        {
            try
            {
                dynamic cell = activeCell;
                object merged = cell.MergeCells;
                if (Convert.ToBoolean(merged))
                    return cell.MergeArea;
            }
            catch { }

            return activeCell;
        }

        private static bool RangeContainsCell(object visibleRange, object cell)
        {
            dynamic vr = visibleRange;
            dynamic c = cell;

            int firstRow = Convert.ToInt32(vr.Row);
            int firstCol = Convert.ToInt32(vr.Column);
            int lastRow = firstRow + Convert.ToInt32(vr.Rows.Count) - 1;
            int lastCol = firstCol + Convert.ToInt32(vr.Columns.Count) - 1;

            int cellFirstRow = Convert.ToInt32(c.Row);
            int cellFirstCol = Convert.ToInt32(c.Column);
            int cellLastRow = cellFirstRow + Convert.ToInt32(c.Rows.Count) - 1;
            int cellLastCol = cellFirstCol + Convert.ToInt32(c.Columns.Count) - 1;

            return cellLastRow >= firstRow && cellFirstRow <= lastRow &&
                   cellLastCol >= firstCol && cellFirstCol <= lastCol;
        }

        private static void ReleaseCom(object obj)
        {
            if (obj == null) return;
            try
            {
                if (Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj);
            }
            catch { }
        }

        public void Dispose()
        {
            _timer.Stop();
            _timer.Tick -= Timer_Tick;
            _timer.Dispose();

            ReleaseMoveSizeHook();

            if (_overlay != null)
            {
                try { _overlay.Close(); } catch { }
                _overlay = null;
            }
        }
    }
}
