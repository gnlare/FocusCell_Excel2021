namespace FocusCell2021.Overlay
{
    internal readonly struct OverlayGeometry
    {
        public OverlayGeometry(
            double paneLeft, double paneTop, double paneWidth, double paneHeight,
            double cellLeft, double cellTop, double cellWidth, double cellHeight,
            double dpiScale)
        {
            PaneLeft = paneLeft;
            PaneTop = paneTop;
            PaneWidth = paneWidth;
            PaneHeight = paneHeight;
            CellLeft = cellLeft;
            CellTop = cellTop;
            CellWidth = cellWidth;
            CellHeight = cellHeight;
            DpiScale = dpiScale;
        }

        public double PaneLeft { get; }
        public double PaneTop { get; }
        public double PaneWidth { get; }
        public double PaneHeight { get; }
        public double CellLeft { get; }
        public double CellTop { get; }
        public double CellWidth { get; }
        public double CellHeight { get; }
        public double DpiScale { get; }
    }
}
