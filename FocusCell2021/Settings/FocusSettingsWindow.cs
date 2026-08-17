using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using FocusCell2021.Localization;
using Forms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace FocusCell2021.Settings
{
    internal sealed class FocusSettingsWindow : Window
    {
        private readonly FocusSettings _settings;
        private readonly Action _applyCallback;
        private readonly IntPtr _ownerHwnd;

        private Button _colorButton;
        private Slider _opacitySlider;
        private TextBlock _opacityValue;
        private ComboBox _modeCombo;
        private CheckBox _showBorderCheck;
        private Slider _borderThicknessSlider;
        private TextBlock _borderThicknessValue;
        private Slider _borderOpacitySlider;
        private TextBlock _borderOpacityValue;
        private ComboBox _refreshCombo;
        private CheckBox _hideWhileMovingCheck;
        private string _selectedColor;

        public FocusSettingsWindow(FocusSettings settings, IntPtr ownerHwnd, Action applyCallback)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _ownerHwnd = ownerHwnd;
            _applyCallback = applyCallback;
            _selectedColor = settings.HighlightColor;

            Title = AppText.WindowTitle;
            Width = 430;
            Height = 560;
            MinWidth = 430;
            MinHeight = 560;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar = false;

            Content = BuildUi();
            if (_ownerHwnd != IntPtr.Zero)
                new WindowInteropHelper(this).Owner = _ownerHwnd;
            LoadFromSettings();
        }


        private UIElement BuildUi()
        {
            var root = new DockPanel { Margin = new Thickness(18) };

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 16, 0, 0)
            };

            var resetButton = MakeButton(AppText.Defaults, 76);
            resetButton.Click += (_, __) => ResetControls();
            buttons.Children.Add(resetButton);

            var applyButton = MakeButton(AppText.Apply, 76);
            applyButton.Margin = new Thickness(8, 0, 0, 0);
            applyButton.Click += (_, __) => ApplyToSettings();
            buttons.Children.Add(applyButton);

            var cancelButton = MakeButton(AppText.Cancel, 76);
            cancelButton.Margin = new Thickness(8, 0, 0, 0);
            cancelButton.Click += (_, __) => Close();
            buttons.Children.Add(cancelButton);

            var okButton = MakeButton(AppText.Ok, 76);
            okButton.Margin = new Thickness(8, 0, 0, 0);
            okButton.IsDefault = true;
            okButton.Click += (_, __) =>
            {
                ApplyToSettings();
                DialogResult = true;
                Close();
            };
            buttons.Children.Add(okButton);

            DockPanel.SetDock(buttons, Dock.Bottom);
            root.Children.Add(buttons);

            var content = new StackPanel();
            root.Children.Add(content);

            content.Children.Add(MakeHeading(AppText.HighlightHeading));

            var colorRow = MakeRow(AppText.HighlightColor);
            _colorButton = new Button
            {
                Width = 145,
                Height = 28,
                HorizontalAlignment = HorizontalAlignment.Left
            };
            _colorButton.Click += (_, __) => ChooseColor();
            Grid.SetColumn(_colorButton, 1);
            colorRow.Children.Add(_colorButton);
            content.Children.Add(colorRow);

            var opacityRow = MakeRow(AppText.HighlightOpacity);
            var opacityPanel = new DockPanel();
            _opacityValue = new TextBlock { Width = 52, TextAlignment = TextAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
            DockPanel.SetDock(_opacityValue, Dock.Right);
            opacityPanel.Children.Add(_opacityValue);
            _opacitySlider = new Slider { Minimum = 3, Maximum = 75, TickFrequency = 1, IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center };
            _opacitySlider.ValueChanged += (_, __) => _opacityValue.Text = $"{_opacitySlider.Value:0}%";
            opacityPanel.Children.Add(_opacitySlider);
            Grid.SetColumn(opacityPanel, 1);
            opacityRow.Children.Add(opacityPanel);
            content.Children.Add(opacityRow);

            var modeRow = MakeRow(AppText.HighlightRange);
            _modeCombo = new ComboBox { Width = 180, Height = 28, HorizontalAlignment = HorizontalAlignment.Left };
            _modeCombo.Items.Add(new ComboBoxItem { Content = AppText.RowAndColumn, Tag = FocusHighlightMode.RowAndColumn });
            _modeCombo.Items.Add(new ComboBoxItem { Content = AppText.RowOnly, Tag = FocusHighlightMode.RowOnly });
            _modeCombo.Items.Add(new ComboBoxItem { Content = AppText.ColumnOnly, Tag = FocusHighlightMode.ColumnOnly });
            Grid.SetColumn(_modeCombo, 1);
            modeRow.Children.Add(_modeCombo);
            content.Children.Add(modeRow);

            content.Children.Add(MakeSeparator());
            content.Children.Add(MakeHeading(AppText.BorderHeading));

            _showBorderCheck = new CheckBox { Content = AppText.ShowBorder, Margin = new Thickness(0, 4, 0, 8) };
            _showBorderCheck.Checked += (_, __) => UpdateBorderControlState();
            _showBorderCheck.Unchecked += (_, __) => UpdateBorderControlState();
            content.Children.Add(_showBorderCheck);

            var thicknessRow = MakeRow(AppText.BorderThickness);
            var thicknessPanel = new DockPanel();
            _borderThicknessValue = new TextBlock { Width = 52, TextAlignment = TextAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
            DockPanel.SetDock(_borderThicknessValue, Dock.Right);
            thicknessPanel.Children.Add(_borderThicknessValue);
            _borderThicknessSlider = new Slider { Minimum = 1, Maximum = 6, TickFrequency = 0.5, IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center };
            _borderThicknessSlider.ValueChanged += (_, __) => _borderThicknessValue.Text = $"{_borderThicknessSlider.Value:0.0}px";
            thicknessPanel.Children.Add(_borderThicknessSlider);
            Grid.SetColumn(thicknessPanel, 1);
            thicknessRow.Children.Add(thicknessPanel);
            content.Children.Add(thicknessRow);

            var borderOpacityRow = MakeRow(AppText.BorderOpacity);
            var borderOpacityPanel = new DockPanel();
            _borderOpacityValue = new TextBlock { Width = 52, TextAlignment = TextAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
            DockPanel.SetDock(_borderOpacityValue, Dock.Right);
            borderOpacityPanel.Children.Add(_borderOpacityValue);
            _borderOpacitySlider = new Slider { Minimum = 10, Maximum = 100, TickFrequency = 5, IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center };
            _borderOpacitySlider.ValueChanged += (_, __) => _borderOpacityValue.Text = $"{_borderOpacitySlider.Value:0}%";
            borderOpacityPanel.Children.Add(_borderOpacitySlider);
            Grid.SetColumn(borderOpacityPanel, 1);
            borderOpacityRow.Children.Add(borderOpacityPanel);
            content.Children.Add(borderOpacityRow);

            content.Children.Add(MakeSeparator());
            content.Children.Add(MakeHeading(AppText.BehaviorHeading));

            _hideWhileMovingCheck = new CheckBox
            {
                Content = AppText.HideWhileMoving,
                Margin = new Thickness(0, 4, 0, 10)
            };
            content.Children.Add(_hideWhileMovingCheck);

            var refreshRow = MakeRow(AppText.RefreshInterval);
            _refreshCombo = new ComboBox { Width = 180, Height = 28, HorizontalAlignment = HorizontalAlignment.Left };
            _refreshCombo.Items.Add(new ComboBoxItem { Content = AppText.RefreshFast, Tag = 35 });
            _refreshCombo.Items.Add(new ComboBoxItem { Content = AppText.RefreshDefault, Tag = 55 });
            _refreshCombo.Items.Add(new ComboBoxItem { Content = AppText.RefreshPowerSaving, Tag = 100 });
            _refreshCombo.Items.Add(new ComboBoxItem { Content = AppText.RefreshSlow, Tag = 150 });
            Grid.SetColumn(_refreshCombo, 1);
            refreshRow.Children.Add(_refreshCombo);
            content.Children.Add(refreshRow);

            var note = new TextBlock
            {
                Text = AppText.SettingsNote,
                Margin = new Thickness(0, 18, 0, 0),
                Foreground = Brushes.DimGray,
                TextWrapping = TextWrapping.Wrap
            };
            content.Children.Add(note);

            return root;
        }

        private static Button MakeButton(string text, double width)
        {
            return new Button { Content = text, Width = width, Height = 30 };
        }

        private static TextBlock MakeHeading(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontWeight = FontWeights.SemiBold,
                FontSize = 14,
                Margin = new Thickness(0, 2, 0, 8)
            };
        }

        private static Border MakeSeparator()
        {
            return new Border
            {
                Height = 1,
                Background = Brushes.LightGray,
                Margin = new Thickness(0, 14, 0, 14)
            };
        }

        private static Grid MakeRow(string label)
        {
            var grid = new Grid { Margin = new Thickness(0, 4, 0, 7) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(130) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var text = new TextBlock { Text = label, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(text, 0);
            grid.Children.Add(text);
            return grid;
        }

        private void LoadFromSettings()
        {
            _selectedColor = _settings.HighlightColor;
            UpdateColorButton();
            _opacitySlider.Value = Math.Round(_settings.Opacity * 100);
            SelectMode(_settings.Mode);
            _showBorderCheck.IsChecked = _settings.ShowCellBorder;
            _borderThicknessSlider.Value = _settings.BorderThickness;
            _borderOpacitySlider.Value = Math.Round(_settings.BorderOpacity * 100);
            _hideWhileMovingCheck.IsChecked = _settings.HideWhileMoving;
            SelectRefresh(_settings.RefreshIntervalMs);
            UpdateBorderControlState();
        }

        private void ResetControls()
        {
            _selectedColor = "#FFD54F";
            UpdateColorButton();
            _opacitySlider.Value = 18;
            SelectMode(FocusHighlightMode.RowAndColumn);
            _showBorderCheck.IsChecked = true;
            _borderThicknessSlider.Value = 2.0;
            _borderOpacitySlider.Value = 95;
            _hideWhileMovingCheck.IsChecked = true;
            SelectRefresh(55);
            UpdateBorderControlState();
        }

        private void ApplyToSettings()
        {
            _settings.HighlightColor = _selectedColor;
            _settings.Opacity = _opacitySlider.Value / 100.0;
            _settings.Mode = GetSelectedMode();
            _settings.ShowCellBorder = _showBorderCheck.IsChecked == true;
            _settings.BorderThickness = _borderThicknessSlider.Value;
            _settings.BorderOpacity = _borderOpacitySlider.Value / 100.0;
            _settings.HideWhileMoving = _hideWhileMovingCheck.IsChecked == true;
            _settings.RefreshIntervalMs = GetSelectedRefresh();
            _applyCallback?.Invoke();
        }

        private void ChooseColor()
        {
            try
            {
                var mediaColor = (Color)ColorConverter.ConvertFromString(_selectedColor);
                using (var dialog = new Forms.ColorDialog())
                {
                    dialog.FullOpen = true;
                    dialog.Color = Drawing.Color.FromArgb(mediaColor.R, mediaColor.G, mediaColor.B);
                    if (dialog.ShowDialog() == Forms.DialogResult.OK)
                    {
                        _selectedColor = $"#{dialog.Color.R:X2}{dialog.Color.G:X2}{dialog.Color.B:X2}";
                        UpdateColorButton();
                    }
                }
            }
            catch
            {
                _selectedColor = "#FFD54F";
                UpdateColorButton();
            }
        }

        private void UpdateColorButton()
        {
            _colorButton.Content = _selectedColor;
            try
            {
                var color = (Color)ColorConverter.ConvertFromString(_selectedColor);
                _colorButton.Background = new SolidColorBrush(color);
                var luminance = (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255.0;
                _colorButton.Foreground = luminance < 0.55 ? Brushes.White : Brushes.Black;
            }
            catch
            {
                _colorButton.Background = Brushes.Gold;
                _colorButton.Foreground = Brushes.Black;
            }
        }

        private void UpdateBorderControlState()
        {
            var enabled = _showBorderCheck.IsChecked == true;
            _borderThicknessSlider.IsEnabled = enabled;
            _borderOpacitySlider.IsEnabled = enabled;
        }

        private void SelectMode(FocusHighlightMode mode)
        {
            foreach (ComboBoxItem item in _modeCombo.Items)
            {
                if (item.Tag is FocusHighlightMode value && value == mode)
                {
                    _modeCombo.SelectedItem = item;
                    return;
                }
            }
            _modeCombo.SelectedIndex = 0;
        }

        private FocusHighlightMode GetSelectedMode()
        {
            if (_modeCombo.SelectedItem is ComboBoxItem item && item.Tag is FocusHighlightMode mode)
                return mode;
            return FocusHighlightMode.RowAndColumn;
        }

        private void SelectRefresh(int value)
        {
            var bestIndex = 1;
            var bestDistance = int.MaxValue;
            for (int i = 0; i < _refreshCombo.Items.Count; i++)
            {
                var item = (ComboBoxItem)_refreshCombo.Items[i];
                var v = Convert.ToInt32(item.Tag, CultureInfo.InvariantCulture);
                var distance = Math.Abs(v - value);
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    bestIndex = i;
                }
            }
            _refreshCombo.SelectedIndex = bestIndex;
        }

        private int GetSelectedRefresh()
        {
            if (_refreshCombo.SelectedItem is ComboBoxItem item)
                return Convert.ToInt32(item.Tag, CultureInfo.InvariantCulture);
            return 55;
        }
    }
}
