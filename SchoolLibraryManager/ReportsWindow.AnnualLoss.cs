using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace SchoolLibraryManager;

public partial class ReportsWindow
{
    private DataTable _annualLossTable = new();
    private TextBlock? _txtAnnualLossSummary;
    private TextBlock? _txtAnnualLossNote;
    private DataGrid? _gridAnnualLoss;

    private void BuildAnnualLossTab()
    {
        var generateButton = new Button
        {
            Content = "בחר שנה והפק דוח הפסדים שנתי",
            Width = 240,
            Height = 34,
            Background = (Brush)new BrushConverter().ConvertFromString("#B91C1C")!,
            Foreground = Brushes.White,
            FontWeight = FontWeights.SemiBold
        };
        generateButton.Click += BtnGenerateAnnualLoss_OnClick;

        var printButton = new Button
        {
            Content = "הדפס דוח 🖨️",
            Width = 140,
            Height = 34,
            Margin = new Thickness(8, 0, 0, 0)
        };
        printButton.Click += BtnPrintAnnualLoss_OnClick;

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
        buttons.Children.Add(generateButton);
        buttons.Children.Add(printButton);

        _txtAnnualLossSummary = new TextBlock
        {
            Text = "לחץ על \"בחר שנה והפק דוח הפסדים שנתי\" כדי לראות כמה ספרים שהושאלו בשנת הלימודים שתבחר טרם הוחזרו.",
            FontSize = 18,
            FontWeight = FontWeights.Bold,
            Foreground = (Brush)new BrushConverter().ConvertFromString("#991B1B")!,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 4)
        };

        _txtAnnualLossNote = new TextBlock
        {
            Foreground = Brushes.DimGray,
            FontSize = 12,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8)
        };

        var header = new StackPanel();
        header.Children.Add(buttons);
        header.Children.Add(_txtAnnualLossSummary);
        header.Children.Add(_txtAnnualLossNote);

        _gridAnnualLoss = new DataGrid
        {
            AutoGenerateColumns = false,
            IsReadOnly = true,
            CanUserAddRows = false,
            HeadersVisibility = DataGridHeadersVisibility.Column,
            GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
            RowHeaderWidth = 0
        };
        _gridAnnualLoss.Columns.Add(new DataGridTextColumn { Header = "ברקוד", Binding = new System.Windows.Data.Binding("Barcode"), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        _gridAnnualLoss.Columns.Add(new DataGridTextColumn { Header = "שם הספר", Binding = new System.Windows.Data.Binding("BookTitle"), Width = new DataGridLength(2, DataGridLengthUnitType.Star) });
        _gridAnnualLoss.Columns.Add(new DataGridTextColumn { Header = "שם החניך", Binding = new System.Windows.Data.Binding("StudentName"), Width = new DataGridLength(2, DataGridLengthUnitType.Star) });
        _gridAnnualLoss.Columns.Add(new DataGridTextColumn { Header = "תאריך השאלה", Binding = new System.Windows.Data.Binding("LoanDate"), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        _gridAnnualLoss.Columns.Add(new DataGridTextColumn { Header = "תאריך יעד", Binding = new System.Windows.Data.Binding("DueDate"), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        _gridAnnualLoss.Columns.Add(new DataGridTextColumn
        {
            Header = "עלות (₪)",
            Binding = new System.Windows.Data.Binding("Price") { StringFormat = "N2", ConverterCulture = CultureInfo.InvariantCulture },
            Width = new DataGridLength(1, DataGridLengthUnitType.Star)
        });

        var layout = new Grid { Margin = new Thickness(8) };
        layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        Grid.SetRow(header, 0);
        Grid.SetRow(_gridAnnualLoss, 1);
        layout.Children.Add(header);
        layout.Children.Add(_gridAnnualLoss);

        TabsReports.Items.Add(new TabItem { Header = "דוח הפסדים שנתי 📉", Content = layout });
    }

    private string? PromptSchoolYear(DependencyObject? anchor)
    {
        string currentYear;
        try
        {
            currentYear = _dbService.GetCurrentSchoolYear();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Annual loss: read current school year");
            currentYear = string.Empty;
        }

        var choices = DatabaseService.GetSchoolYearChoices(currentYear);
        try
        {
            var paymentYears = _dbService.GetPaymentSchoolYears();
            if (paymentYears.Count > 0)
            {
                choices = paymentYears;
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Annual loss: read payment school years");
        }

        var combo = new ComboBox
        {
            IsEditable = true,
            ItemsSource = choices,
            Height = 30,
            Margin = new Thickness(0, 8, 0, 12),
            Text = DatabaseService.TryParseHebrewYear(currentYear, out _) ? currentYear : choices.FirstOrDefault() ?? string.Empty
        };

        var okButton = new Button { Content = "הפק דוח", Width = 100, Height = 30, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
        var cancelButton = new Button { Content = "ביטול", Width = 100, Height = 30, IsCancel = true };
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Left };
        buttons.Children.Add(okButton);
        buttons.Children.Add(cancelButton);

        var panel = new StackPanel { Margin = new Thickness(16) };
        panel.Children.Add(new TextBlock
        {
            Text = "בחר את שנת הלימודים שעליה יופק דוח ההפסדים (לדוגמה: תשפז):",
            TextWrapping = TextWrapping.Wrap
        });
        panel.Children.Add(combo);
        panel.Children.Add(buttons);

        var dialog = new Window
        {
            Title = "בחירת שנה לדוח הפסדים",
            Content = panel,
            Width = 380,
            SizeToContent = SizeToContent.Height,
            ResizeMode = ResizeMode.NoResize,
            FlowDirection = FlowDirection.RightToLeft,
            ShowInTaskbar = false
        };

        var owner = anchor is null ? null : Window.GetWindow(anchor);
        if (owner is { IsVisible: true } && !ReferenceEquals(owner, dialog))
        {
            dialog.Owner = owner;
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
        }
        else
        {
            dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        okButton.Click += (_, _) =>
        {
            if (!DatabaseService.TryParseHebrewYear(combo.Text, out _))
            {
                MessageBox.Show("שנת לימודים לא תקינה. הזן שנה בפורמט כמו תשפז.", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            dialog.DialogResult = true;
        };

        return dialog.ShowDialog() == true ? combo.Text.Trim() : null;
    }

    private void BtnGenerateAnnualLoss_OnClick(object sender, RoutedEventArgs e)
    {
        var selectedYear = PromptSchoolYear(sender as DependencyObject);
        if (selectedYear is null)
        {
            return;
        }

        try
        {
            var summary = _dbService.GetAnnualLossSummary(selectedYear);
            _annualLossTable = summary.Details;
            if (_gridAnnualLoss is not null)
            {
                _gridAnnualLoss.ItemsSource = _annualLossTable.DefaultView;
            }

            if (_txtAnnualLossSummary is not null)
            {
                _txtAnnualLossSummary.Text =
                    $"כמות ספרים שטרם חזרו: {summary.UnreturnedCount:N0} | סך עלות מוערכת: {summary.TotalCost:N2} ₪";
            }

            if (_txtAnnualLossNote is not null)
            {
                var yearText = string.IsNullOrWhiteSpace(summary.SchoolYear) ? "(לא מוגדרת)" : summary.SchoolYear;
                var note = $"שנת לימודים {yearText}: השאלות פעילות מ-{summary.From:dd/MM/yyyy} עד {summary.ToExclusive.AddDays(-1):dd/MM/yyyy}.";
                if (summary.BooksWithoutPrice > 0)
                {
                    note += $" ל-{summary.BooksWithoutPrice} ספרים אין מחיר מוגדר, ולכן הם אינם נכללים בסכום.";
                }

                _txtAnnualLossNote.Text = note;
            }

            SetStatus("דוח הפסדים שנתי הופק.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Annual loss report failed");
            SetStatus("הפקת דוח הפסדים שנתי נכשלה.");
        }
    }

    private void BtnPrintAnnualLoss_OnClick(object sender, RoutedEventArgs e)
    {
        if (_annualLossTable.Rows.Count == 0)
        {
            SetStatus("אין נתונים להדפסה. יש להפיק את הדוח תחילה.");
            return;
        }

        PrintDataTable(
            "דוח הפסדים שנתי - ספרים שטרם חזרו",
            _annualLossTable,
            ["Barcode", "BookTitle", "StudentName", "LoanDate", "DueDate", "Price"],
            ["ברקוד", "שם הספר", "שם החניך", "תאריך השאלה", "תאריך יעד", "עלות (₪)"]);
    }
}
