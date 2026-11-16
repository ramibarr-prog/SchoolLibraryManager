using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace SchoolLibraryManager;

public partial class ReportsWindow : Window
{
    private readonly DatabaseService _dbService;
    private DataTable _overdueTable = new();
    private DataTable _unpaidTable = new();
    private DataTable _activeLoansTable = new();

    public ReportsWindow(DatabaseService dbService)
    {
        InitializeComponent();
        _dbService = dbService;
        BuildAnnualLossTab();
        Loaded += ReportsWindow_OnLoaded;
    }

    public void ReloadAll()
    {
        ReportsWindow_OnLoaded(this, new RoutedEventArgs());
    }

    private void ReportsWindow_OnLoaded(object sender, RoutedEventArgs e)
    {
        try
        {
            try { LoadClassFilters(); } catch (Exception ex) { ErrorLogger.Log(ex, "Load class filters failed"); }
            try { LoadInventoryReports(); } catch (Exception ex) { ErrorLogger.Log(ex, "Load inventory failed"); }
            LoadOverdueReport();
            LoadUnpaidStudentsReport();
            LoadActiveLoansReport();
            SetStatus("הדוחות נטענו בהצלחה.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Reports window load failed");
            SetStatus("טעינת הדוחות נכשלה.");
        }
    }

    private void LoadClassFilters()
    {
        var classes = _dbService.GetClasses();
        var items = new List<ClassFilterItem> { new("הכל", "הכל") };
        foreach (DataRow row in classes.Rows)
        {
            var id = row["ClassID"]?.ToString()?.Trim() ?? string.Empty;
            if (id.Length == 0)
            {
                continue;
            }

            var name = row["ClassName"]?.ToString()?.Trim();
            items.Add(new ClassFilterItem(id, string.IsNullOrWhiteSpace(name) ? ClassFormatter.ToDisplay(id) : name));
        }

        CmbOverdueClass.ItemsSource = items.ToList();
        CmbOverdueClass.SelectedValue = "הכל";

        CmbUnpaidClass.ItemsSource = items.ToList();
        CmbUnpaidClass.SelectedValue = "הכל";

        CmbActiveLoansClass.ItemsSource = items.ToList();
        CmbActiveLoansClass.SelectedValue = "הכל";
    }

    private void BtnRefreshActiveLoans_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            LoadActiveLoansReport();
            SetStatus("דוח השאלות פעילות עודכן.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Refresh active loans report failed");
            SetStatus("רענון דוח השאלות פעילות נכשל.");
        }
    }

    private void LoadActiveLoansReport()
    {
        var selectedClass = CmbActiveLoansClass.SelectedValue?.ToString();
        _activeLoansTable = _dbService.GetActiveLoansReport(selectedClass);
        GridActiveLoans.ItemsSource = _activeLoansTable.DefaultView;
    }

    private void BtnPrintActiveLoans_OnClick(object sender, RoutedEventArgs e)
    {
        if (_activeLoansTable.Rows.Count == 0)
        {
            SetStatus("אין נתונים להדפסת דוח השאלות פעילות.");
            return;
        }

        PrintDataTable(
            "דוח השאלות פעילות",
            _activeLoansTable,
            ["ClassName", "StudentName", "stuPhone", "BookTitle", "Barcode", "LoanDate", "DueDate"],
            ["כיתה", "חניך", "טלפון", "שם הספר", "ברקוד", "תאריך השאלה", "תאריך החזרה"]);
    }

    private void BtnRefreshOverdue_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            LoadOverdueReport();
            SetStatus("דוח האיחורים עודכן.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Refresh overdue report failed");
            SetStatus("רענון דוח איחורים נכשל.");
        }
    }

    private void LoadOverdueReport()
    {
        var selectedClass = CmbOverdueClass.SelectedValue?.ToString();
        _overdueTable = _dbService.GetOverdueLoans(selectedClass);
        GridOverdue.ItemsSource = _overdueTable.DefaultView;
    }

    private void LoadInventoryReports()
    {
        var overview = _dbService.GetInventoryOverview();
        TxtKpiTotalTitles.Text = overview.TotalTitles.ToString("N0");
        TxtKpiTotalCopies.Text = overview.TotalCopies.ToString("N0");
        TxtKpiActiveLoans.Text = overview.ActiveLoans.ToString("N0");
        TxtKpiAvailableCopies.Text = overview.AvailableCopies.ToString("N0");

        var bySubject = _dbService.GetInventoryBySubject();
        GridInventoryBySubject.ItemsSource = bySubject.DefaultView;
    }

    private void BtnRefreshUnpaid_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            LoadUnpaidStudentsReport();
            SetStatus("דוח הגבייה עודכן.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Refresh unpaid students report failed");
            SetStatus("רענון דוח גבייה נכשל.");
        }
    }

    private void LoadUnpaidStudentsReport()
    {
        var selectedClass = CmbUnpaidClass.SelectedValue?.ToString();
        _unpaidTable = _dbService.GetUnpaidStudentsReport(selectedClass);
        EnsureFullNameColumn(_unpaidTable);

        GridUnpaidStudents.ItemsSource = _unpaidTable.DefaultView;
        string yearText;
        try
        {
            yearText = _dbService.GetCurrentSchoolYear();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Reports school year title");
            yearText = "(לא מוגדרת)";
        }

        TxtUnpaidTitle.Text = $"חניכים שטרם שילמו דמי ספרייה לשנת {yearText}";
    }

    private static void EnsureFullNameColumn(DataTable dt)
    {
        if (!dt.Columns.Contains("FullName"))
        {
            dt.Columns.Add("FullName", typeof(string));
        }

        foreach (DataRow row in dt.Rows)
        {
            var firstName = row.Table.Columns.Contains("stuFname") ? row["stuFname"]?.ToString() : string.Empty;
            var lastName = row.Table.Columns.Contains("stuLname") ? row["stuLname"]?.ToString() : string.Empty;
            row["FullName"] = $"{firstName} {lastName}".Trim();
        }
    }

    private void BtnPrintOverdue_OnClick(object sender, RoutedEventArgs e)
    {
        if (_overdueTable.Rows.Count == 0)
        {
            SetStatus("אין נתונים להדפסת דוח איחורים.");
            return;
        }

        PrintDataTable(
            "דוח ספרים באיחור",
            _overdueTable,
            ["Barcode", "BookTitle", "StudentName", "ClassName", "stuPhone", "LoanDate", "DueDate", "DaysOverdue"],
            ["ברקוד", "שם הספר", "חניך", "כיתה", "טלפון", "תאריך השאלה", "תאריך יעד", "ימי איחור"]);
    }

    private void BtnPrintUnpaid_OnClick(object sender, RoutedEventArgs e)
    {
        if (_unpaidTable.Rows.Count == 0)
        {
            SetStatus("אין נתונים להדפסת רשימת חובות.");
            return;
        }

        PrintDataTable(
            "רשימת חניכים שטרם שילמו",
            _unpaidTable,
            ["stuID", "FullName", "ClassName", "stuPhone", "ActiveLoansCount"],
            ["ת.ז", "שם מלא", "כיתה", "טלפון", "השאלות פעילות"]);
    }

    private void PrintDataTable(string title, DataTable data, string[] columns, string[] headers)
    {
        var printDialog = new PrintDialog();
        if (printDialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var document = BuildDocument(title, data, columns, headers);

            const double marginPx = 60;
            var pageWidth = printDialog.PrintableAreaWidth > 0 ? printDialog.PrintableAreaWidth : 793;
            var pageHeight = printDialog.PrintableAreaHeight > 0 ? printDialog.PrintableAreaHeight : 1122;
            document.PageWidth = pageWidth;
            document.PageHeight = pageHeight;
            document.ColumnWidth = pageWidth;
            document.PagePadding = new Thickness(marginPx);

            var paginator = ((IDocumentPaginatorSource)document).DocumentPaginator;
            paginator.PageSize = new Size(pageWidth, pageHeight);
            printDialog.PrintDocument(paginator, title);
            SetStatus($"הדוח '{title}' נשלח להדפסה.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"Print report failed: {title}");
            SetStatus("הדפסת הדוח נכשלה.");
        }
    }

    private static FlowDocument BuildDocument(string title, DataTable data, string[] columns, string[] headers)
    {
        var document = new FlowDocument
        {
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 11,
            PagePadding = new Thickness(30),
            FlowDirection = FlowDirection.RightToLeft
        };

        document.Blocks.Add(new Paragraph(new Run(title))
        {
            FontSize = 18,
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(0, 0, 0, 12)
        });

        var table = new Table { CellSpacing = 0 };
        foreach (var column in columns)
        {
            table.Columns.Add(new TableColumn { Width = new GridLength(GetPrintColumnWeight(column), GridUnitType.Star) });
        }

        var rowGroup = new TableRowGroup();
        table.RowGroups.Add(rowGroup);

        var headerRow = new TableRow();
        for (var i = 0; i < headers.Length; i++)
        {
            headerRow.Cells.Add(new TableCell(new Paragraph(new Run(headers[i]))) { FontWeight = FontWeights.Bold });
        }
        rowGroup.Rows.Add(headerRow);

        foreach (DataRow row in data.Rows)
        {
            var dataRow = new TableRow();
            foreach (var column in columns)
            {
                var value = row.Table.Columns.Contains(column) ? row[column]?.ToString() : string.Empty;
                dataRow.Cells.Add(new TableCell(new Paragraph(new Run(value ?? string.Empty))));
            }

            rowGroup.Rows.Add(dataRow);
        }

        document.Blocks.Add(table);
        SchoolBranding.AddLetterhead(document);
        return document;
    }

    private static double GetPrintColumnWeight(string column) => column switch
    {
        "BookTitle" => 3.2,
        "StudentName" or "FullName" => 2.6,
        "Barcode" => 1.2,
        "stuID" => 1.3,
        "ClassName" => 1.1,
        "stuPhone" => 1.5,
        "LoanDate" or "DueDate" => 1.4,
        "DaysOverdue" or "ActiveLoansCount" => 1.0,
        _ => 1.5
    };

    private void SetStatus(string message)
    {
        TxtStatus.Text = message;
    }

}

public sealed record ClassFilterItem(string ClassID, string ClassName);
