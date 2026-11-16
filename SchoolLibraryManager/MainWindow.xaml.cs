using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace SchoolLibraryManager;

public partial class MainWindow : Window
{
    private readonly DatabaseService _databaseService = new();
    private UpdateService? _updateService;

    private readonly DeskView _deskView;
    private readonly BooksView _booksView;
    private readonly StudentsView _studentsView;
    private readonly ReportsView _reportsView;
    private readonly ImportSyncView _importSyncView;

    public MainWindow()
    {
        InitializeComponent();

        var appVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        TxtVersion.Text = appVersion is null ? string.Empty : $"גרסה {appVersion.Major}.{appVersion.Minor}.{appVersion.Build}";

        _deskView = new DeskView(_databaseService);
        _booksView = new BooksView(_databaseService);
        _studentsView = new StudentsView(_databaseService);
        _reportsView = new ReportsView(_databaseService);
        _importSyncView = new ImportSyncView(_databaseService);

        try
        {
            var dbPath = _databaseService.DatabasePath;
            var lastWrite = System.IO.File.GetLastWriteTime(dbPath);
            Title = Title + " | " + dbPath + " | " + lastWrite.ToString("yyyy-MM-dd HH:mm");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "MainWindow title database path");
        }

        Loaded += OnLoaded;
        Closing += MainWindow_Closing;
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (App.IsMaintenance || App.IsRestoring)
        {
            e.Cancel = true;
            MessageBox.Show(this, "מתבצעת פעולת תחזוקה במסד הנתונים (דחיסה או שחזור). יש להמתין לסיומה ואז לסגור את התוכנה.", "פעולה מתבצעת", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    public void NavigateToDeskWithStudent(string stuId)
    {
        NavigateTo("Desk");
        _deskView.FocusStudentById(stuId);
    }

    public void NavigateToStudentsWithStudent(string stuId)
    {
        NavigateTo("Students");
        _studentsView.FocusStudentById(stuId);
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        NavigateTo("Desk");

        var backupWarning = DatabaseService.ConsumeBackupFailureWarning();
        if (!string.IsNullOrWhiteSpace(backupWarning))
        {
            MessageBox.Show(this,
                "אזהרה: הגיבוי האחרון של מסד הנתונים נכשל, ולכן אין כרגע גיבוי עדכני.\nמומלץ לפנות לגורם התומך, או לבדוק שיש מקום פנוי בדיסק ושהתיקייה Backups ניתנת לכתיבה.\n\nפרטים: " + backupWarning,
                "גיבוי נכשל",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        _updateService = new UpdateService(_databaseService, this);
        if (!App.IsMaintenance && !App.IsRestoring)
        {
            _ = _updateService.CheckAsync(interactive: false);
        }
    }

    private async void BtnCheckUpdates_Click(object sender, RoutedEventArgs e)
    {
        _updateService ??= new UpdateService(_databaseService, this);
        await _updateService.CheckAsync(interactive: true);
    }

    private void NavButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not ToggleButton { Tag: string viewKey })
        {
            return;
        }

        NavigateTo(viewKey);
    }

    private void NavigateTo(string viewKey)
    {
        MainContentHost.Content = viewKey switch
        {
            "Desk" => _deskView,
            "Books" => _booksView,
            "Students" => _studentsView,
            "Reports" => _reportsView,
            "Import" => _importSyncView,
            _ => _deskView
        };

        BtnNavDesk.IsChecked = viewKey == "Desk";
        BtnNavBooks.IsChecked = viewKey == "Books";
        BtnNavStudents.IsChecked = viewKey == "Students";
        BtnNavReports.IsChecked = viewKey == "Reports";
        BtnNavImport.IsChecked = viewKey == "Import";

        TxtShellStatus.Text = viewKey switch
        {
            "Desk" => "דלפק השאלות פעיל",
            "Books" => "ניהול כותרים ועותקים פעיל",
            "Students" => "ניהול חניכים ותשלומים פעיל",
            "Reports" => "מרכז דוחות פעיל",
            "Import" => "מרכז ייבוא וסנכרון פעיל",
            _ => "מוכן"
        };

        if (viewKey == "Desk")
        {
            txtShortcutsHelp.Text = "[F2] חניך | [F3] ברקוד ספר | [F4] החזרה מהירה | [Esc] איפוס דלפק";
            txtShortcutsHelp.Visibility = Visibility.Visible;
        }
        else
        {
            txtShortcutsHelp.Text = string.Empty;
            txtShortcutsHelp.Visibility = Visibility.Collapsed;
        }
    }

    public void SetMaintenanceMode(bool on)
    {
        App.IsMaintenance = on;
        BtnNavDesk.IsEnabled = !on;
        BtnNavBooks.IsEnabled = !on;
        BtnNavStudents.IsEnabled = !on;
        BtnNavReports.IsEnabled = !on;
        BtnNavImport.IsEnabled = !on;
    }

    private void BtnQuickBackup_Click(object sender, RoutedEventArgs e)
    {
        if (App.IsMaintenance)
        {
            return;
        }

        var backupPath = _databaseService.CreateBackup("Manual");
        if (string.IsNullOrEmpty(backupPath))
        {
            TxtShellStatus.Text = "גיבוי נכשל.";
            MessageBox.Show(this, "יצירת הגיבוי נכשלה. פרטים ביומן השגיאות.", "גיבוי", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        TxtShellStatus.Text = "גיבוי בוצע בהצלחה.";
        MessageBox.Show(this, $"הגיבוי נשמר בהצלחה:\n{System.IO.Path.GetFileName(backupPath)}", "גיבוי", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (App.IsMaintenance)
        {
            return;
        }

        if (_deskView.HandleHotkey(e.Key))
        {
            e.Handled = true;
        }
    }
}
