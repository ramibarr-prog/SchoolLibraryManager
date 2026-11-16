using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.Media;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace SchoolLibraryManager;

public partial class DeskView : UserControl
{
    private readonly DatabaseService _databaseService;
    private Student? _currentStudent;
    private bool _currentStudentPaidForYear;
    private bool _quickReturnMode;
    private readonly ObservableCollection<ReturnedBookItem> _returnedThisSession = new();
    private readonly ObservableCollection<BorrowCartItem> _borrowCart = new();
    private readonly ObservableCollection<ReturnCartItem> _returnCart = new();
    private readonly HashSet<string> _borrowCartLookup = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _returnCartLookup = new(StringComparer.OrdinalIgnoreCase);
    private readonly ObservableCollection<StudentSuggestionItem> _studentSuggestions = new();
    private bool _suppressStudentSearchEvents;
    private string? _pendingExternalStudentId;
    private readonly bool _currentStudentHasOverdueLoans;

    public DeskView(DatabaseService databaseService)
    {
        InitializeComponent();
        _databaseService = databaseService;
        Loaded += OnLoaded;
    }

    public void FocusStudentById(string stuId)
    {
        var normalizedId = (stuId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedId))
        {
            txtStudentSearch.Focus();
            return;
        }

        _pendingExternalStudentId = normalizedId;

        if (!IsLoaded)
        {
            return;
        }

        RunDeskStudentSearch(normalizedId);
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        ResetPaymentDisplay();
        gridSessionReturns.ItemsSource = _returnedThisSession;
        dgBorrowCart.ItemsSource = _borrowCart;
        dgReturnCart.ItemsSource = _returnCart;
        lstStudentSuggestions.ItemsSource = _studentSuggestions;
        SetQuickReturnMode(false);

        if (!string.IsNullOrWhiteSpace(_pendingExternalStudentId))
        {
            var stuId = _pendingExternalStudentId;
            _pendingExternalStudentId = null;
            RunDeskStudentSearch(stuId);
            return;
        }

        txtStudentSearch.Focus();
    }

    private void RunDeskStudentSearch(string? stuId)
    {
        var normalizedId = (stuId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedId))
        {
            return;
        }

        _suppressStudentSearchEvents = true;
        txtStudentSearch.Text = normalizedId;
        _suppressStudentSearchEvents = false;

        SearchStudent();
    }

    private void RefreshStudentSuggestions()
    {
        if (_suppressStudentSearchEvents)
        {
            return;
        }

        var query = (txtStudentSearch.Text ?? string.Empty).Trim();
        _studentSuggestions.Clear();
        if (string.IsNullOrWhiteSpace(query))
        {
            popupStudentSuggestions.IsOpen = false;
            return;
        }

        var dt = _databaseService.SearchStudents(query);
        foreach (DataRow row in dt.Rows.Cast<DataRow>()
                     .Where(r => !string.Equals((r["stuClassID"]?.ToString() ?? string.Empty).Trim(), "9999", StringComparison.OrdinalIgnoreCase))
                     .Take(10))
        {
            var stuId = row["stuID"]?.ToString() ?? string.Empty;
            var firstName = row["stuFname"]?.ToString() ?? string.Empty;
            var lastName = row["stuLname"]?.ToString() ?? string.Empty;
            var classCode = row["stuClassID"]?.ToString() ?? string.Empty;
            var className = dt.Columns.Contains("ClassName") ? row["ClassName"]?.ToString() ?? string.Empty : string.Empty;
            _studentSuggestions.Add(new StudentSuggestionItem
            {
                StuId = stuId,
                FirstName = firstName,
                LastName = lastName,
                ClassCode = classCode,
                DisplayText = $"{firstName} {lastName}".Trim() + $" - כיתה {ClassDisplay(className, classCode)} (ת.ז: {stuId})"
            });
        }

        popupStudentSuggestions.IsOpen = _studentSuggestions.Count > 0;
        lstStudentSuggestions.SelectedIndex = -1;
    }

    private static bool IsDescendantOf(DependencyObject? child, DependencyObject? parent)
    {
        while (child is not null)
        {
            if (ReferenceEquals(child, parent))
            {
                return true;
            }

            child = child switch
            {
                Visual or Visual3D => VisualTreeHelper.GetParent(child),
                _ => null
            };
        }

        return false;
    }

    private void DeskView_OnPreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (!popupStudentSuggestions.IsOpen)
        {
            return;
        }

        if (e.OriginalSource is DependencyObject source &&
            (IsDescendantOf(source, txtStudentSearch) || IsDescendantOf(source, lstStudentSuggestions)))
        {
            return;
        }

        CloseStudentSuggestions();
    }

    private void SelectStudentSuggestion(StudentSuggestionItem suggestion)
    {
        var student = _databaseService.FindStudent(suggestion.StuId);
        if (student is null)
        {
            SetStatus("לא נמצא חניך לפי החיפוש.", StatusKind.Error);
            SystemSounds.Hand.Play();
            return;
        }

        LoadStudentCard(student);
        popupStudentSuggestions.IsOpen = false;
        _studentSuggestions.Clear();
        SetStatus($"נטען חניך: {student.FullName}", StatusKind.Success);
        SystemSounds.Asterisk.Play();
        txtBorrowBarcode.Focus();
        txtBorrowBarcode.SelectAll();
    }

    private void CloseStudentSuggestions()
    {
        popupStudentSuggestions.IsOpen = false;
        _studentSuggestions.Clear();
    }

    private bool TryLoadStudentByDirectId(string input)
    {
        var normalized = (input ?? string.Empty).Trim();
        if (normalized.Length < 8 || normalized.Any(ch => !char.IsDigit(ch)))
        {
            return false;
        }

        var student = _databaseService.FindStudent(normalized);
        if (student is null)
        {
            return false;
        }

        LoadStudentCard(student);
        CloseStudentSuggestions();
        SetStatus($"נטען חניך: {student.FullName}", StatusKind.Success);
        SystemSounds.Asterisk.Play();
        txtBorrowBarcode.Focus();
        txtBorrowBarcode.SelectAll();
        return true;
    }

    public bool HandleHotkey(Key key)
    {
        switch (key)
        {
            case Key.F2:
                ExitQuickReturnMode();
                txtStudentSearch.Focus();
                txtStudentSearch.SelectAll();
                return true;
            case Key.F3:
                txtBorrowBarcode.Focus();
                txtBorrowBarcode.SelectAll();
                return true;
            case Key.F4:
                ActivateQuickReturnMode();
                return true;
            case Key.Escape:
                ExitQuickReturnMode();
                return true;
            default:
                return false;
        }
    }

    private void ActivateQuickReturnMode()
    {
        ResetStudentSelection();
        SetQuickReturnMode(true);
        txtQuickReturnBarcode.Focus();
        txtQuickReturnBarcode.SelectAll();
        SetStatus("מצב החזרות פעיל - סרוק ספרים ברצף.", StatusKind.Info);
    }

    private void ExitQuickReturnMode()
    {
        SetQuickReturnMode(false);
        txtQuickReturnBarcode.Clear();
        txtStudentSearch.Focus();
        txtStudentSearch.SelectAll();
    }

    private void SetQuickReturnMode(bool enabled)
    {
        _quickReturnMode = enabled;
        if (tabDeskActions != null)
        {
            if (enabled && tabDeskActions.SelectedIndex != 2)
            {
                tabDeskActions.SelectedIndex = 2;
            }
            else if (!enabled && tabDeskActions.SelectedIndex == 2)
            {
                tabDeskActions.SelectedIndex = 0;
            }
        }

        studentCardPanel.Opacity = enabled ? 0.45 : 1.0;
        studentCardPanel.IsEnabled = !enabled;
        txtStudentSearch.Opacity = enabled ? 0.45 : 1.0;

        dgActiveLoans.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        txtActiveLoansHeader.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        gridSessionReturns.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        txtQuickReturnLogHeader.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;

        if (enabled)
        {
            txtQuickReturnBarcode.Focus();
            txtQuickReturnBarcode.SelectAll();
        }
    }

    private void AddQuickReturnLog(string barcode, string title, string studentName, string className)
    {
        _returnedThisSession.Insert(0, new ReturnedBookItem
        {
            ReturnTime = DateTime.Now.ToString("HH:mm:ss"),
            Barcode = barcode,
            Title = string.IsNullOrWhiteSpace(title) ? "ללא כותרת" : title,
            StudentName = string.IsNullOrWhiteSpace(studentName) ? "חניך לא ידוע" : studentName,
            ClassName = string.IsNullOrWhiteSpace(className) ? "-" : ClassFormatter.ToDisplayName(className, className)
        });

        while (_returnedThisSession.Count > 120)
        {
            _returnedThisSession.RemoveAt(_returnedThisSession.Count - 1);
        }
    }

    private (string Title, string BorrowerName, string ClassDisplay) LookupBorrowedCopyDetails(string barcode)
    {
        using var connection = _databaseService.CreateOpenConnection();
        var sql = @"SELECT TOP 1 IIF(b.bkName IS NULL, '', b.bkName) AS BookTitle,
                           Trim(s.stuFname & ' ' & s.stuLname) AS StudentName,
                           IIF(c.ClassName IS NULL, s.stuClassID, c.ClassName) AS ClassName
                    FROM (([Transaction] t
                    LEFT JOIN [BookCopy] bc ON t.Barcode = bc.Barcode)
                    LEFT JOIN [Book] b ON bc.bkID = b.bkID)
                    LEFT JOIN ([Student] s LEFT JOIN [Classes] c ON s.stuClassID = c.ClassID) ON t.stuID = s.stuID
                    WHERE t.Barcode = ? AND t.History = '0'";

        using var cmd = new OleDbCommand(sql, connection);
        cmd.Parameters.AddWithValue("@p1", barcode);
        using var reader = cmd.ExecuteReader();
        if (reader is null || !reader.Read())
        {
            return (string.Empty, string.Empty, string.Empty);
        }

        return (
            (reader["BookTitle"]?.ToString() ?? string.Empty).Trim(),
            (reader["StudentName"]?.ToString() ?? string.Empty).Trim(),
            (reader["ClassName"]?.ToString() ?? string.Empty).Trim());
    }

    private void BtnSearchStudent_OnClick(object sender, RoutedEventArgs e) => SearchStudent();

    private void TabDeskActions_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!ReferenceEquals(e.Source, tabDeskActions))
        {
            return;
        }

        var isQuickReturnTab = tabDeskActions.SelectedIndex == 2;
        if (isQuickReturnTab == _quickReturnMode)
        {
            return;
        }

        if (isQuickReturnTab)
        {
            ActivateQuickReturnMode();
        }
        else
        {
            ExitQuickReturnMode();
        }
    }

    private void TxtStudentSearch_OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            CloseStudentSuggestions();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Down && popupStudentSuggestions.IsOpen && _studentSuggestions.Count > 0)
        {
            lstStudentSuggestions.SelectedIndex = 0;
            lstStudentSuggestions.Focus();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Enter)
        {
            if (popupStudentSuggestions.IsOpen && lstStudentSuggestions.SelectedItem is StudentSuggestionItem selected)
            {
                SelectStudentSuggestion(selected);
                e.Handled = true;
                return;
            }

            if (TryLoadStudentByDirectId(txtStudentSearch.Text))
            {
                e.Handled = true;
                return;
            }

            SearchStudent();
            e.Handled = true;
        }
    }

    private void TxtStudentSearch_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        RefreshStudentSuggestions();
    }

    private void LstStudentSuggestions_OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (lstStudentSuggestions.SelectedItem is StudentSuggestionItem selected)
        {
            SelectStudentSuggestion(selected);
        }
    }

    private void LstStudentSuggestions_OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && lstStudentSuggestions.SelectedItem is StudentSuggestionItem selected)
        {
            SelectStudentSuggestion(selected);
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Escape)
        {
            CloseStudentSuggestions();
            txtStudentSearch.Focus();
            txtStudentSearch.SelectAll();
            e.Handled = true;
        }
    }

    private void BtnBorrow_OnClick(object sender, RoutedEventArgs e) => AddBorrowToCart();

    private void TxtBorrowBarcode_OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            AddBorrowToCart();
            e.Handled = true;
        }
    }

    private void BtnReturnSingle_OnClick(object sender, RoutedEventArgs e) => AddReturnToCart(txtManualReturnBarcode.Text);

    private void TxtManualReturnBarcode_OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            AddReturnToCart(txtManualReturnBarcode.Text);
            e.Handled = true;
        }
    }

    private void BtnEnterQuickReturn_OnClick(object sender, RoutedEventArgs e) => ActivateQuickReturnMode();

    private void BtnReturn_OnClick(object sender, RoutedEventArgs e) => ReturnCopy(txtQuickReturnBarcode.Text, forceContinuousMode: true);

    private void TxtQuickReturnBarcode_OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            ReturnCopy(txtQuickReturnBarcode.Text, forceContinuousMode: true);
            e.Handled = true;
        }
    }

    private void BtnExitQuickReturn_Click(object sender, RoutedEventArgs e) => ExitQuickReturnMode();

    private void BtnBorrowBatch_OnClick(object sender, RoutedEventArgs e) => RunGuarded(sender, BorrowBatchFromCart);

    private void BtnReturnBatch_OnClick(object sender, RoutedEventArgs e) => RunGuarded(sender, ReturnBatchFromCart);

    private void BtnBorrowExecute_OnClick(object sender, RoutedEventArgs e) => RunGuarded(sender, BorrowBatchFromCart);

    private void BtnReturnExecute_OnClick(object sender, RoutedEventArgs e) => RunGuarded(sender, ReturnBatchFromCart);

    private static void RunGuarded(object sender, Action action)
    {
        var button = sender as System.Windows.Controls.Button;
        if (button is not null)
        {
            if (!button.IsEnabled)
            {
                return;
            }

            button.IsEnabled = false;
        }

        try
        {
            action();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Desk batch action failed");
            MessageBox.Show($"הפעולה נכשלה: {ex.Message}", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (button is not null)
            {
                button.IsEnabled = true;
            }
        }
    }

    private void BtnRemoveBorrowCartItem_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: BorrowCartItem item })
        {
            return;
        }

        RemoveBorrowCartItem(item);
    }

    private void BtnRemoveReturnCartItem_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: ReturnCartItem item })
        {
            return;
        }

        RemoveReturnCartItem(item);
    }

    private void RemoveBorrowCartItem(BorrowCartItem item)
    {
        if (!_borrowCart.Remove(item))
        {
            return;
        }

        _borrowCartLookup.Remove(item.Barcode);
        SetStatus("הפריט הוסר מעגלת ההשאלה.", StatusKind.Info);
    }

    private void RemoveReturnCartItem(ReturnCartItem item)
    {
        if (!_returnCart.Remove(item))
        {
            return;
        }

        _returnCartLookup.Remove(item.Barcode);
        SetStatus("הפריט הוסר מעגלת ההחזרה.", StatusKind.Info);
    }

    private void ChkAllowOneTimeBorrow_OnChanged(object sender, RoutedEventArgs e)
    {
        if (_currentStudent is null)
        {
            chkAllowOneTimeBorrow.IsChecked = false;
            return;
        }

        UpdateBorrowAvailability();
    }

    private void AddBorrowToCart()
    {
        if (_currentStudent is null)
        {
            SetStatus("יש לבחור חניך תחילה", StatusKind.Error);
            SystemSounds.Hand.Play();
            txtStudentSearch.Focus();
            txtStudentSearch.SelectAll();
            return;
        }

        if (!_currentStudentPaidForYear && chkAllowOneTimeBorrow.IsChecked != true)
        {
            SetStatus("החניך לא שילם. לא ניתן להשאיל ללא אישור חריג.", StatusKind.Error);
            SystemSounds.Hand.Play();
            txtBorrowBarcode.Clear();
            txtBorrowBarcode.Focus();
            return;
        }

        var barcode = (txtBorrowBarcode.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(barcode))
        {
            SetStatus("יש להזין ברקוד עותק.", StatusKind.Error);
            txtBorrowBarcode.Clear();
            txtBorrowBarcode.Focus();
            return;
        }

        try
        {
            var ok = _databaseService.TryGetBorrowCartItem(barcode, out var title, out _, out var message, _currentStudent?.StuId);            if (!ok)
            {
                SetStatus(message, StatusKind.Error);
                SystemSounds.Hand.Play();
                txtBorrowBarcode.Clear();
                txtBorrowBarcode.Focus();
                return;
            }

            if (_borrowCartLookup.Contains(barcode))
            {
                SetStatus("העותק כבר קיים בעגלת ההשאלה.", StatusKind.Info);
                txtBorrowBarcode.Clear();
                txtBorrowBarcode.Focus();
                return;
            }

            _borrowCart.Add(new BorrowCartItem { IsSelected = true, Barcode = barcode, Title = title });
            _borrowCartLookup.Add(barcode);
            SetStatus("העותק נוסף לעגלת ההשאלה.", StatusKind.Success);
            SystemSounds.Asterisk.Play();
            txtBorrowBarcode.Clear();
            txtBorrowBarcode.Focus();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Add borrow cart item failed");
            SetStatus(ex.Message, StatusKind.Error);
            SystemSounds.Hand.Play();
            txtBorrowBarcode.Clear();
            txtBorrowBarcode.Focus();
        }
    }

    private void BorrowBatchFromCart()
    {
        if (_currentStudent is null)
        {
            SetStatus("יש לבחור חניך לפני השאלה מרוכזת.", StatusKind.Error);
            txtStudentSearch.Focus();
            txtStudentSearch.SelectAll();
            return;
        }

        if (!_currentStudentPaidForYear && chkAllowOneTimeBorrow.IsChecked != true)
        {
            SetStatus("החניך לא שילם. לא ניתן להשאיל ללא אישור חריג.", StatusKind.Error);
            SystemSounds.Hand.Play();
            txtBorrowBarcode.Focus();
            txtBorrowBarcode.SelectAll();
            return;
        }

        var selectedBarcodes = _borrowCart.Where(x => x.IsSelected).Select(x => x.Barcode).ToList();
        if (selectedBarcodes.Count == 0)
        {
            SetStatus("לא נבחרו עותקים לעגלת ההשאלה.", StatusKind.Error);
            return;
        }

        var success = _databaseService.TryBorrowCopiesBatch(_currentStudent.StuId, selectedBarcodes, out var message);
        SetStatus(success ? $"{message} עבור החניך {_currentStudent.FullName}" : message, success ? StatusKind.Success : StatusKind.Error);
        if (!success)
        {
            SystemSounds.Hand.Play();
            txtBorrowBarcode.Focus();
            txtBorrowBarcode.SelectAll();
            return;
        }

        foreach (var item in _borrowCart.Where(x => x.IsSelected).ToList())
        {
            _borrowCart.Remove(item);
            _borrowCartLookup.Remove(item.Barcode);
        }

        chkAllowOneTimeBorrow.IsChecked = false;
        LoadActiveLoans(_currentStudent.StuId);
        txtBorrowBarcode.Clear();
        txtBorrowBarcode.Focus();
        SystemSounds.Asterisk.Play();
    }

    private void AddReturnToCart(string? barcode, bool fromQuickMode = false)
    {
        var normalizedBarcode = (barcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedBarcode))
        {
            SetStatus("יש להזין ברקוד עותק.", StatusKind.Error);
            SystemSounds.Hand.Play();
            if (fromQuickMode)
            {
                txtQuickReturnBarcode.Focus();
                txtQuickReturnBarcode.SelectAll();
            }
            else
            {
                txtManualReturnBarcode.Focus();
                txtManualReturnBarcode.SelectAll();
            }

            return;
        }

        try
        {
            var ok = _databaseService.TryGetReturnCartItem(normalizedBarcode, out var title, out var borrowerName, out var message);
            if (!ok)
            {
                SetStatus(message, StatusKind.Error);
                SystemSounds.Hand.Play();
                if (fromQuickMode)
                {
                    txtQuickReturnBarcode.Focus();
                    txtQuickReturnBarcode.SelectAll();
                }
                else
                {
                    txtManualReturnBarcode.Focus();
                    txtManualReturnBarcode.SelectAll();
                }

                return;
            }

            if (_returnCartLookup.Contains(normalizedBarcode))
            {
                SetStatus("העותק כבר קיים בעגלת ההחזרה.", StatusKind.Info);
                if (fromQuickMode)
                {
                    txtQuickReturnBarcode.Clear();
                    txtQuickReturnBarcode.Focus();
                }
                else
                {
                    txtManualReturnBarcode.Clear();
                    txtManualReturnBarcode.Focus();
                }

                return;
            }

            _returnCart.Add(new ReturnCartItem
            {
                IsSelected = true,
                Barcode = normalizedBarcode,
                Title = title,
                BorrowerName = borrowerName
            });
            _returnCartLookup.Add(normalizedBarcode);
            SetStatus("העותק נוסף לעגלת ההחזרה.", StatusKind.Success);
            SystemSounds.Asterisk.Play();

            if (fromQuickMode)
            {
                txtQuickReturnBarcode.Clear();
                txtQuickReturnBarcode.Focus();
            }
            else
            {
                txtManualReturnBarcode.Clear();
                txtManualReturnBarcode.Focus();
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Add return cart item failed");
            SetStatus(ex.Message, StatusKind.Error);
            SystemSounds.Hand.Play();
        }
    }

    private bool BlockScrappedOrLostReturn(string barcode)
    {
        if (!_databaseService.IsCopyMarkedScrappedOrLost(barcode))
        {
            return false;
        }

        SystemSounds.Hand.Play();
        SetStatus("עותק זה מוגדר כ'גרוט' או 'אבוד'. יש לשחזר אותו במסך כותרים ועותקים.", StatusKind.Error);
        return true;
    }

    private void ReturnBatchFromCart()
    {
        var selectedBarcodes = _returnCart.Where(x => x.IsSelected).Select(x => x.Barcode).ToList();
        if (selectedBarcodes.Count == 0)
        {
            SetStatus("לא נבחרו עותקים לעגלת ההחזרה.", StatusKind.Error);
            return;
        }

        foreach (var selectedBarcode in selectedBarcodes)
        {
            if (BlockScrappedOrLostReturn(selectedBarcode))
            {
                return;
            }
        }

        var success = _databaseService.TryReturnCopiesBatch(selectedBarcodes, out var message);
        SetStatus(message, success ? StatusKind.Success : StatusKind.Error);
        if (!success)
        {
            SystemSounds.Hand.Play();
            txtManualReturnBarcode.Focus();
            txtManualReturnBarcode.SelectAll();
            return;
        }

        foreach (var item in _returnCart.Where(x => x.IsSelected).ToList())
        {
            _returnCart.Remove(item);
            _returnCartLookup.Remove(item.Barcode);
        }

        if (_currentStudent is not null)
        {
            LoadActiveLoans(_currentStudent.StuId);
        }

        txtManualReturnBarcode.Clear();
        txtManualReturnBarcode.Focus();
        SystemSounds.Asterisk.Play();
    }

    private void BtnPrintActiveLoans_OnClick(object sender, RoutedEventArgs e) => PrintActiveLoans();

    private void DgActiveLoans_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        try
        {
            if (dgActiveLoans.SelectedItem is not DataRowView selectedRow)
            {
                return;
            }

            var stuId = selectedRow.Row.Table.Columns.Contains("stuID")
                ? selectedRow["stuID"]?.ToString()?.Trim() ?? string.Empty
                : string.Empty;

            if (string.IsNullOrWhiteSpace(stuId))
            {
                stuId = (_currentStudent?.StuId ?? txtStuId.Text ?? string.Empty).Trim();
            }

            if (string.IsNullOrWhiteSpace(stuId))
            {
                return;
            }

            if (Application.Current.MainWindow is MainWindow mainWindow)
            {
                mainWindow.NavigateToStudentsWithStudent(stuId);
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Desk active loans double-click navigation failed");
            SetStatus("שגיאה במעבר לכרטיס החניך.", StatusKind.Error);
        }
    }

    private void SearchStudent()
    {
        SetQuickReturnMode(false);
        var ambiguousMatches = _studentSuggestions.Count > 1;
        CloseStudentSuggestions();
        try
        {
            var student = _databaseService.FindStudent(txtStudentSearch.Text);
            if (student is null)
            {
                _currentStudent = null;
                ClearStudentDisplay();
                SetStatus(ambiguousMatches
                    ? "נמצאו מספר חניכים. יש לבחור חניך מהרשימה."
                    : "לא נמצא חניך לפי החיפוש.", StatusKind.Error);
                SystemSounds.Hand.Play();
                txtStudentSearch.Focus();
                txtStudentSearch.SelectAll();
                return;
            }

            LoadStudentCard(student);
            SetStatus($"נטען חניך: {student.FullName}", StatusKind.Success);
            SystemSounds.Asterisk.Play();
            txtBorrowBarcode.Focus();
            txtBorrowBarcode.SelectAll();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Search student failed");
            SetStatus("שגיאה בחיפוש חניך.", StatusKind.Error);
            SystemSounds.Hand.Play();
        }
    }

    private static string ClassDisplay(string? className, string? classCode)
    {
        return string.IsNullOrWhiteSpace(className)
            ? ClassFormatter.ToDisplay(classCode)
            : ClassFormatter.ToDisplayName(className, classCode);
    }

    private void LoadStudentCard(Student student)
    {
        if (_currentStudent is not null && !string.Equals(_currentStudent.StuId, student.StuId, StringComparison.Ordinal))
        {
            _borrowCart.Clear();
            _borrowCartLookup.Clear();
        }

        _suppressStudentSearchEvents = true;
        try
        {
            _currentStudent = student;
            txtStuId.Text = student.StuId;
            txtFirstName.Text = student.FirstName;
            txtLastName.Text = student.LastName;
            txtClass.Text = ClassDisplay(student.ClassName, student.ClassCode);
            txtStudentSearch.Text = student.StuId;
        }
        finally
        {
            _suppressStudentSearchEvents = false;
        }

        chkAllowOneTimeBorrow.IsChecked = false;
        UpdateStudentPaymentStatus(student.StuId);
        LoadActiveLoans(student.StuId);
        UpdateBorrowAvailability();
    }

    private void LoadActiveLoans(string stuId)
    {
        DataTable loans = _databaseService.GetActiveLoansByStudent(stuId);

        bool hasOverdue = loans.AsEnumerable().Any(row => row["IsOverdue"]?.ToString() == "1");
        SetOverdueWarningVisibility(hasOverdue);
        if (hasOverdue) { SystemSounds.Exclamation.Play(); }

        dgActiveLoans.ItemsSource = loans.DefaultView;
        UpdateBorrowAvailability();
    }

    private void UpdateStudentPaymentStatus(string stuId)
    {
        var isExempt = _databaseService.IsStudentExemptForCurrentSchoolYear(stuId);
        _currentStudentPaidForYear = isExempt || _databaseService.IsPaymentCompletedForCurrentSchoolYear(stuId);

        if (isExempt)
        {
            paymentBadge.Visibility = Visibility.Collapsed;
            paymentBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB"));
            txtPaymentStatus.Text = "פטור";
            return;
        }

        if (_currentStudentPaidForYear)
        {
            paymentBadge.Visibility = Visibility.Collapsed;
            paymentBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#16A34A"));
            txtPaymentStatus.Text = "שולם";
            return;
        }

        paymentBadge.Visibility = Visibility.Visible;
        paymentBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DC2626"));
        txtPaymentStatus.Text = "לא שולם";
    }

    private void UpdateBorrowAvailability()
    {
        var hasStudent = _currentStudent is not null;
        var allowOverride = chkAllowOneTimeBorrow.IsChecked == true;
        var canBorrow = hasStudent && (_currentStudentPaidForYear || allowOverride);

        btnBorrow.IsEnabled = canBorrow || !hasStudent;
        txtBorrowBarcode.IsEnabled = true;
    }

    private void BorrowCopy()
    {
        if (_currentStudent is null)
        {
            SetStatus("יש לבחור חניך תחילה", StatusKind.Error);
            SystemSounds.Hand.Play();
            txtStudentSearch.Focus();
            txtStudentSearch.SelectAll();
            return;
        }

        try
        {
            var success = _databaseService.TryBorrowCopy(_currentStudent.StuId, txtBorrowBarcode.Text, out var message);
            SetStatus(success ? $"{message} עבור החניך {_currentStudent.FullName}" : message, success ? StatusKind.Success : StatusKind.Error);

            if (!success)
            {
                SystemSounds.Hand.Play();
                txtBorrowBarcode.Focus();
                txtBorrowBarcode.SelectAll();
                return;
            }

            SystemSounds.Asterisk.Play();
            txtBorrowBarcode.Clear();
            chkAllowOneTimeBorrow.IsChecked = false;
            txtBorrowBarcode.Focus();
            LoadActiveLoans(_currentStudent.StuId);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Borrow copy failed");
            SetStatus(ex.Message, StatusKind.Error);
            SystemSounds.Hand.Play();
            txtBorrowBarcode.Focus();
            txtBorrowBarcode.SelectAll();
        }
    }

    private void ReturnCopy(string? barcode, bool forceContinuousMode = false)
    {
        var normalizedBarcode = (barcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedBarcode))
        {
            SetStatus("יש להזין ברקוד עותק.", StatusKind.Error);
            SystemSounds.Hand.Play();
            if (forceContinuousMode)
            {
                txtQuickReturnBarcode.Focus();
                txtQuickReturnBarcode.SelectAll();
            }
            else
            {
                txtManualReturnBarcode.Focus();
                txtManualReturnBarcode.SelectAll();
            }
            return;
        }

        try
        {
            var preReturn = forceContinuousMode ? LookupBorrowedCopyDetails(normalizedBarcode) : (string.Empty, string.Empty, string.Empty);
            var outcome = _databaseService.TryReturnCopyWithBorrowerInfo(normalizedBarcode, out var message, out var borrowerInfo);
            var success = outcome == DatabaseService.ReturnOutcome.Ok;
            SetStatus(message, success ? StatusKind.Success : StatusKind.Error);

            if (!success)
            {
                SystemSounds.Hand.Play();
                if (forceContinuousMode)
                {
                    if (outcome == DatabaseService.ReturnOutcome.NotBorrowed)
                    {
                        SetStatus("עותק זה אינו רשום כמושאל.", StatusKind.Error);
                    }

                    txtQuickReturnBarcode.Focus();
                    txtQuickReturnBarcode.SelectAll();
                    return;
                }

                txtManualReturnBarcode.Focus();
                txtManualReturnBarcode.SelectAll();
                return;
            }

            SystemSounds.Asterisk.Play();
            if (forceContinuousMode)
            {
                txtQuickReturnBarcode.Clear();
                txtQuickReturnBarcode.Focus();
                AddQuickReturnLog(
                    normalizedBarcode,
                    preReturn.Item1,
                    borrowerInfo?.FullName ?? preReturn.Item2,
                    preReturn.Item3);
            }
            else
            {
                txtManualReturnBarcode.Clear();
                txtManualReturnBarcode.Focus();
            }

            if (_currentStudent is not null)
            {
                LoadActiveLoans(_currentStudent.StuId);
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Return copy failed");
            SetStatus(ex.Message, StatusKind.Error);
            SystemSounds.Hand.Play();
            if (forceContinuousMode)
            {
                txtQuickReturnBarcode.Focus();
                txtQuickReturnBarcode.SelectAll();
            }
            else
            {
                txtManualReturnBarcode.Focus();
                txtManualReturnBarcode.SelectAll();
            }
        }
    }

    private void ProcessBarcodeScan()
    {
        if (_quickReturnMode)
        {
            ReturnCopy(txtQuickReturnBarcode.Text, forceContinuousMode: true);
            return;
        }

        if (_currentStudent is not null)
        {
            BorrowCopy();
        }
        else
        {
            SetStatus("יש לבחור חניך תחילה", StatusKind.Error);
            SystemSounds.Hand.Play();
            txtStudentSearch.Focus();
            txtStudentSearch.SelectAll();
        }
    }

    private void PrintActiveLoans()
    {
        if (_currentStudent is null)
        {
            SetStatus("בחר חניך לפני הדפסה.", StatusKind.Error);
            return;
        }

        try
        {
            var loans = _databaseService.GetActiveLoansByStudent(_currentStudent.StuId);
            if (loans.Rows.Count == 0)
            {
                SetStatus("אין השאלות פעילות להדפסה.", StatusKind.Error);
                return;
            }

            var document = BuildLoanReportDocument(_currentStudent, loans);
            var printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                printDialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, "Active Student Loans");
                SetStatus("דוח השאלות נשלח להדפסה.", StatusKind.Success);
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Print active loans failed");
            SetStatus("שגיאה בהדפסת פירוט השאלות.", StatusKind.Error);
        }
    }

    private static FlowDocument BuildLoanReportDocument(Student student, DataTable loans)
    {
        var document = new FlowDocument
        {
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 12,
            PagePadding = new Thickness(40),
            FlowDirection = FlowDirection.RightToLeft
        };

        document.Blocks.Add(new Paragraph(new Run("פירוט השאלות פעילות"))
        {
            FontSize = 18,
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(0, 0, 0, 8)
        });

        document.Blocks.Add(new Paragraph(new Run($"חניך: {student.FullName} | ת.ז: {student.StuId} | כיתה: {ClassDisplay(student.ClassName, student.ClassCode)}"))
        {
            Margin = new Thickness(0, 0, 0, 12)
        });

        var table = new Table();
        table.Columns.Add(new TableColumn { Width = new GridLength(120) });
        table.Columns.Add(new TableColumn { Width = new GridLength(260) });
        table.Columns.Add(new TableColumn { Width = new GridLength(140) });

        var rowGroup = new TableRowGroup();
        table.RowGroups.Add(rowGroup);

        var header = new TableRow();
        header.Cells.Add(new TableCell(new Paragraph(new Run("ברקוד"))) { FontWeight = FontWeights.Bold });
        header.Cells.Add(new TableCell(new Paragraph(new Run("כותר"))) { FontWeight = FontWeights.Bold });
        header.Cells.Add(new TableCell(new Paragraph(new Run("תאריך החזרה"))) { FontWeight = FontWeights.Bold });
        rowGroup.Rows.Add(header);

        foreach (DataRow row in loans.Rows)
        {
            var barcode = row.Table.Columns.Contains("Barcode") ? row["Barcode"]?.ToString() : string.Empty;
            var title = row.Table.Columns.Contains("Title") ? row["Title"]?.ToString() : string.Empty;
            var dateToReturn = row.Table.Columns.Contains("DateToReturn") ? row["DateToReturn"]?.ToString() : string.Empty;

            var dataRow = new TableRow();
            dataRow.Cells.Add(new TableCell(new Paragraph(new Run(barcode ?? string.Empty))));
            dataRow.Cells.Add(new TableCell(new Paragraph(new Run(title ?? string.Empty))));
            dataRow.Cells.Add(new TableCell(new Paragraph(new Run(dateToReturn ?? string.Empty))));
            rowGroup.Rows.Add(dataRow);
        }

        document.Blocks.Add(table);
        SchoolBranding.AddLetterhead(document);
        return document;
    }

    private void ClearStudentDisplay()
    {
        _currentStudent = null;
        _borrowCart.Clear();
        _borrowCartLookup.Clear();
        txtStuId.Clear();
        txtFirstName.Clear();
        txtLastName.Clear();
        txtClass.Clear();
        txtStudentSearch.Clear();
        CloseStudentSuggestions();
        dgActiveLoans.ItemsSource = null;
        chkAllowOneTimeBorrow.IsChecked = false;
        _currentStudentPaidForYear = false;
        SetOverdueWarningVisibility(false);
        ResetPaymentDisplay();
        UpdateBorrowAvailability();
    }

    private void ResetStudentSelection()
    {
        _currentStudent = null;
        _borrowCart.Clear();
        _borrowCartLookup.Clear();
        txtStuId.Clear();
        txtFirstName.Clear();
        txtLastName.Clear();
        txtClass.Clear();
        CloseStudentSuggestions();
        dgActiveLoans.ItemsSource = null;
        chkAllowOneTimeBorrow.IsChecked = false;
        _currentStudentPaidForYear = false;
        SetOverdueWarningVisibility(false);
        ResetPaymentDisplay();
        UpdateBorrowAvailability();
    }

    private void ResetDesk()
    {
        ClearStudentDisplay();
        SetQuickReturnMode(false);
        txtBorrowBarcode.Clear();
        txtManualReturnBarcode.Clear();
        txtQuickReturnBarcode.Clear();
        _returnedThisSession.Clear();
        SetStatus("הדלפק אופס.", StatusKind.Info);
        txtStudentSearch.Focus();
        txtStudentSearch.SelectAll();
    }

    private void ResetPaymentDisplay()
    {
        paymentBadge.Visibility = Visibility.Collapsed;
        paymentBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DC2626"));
        txtPaymentStatus.Text = "לא שולם";
    }

    private void SetOverdueWarningVisibility(bool isVisible)
    {
        if (FindName("borderOverdueWarning") is Border warningBorder)
        {
            warningBorder.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private void SetStatus(string message, StatusKind kind)
    {
        txtStatus.Text = message;
        statusBadge.Background = kind switch
        {
            StatusKind.Success => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#16A34A")),
            StatusKind.Error => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DC2626")),
            _ => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1D4ED8"))
        };
    }

    private enum StatusKind
    {
        Info,
        Success,
        Error
    }

    public class ReturnedBookItem
    {
        public string ReturnTime { get; set; } = "";
        public string Barcode { get; set; } = "";
        public string Title { get; set; } = "";
        public string StudentName { get; set; } = "";
        public string ClassName { get; set; } = "";
    }

    public class BorrowCartItem
    {
        public bool IsSelected { get; set; }
        public string Barcode { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
    }

    public class ReturnCartItem
    {
        public bool IsSelected { get; set; }
        public string Barcode { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string BorrowerName { get; set; } = string.Empty;
    }

    private sealed class StudentSuggestionItem
    {
        public string StuId { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string ClassCode { get; set; } = string.Empty;
        public string DisplayText { get; set; } = string.Empty;
    }
}
