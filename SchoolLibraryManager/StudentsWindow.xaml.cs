using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace SchoolLibraryManager;

public partial class StudentsWindow : Window
{
    private readonly DatabaseService _dbService;
    private DataTable _studentsTable = new();
    private DataTable _cardClassesTable = new();
    private string? _selectedStudentId;
    private bool _isStudentCardEditMode;
    private bool _suppressExemptionEvents;
    private bool _isInitializing;
    private bool _showInactiveStudents;
    private readonly DispatcherTimer _searchDebounceTimer;
    private const string AllClassesFilter = "כל הכיתות";
    private const string OtherClassDisplay = "--אחר--";
    private const string OtherClassCode = "0000";

    private static bool IsOtherClassValue(string? value)
    {
        var normalized = (value ?? string.Empty).Trim();
        return string.IsNullOrWhiteSpace(normalized)
            || string.Equals(normalized, "0", StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, OtherClassCode, StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, OtherClassDisplay, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeClassIdForStorage(string? classId)
    {
        var normalized = (classId ?? string.Empty).Trim();
        return IsOtherClassValue(normalized) ? OtherClassCode : normalized;
    }

    private static string ToClassDisplay(string? className, string? classId)
    {
        var normalizedName = (className ?? string.Empty).Trim();
        var normalizedId = (classId ?? string.Empty).Trim();

        if (ClassFormatter.ToDisplayName(normalizedName, normalizedId) == ClassFormatter.TeacherDisplay)
        {
            return ClassFormatter.TeacherDisplay;
        }

        if (!string.IsNullOrWhiteSpace(normalizedName) && !IsOtherClassValue(normalizedName))
        {
            return normalizedName;
        }

        return IsOtherClassValue(normalizedId) ? OtherClassDisplay : normalizedId;
    }

    private void NormalizeClassDisplayColumns()
    {
        if (_studentsTable.Columns.Count == 0)
        {
            return;
        }

        var hasClassId = _studentsTable.Columns.Contains("stuClassID");
        var hasClassName = _studentsTable.Columns.Contains("ClassName");
        if (!hasClassId && !hasClassName)
        {
            return;
        }

        foreach (DataRow row in _studentsTable.Rows)
        {
            var classId = hasClassId ? NormalizeClassIdForStorage(row["stuClassID"]?.ToString()) : string.Empty;
            if (hasClassId)
            {
                row["stuClassID"] = classId;
            }

            if (hasClassName)
            {
                row["ClassName"] = ToClassDisplay(row["ClassName"]?.ToString(), classId);
            }
        }
    }

    public StudentsWindow(DatabaseService dbService)
    {
        InitializeComponent();
        _dbService = dbService;
        DataContext ??= this;
        _searchDebounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(250)
        };
        _searchDebounceTimer.Tick += SearchDebounceTimer_Tick;

        if (Content is FrameworkElement rootElement)
        {
            rootElement.Loaded += StudentsWindow_OnLoaded;
        }
        else
        {
            Dispatcher.BeginInvoke(new Action(() => StudentsWindow_OnLoaded(this, new RoutedEventArgs())), DispatcherPriority.Loaded);
        }
    }

    public void FocusStudentById(string stuId)
    {
        var normalizedStuId = (stuId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedStuId))
        {
            TxtStudentSearch.Focus();
            return;
        }

        void ApplyFocus()
        {
            _searchDebounceTimer.Stop();
            TxtStudentSearch.Text = normalizedStuId;
            CmbClassFilter.SelectedIndex = 0;
            PerformSearch(normalizedStuId);
            TxtStudentSearch.Focus();
            TxtStudentSearch.SelectAll();
        }

        if (_isInitializing || _studentsTable.Columns.Count == 0)
        {
            Dispatcher.BeginInvoke((Action)ApplyFocus, DispatcherPriority.Background);
            return;
        }

        ApplyFocus();
    }

    private void StudentsWindow_OnLoaded(object sender, RoutedEventArgs e)
    {
        _isInitializing = true;
        try
        {
            LoadStudentsData();
            LoadClassFilter();
            _showInactiveStudents = ChkShowInactiveStudents.IsChecked == true;
            TxtStudentSearch.Focus();
            SetStatus("נטען חלון ניהול חניכים בהצלחה.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Students window load failed");
            SetStatus("טעינת חלון ניהול חניכים נכשלה.");
        }
        finally
        {
            _isInitializing = false;
        }

        PerformSearch();
    }

    private void LoadClassFilter()
    {
        var classFilterTable = new DataTable();
        classFilterTable.Columns.Add("ClassID", typeof(string));
        classFilterTable.Columns.Add("ClassName", typeof(string));
        var classIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var allRow = classFilterTable.NewRow();
        allRow["ClassID"] = string.Empty;
        allRow["ClassName"] = AllClassesFilter;
        classFilterTable.Rows.Add(allRow);

        var distinctClasses = _studentsTable.AsEnumerable()
            .Select(row =>
            {
                var classId = NormalizeClassIdForStorage(row["stuClassID"]?.ToString());
                var className = ToClassDisplay(row["ClassName"]?.ToString(), classId);

                return new { ClassId = classId, ClassName = className };
            })
            .Where(item => !string.IsNullOrWhiteSpace(item.ClassId))
            .GroupBy(item => item.ClassId, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(item => item.ClassName, StringComparer.CurrentCultureIgnoreCase);

        foreach (var item in distinctClasses)
        {
            var row = classFilterTable.NewRow();
            row["ClassID"] = item.ClassId;
            row["ClassName"] = item.ClassName;
            classFilterTable.Rows.Add(row);
            classIds.Add(item.ClassId);
        }

        var classesLookup = _dbService.GetClasses();
        foreach (DataRow classRow in classesLookup.Rows)
        {
            var classId = NormalizeClassIdForStorage(classRow["ClassID"]?.ToString());
            if (classId.Length is > 0 and < 4 && classId.All(char.IsDigit))
            {
                classId = classId.PadLeft(4, '0');
            }

            if (string.IsNullOrWhiteSpace(classId) || classIds.Contains(classId))
            {
                continue;
            }

            var className = ToClassDisplay(classRow["ClassName"]?.ToString(), classId);
            var row = classFilterTable.NewRow();
            row["ClassID"] = classId;
            row["ClassName"] = className;
            classFilterTable.Rows.Add(row);
            classIds.Add(classId);
        }

        if (!classIds.Contains(ClassFormatter.TeacherCode))
        {
            var teacherRow = classFilterTable.NewRow();
            teacherRow["ClassID"] = ClassFormatter.TeacherCode;
            teacherRow["ClassName"] = ClassFormatter.TeacherDisplay;
            classFilterTable.Rows.Add(teacherRow);
            classIds.Add(ClassFormatter.TeacherCode);
        }

        CmbClassFilter.DisplayMemberPath = "ClassName";
        CmbClassFilter.SelectedValuePath = "ClassID";
        CmbClassFilter.ItemsSource = classFilterTable.DefaultView;
        CmbClassFilter.SelectedIndex = 0;

        _cardClassesTable = classFilterTable.Copy();
        var allClassRows = _cardClassesTable.Select($"ClassName = '{AllClassesFilter.Replace("'", "''", StringComparison.Ordinal)}'");
        foreach (var row in allClassRows)
        {
            _cardClassesTable.Rows.Remove(row);
        }

        if (CmbCardClass != null)
        {
            CmbCardClass.DisplayMemberPath = "ClassName";
            CmbCardClass.SelectedValuePath = "ClassID";
            CmbCardClass.ItemsSource = _cardClassesTable.DefaultView;
        }
    }

    private void LoadStudentsData()
    {
        _studentsTable = _dbService.SearchStudents(string.Empty, null);
        NormalizeClassDisplayColumns();
        _studentsTable.DefaultView.RowFilter = string.Empty;
        GridStudents.ItemsSource = _studentsTable.DefaultView;
        var columns = string.Join(", ", _studentsTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
        Debug.WriteLine($"StudentsView loaded columns: {columns}");
    }

    private static string EscapeForLike(string value)
    {
        return value.Replace("[", "[[]", StringComparison.Ordinal).Replace("*", "[*]", StringComparison.Ordinal).Replace("%", "[%]", StringComparison.Ordinal);
    }

    private static string EscapeForRowFilter(string value)
    {
        return value.Replace("'", "''", StringComparison.Ordinal);
    }

    private void PerformSearch(string? preferredStudentId = null)
    {
        if (_studentsTable.Columns.Count == 0)
        {
            GridStudents.ItemsSource = null;
            ClearStudentCard();
            return;
        }

        try
        {
            var view = _studentsTable.DefaultView;
            var query = EscapeForLike(EscapeForRowFilter((TxtStudentSearch.Text ?? string.Empty).Trim()));
            var classFilter = CmbClassFilter.SelectedValue?.ToString()?.Trim();
            var isAllClasses = string.IsNullOrWhiteSpace(classFilter) || string.Equals(classFilter, AllClassesFilter, StringComparison.OrdinalIgnoreCase) || string.Equals(classFilter, "0", StringComparison.OrdinalIgnoreCase);

            var filters = new List<string>();
            if (_showInactiveStudents)
            {
                filters.Add("([stuClassID] IS NOT NULL AND Convert([stuClassID], 'System.String') = '9999')");
            }
            else
            {
                filters.Add("([stuClassID] IS NULL OR Convert([stuClassID], 'System.String') <> '9999')");
            }

            if (!isAllClasses)
            {
                filters.Add($"Convert([stuClassID], 'System.String') = '{EscapeForRowFilter(classFilter!)}'");
            }

            if (!string.IsNullOrWhiteSpace(query))
            {
                filters.Add($"([stuFname] LIKE '%{query}%' OR [stuLname] LIKE '%{query}%' OR Convert([stuID], 'System.String') LIKE '%{query}%')");
            }

            view.RowFilter = filters.Count > 0 ? string.Join(" AND ", filters) : string.Empty;
            GridStudents.ItemsSource = view;

            if (view.Count == 0)
            {
                _selectedStudentId = null;
                ClearStudentCard();
                return;
            }

            DataRowView? targetRow = null;
            var idToSelect = preferredStudentId ?? _selectedStudentId;
            if (!string.IsNullOrWhiteSpace(idToSelect))
            {
                targetRow = view
                    .Cast<DataRowView>()
                    .FirstOrDefault(r => string.Equals(r["stuID"]?.ToString(), idToSelect, StringComparison.OrdinalIgnoreCase));
            }

            GridStudents.SelectedItem = targetRow ?? view[0];
            GridStudents.ScrollIntoView(GridStudents.SelectedItem);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"PerformSearch failed: {ex.Message}");
            ErrorLogger.Log(ex, "Student filter failed");
            SetStatus("אירעה שגיאה בסינון החניכים.");
        }
    }

    private void SearchDebounceTimer_Tick(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        try
        {
            PerformSearch();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"SearchDebounceTimer_Tick failed: {ex.Message}");
            ErrorLogger.Log(ex, "Debounced student search failed");
            SetStatus("אירעה שגיאה בסינון החניכים.");
        }
    }

    private void TxtStudentSearch_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializing)
        {
            return;
        }

        _searchDebounceTimer.Stop();
        _searchDebounceTimer.Start();
    }

    private void CmbClassFilter_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializing)
        {
            return;
        }

        _searchDebounceTimer.Stop();
        _searchDebounceTimer.Start();
    }

    private void ChkShowInactiveStudents_OnCheckedChanged(object sender, RoutedEventArgs e)
    {
        if (_isInitializing)
        {
            return;
        }

        _showInactiveStudents = ChkShowInactiveStudents.IsChecked == true;
        _searchDebounceTimer.Stop();

        if (CmbClassFilter.SelectedIndex != 0)
        {
            CmbClassFilter.SelectedIndex = 0;
        }

        PerformSearch(_selectedStudentId);
    }

    private void GridStudents_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        try
        {
            if (_isInitializing)
            {
                return;
            }

            if (GridStudents.SelectedItem is not DataRowView row)
            {
                _selectedStudentId = null;
                ClearStudentCard();
                return;
            }

            if (row.Row is null || row.DataView is null)
            {
                _selectedStudentId = null;
                ClearStudentCard();
                return;
            }

            LoadStudentCard(row);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Grid students selection change failed");
            SetStatus($"אירעה שגיאה בעת בחירת חניך: {ex.Message}");
        }
    }

    private void Row_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        try
        {
            if (_isInitializing)
            {
                return;
            }

            var selectedRow = (sender as DataGridRow)?.Item as DataRowView
                ?? (e.OriginalSource as FrameworkElement)?.DataContext as DataRowView
                ?? GridStudents.SelectedItem as DataRowView;

            if (selectedRow is null)
            {
                return;
            }

            GridStudents.SelectedItem = selectedRow;

            var stuId = GetRowText(selectedRow, "stuID").Trim();
            if (string.IsNullOrWhiteSpace(stuId))
            {
                return;
            }

            var mainWindow = Application.Current.MainWindow as MainWindow;
            mainWindow?.NavigateToDeskWithStudent(stuId);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Student row double-click failed");
            SetStatus($"אירעה שגיאה במעבר לדלפק: {ex.Message}");
        }
    }

    private void GridStudents_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        try
        {
            if (_isInitializing)
            {
                return;
            }

            var source = e.OriginalSource as DependencyObject;
            var row = FindAncestor<DataGridRow>(source);
            if (row?.Item is not DataRowView selectedRow)
            {
                return;
            }

            var stuId = GetRowText(selectedRow, "stuID").Trim();
            if (string.IsNullOrWhiteSpace(stuId))
            {
                return;
            }

            GridStudents.SelectedItem = selectedRow;
            e.Handled = true;

            var mainWindow = Application.Current.MainWindow as MainWindow;
            mainWindow?.NavigateToDeskWithStudent(stuId);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Student grid preview double-click failed");
            SetStatus($"אירעה שגיאה במעבר לדלפק: {ex.Message}");
        }
    }

    private void LoadStudentCard(DataRowView row)
    {
        try
        {
            SetStudentCardEditMode(false);
            _selectedStudentId = GetRowText(row, "stuID");

            var firstName = GetRowText(row, "stuFname");
            var lastName = GetRowText(row, "stuLname");
            var className = GetRowText(row, "ClassName");
            var classId = GetRowText(row, "stuClassID");
            var phone = GetRowText(row, "stuPhone");
            var city = GetRowText(row, "stuCity");
            var remarks = GetRowText(row, "stuRemarks");

            if (TxtCardStuId != null) TxtCardStuId.Text = _selectedStudentId;
            if (TxtCardFirstName != null) TxtCardFirstName.Text = firstName;
            if (TxtCardLastName != null) TxtCardLastName.Text = lastName;
            var normalizedClassId = NormalizeClassIdForStorage(classId);
            if (CmbCardClass != null)
            {
                CmbCardClass.SelectedValue = normalizedClassId;
                if (CmbCardClass.SelectedIndex < 0)
                {
                    CmbCardClass.Text = string.IsNullOrWhiteSpace(className) ? normalizedClassId : className;
                }
            }

            if (TxtCardPhone != null) TxtCardPhone.Text = phone;
            if (TxtCardCity != null) TxtCardCity.Text = city;
            if (TxtCardRemarks != null) TxtCardRemarks.Text = remarks;

            if (string.IsNullOrWhiteSpace(_selectedStudentId))
            {
                _suppressExemptionEvents = true;
                if (ChkExemptCurrentYear != null)
                {
                    ChkExemptCurrentYear.IsChecked = false;
                    ChkExemptCurrentYear.IsEnabled = false;
                }

                _suppressExemptionEvents = false;
                if (GridActiveLoans != null)
                {
                    GridActiveLoans.ItemsSource = null;
                }

                return;
            }

            RefreshStudentTables(_selectedStudentId);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Load student card failed");
            SetStatus($"אירעה שגיאה בטעינת כרטיס החניך: {ex.Message}");
        }
    }

    private static string GetRowText(DataRowView? row, string columnName)
    {
        var columns = row?.Row?.Table?.Columns;
        if (columns is not null && columns.Contains(columnName))
        {
            return row?[columnName]?.ToString()?.Trim() ?? string.Empty;
        }

        return string.Empty;
    }

    private static T? FindAncestor<T>(DependencyObject? child) where T : DependencyObject
    {
        while (child is not null)
        {
            if (child is T target)
            {
                return target;
            }

            child = VisualTreeHelper.GetParent(child);
        }

        return null;
    }

    private void RefreshStudentTables(string stuId)
    {
        _suppressExemptionEvents = true;
        try
        {
            if (ChkExemptCurrentYear != null)
            {
                ChkExemptCurrentYear.IsChecked = _dbService.IsStudentExemptForCurrentSchoolYear(stuId);
                ChkExemptCurrentYear.IsEnabled = true;
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Failed to load exemption status");
            SetStatus($"שגיאה בטעינת סטטוס פטור: {ex.Message}");

            if (ChkExemptCurrentYear != null)
            {
                ChkExemptCurrentYear.IsChecked = false;
                ChkExemptCurrentYear.IsEnabled = false;
            }
        }
        finally
        {
            _suppressExemptionEvents = false;
        }

        try
        {
            UpdatePaymentStatusBadge(stuId);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Failed to update payment status badge");
            SetStatus($"שגיאה בטעינת סטטוס תשלום: {ex.Message}");
        }

        try
        {
            var activeLoans = _dbService.GetActiveLoansByStudent(stuId);
            if (GridActiveLoans != null)
            {
                GridActiveLoans.ItemsSource = activeLoans.DefaultView;
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Failed to load active loans");
            SetStatus($"שגיאה בטעינת השאלות פעילות: {ex.Message}");
            if (GridActiveLoans != null)
            {
                GridActiveLoans.ItemsSource = null;
            }
        }
    }

    private void BtnClearSearch_OnClick(object sender, RoutedEventArgs e)
    {
        TxtStudentSearch.Clear();
        CmbClassFilter.SelectedIndex = 0;
        PerformSearch(_selectedStudentId);
        TxtStudentSearch.Focus();
    }

    private void BtnDeleteSelectedStudent_OnClick(object sender, RoutedEventArgs e)
    {
        var stuId = (_selectedStudentId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(stuId))
        {
            SetStatus("יש לבחור חניך למחיקה.");
            return;
        }

        try
        {
            var studentName = $"{TxtCardFirstName.Text} {TxtCardLastName.Text}".Trim();
            if (!_dbService.CanDeleteStudent(stuId, out var validationMessage))
            {
                SetStatus(validationMessage);
                MessageBox.Show(validationMessage, "לא ניתן למחוק חניך", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                "האם לסמן את החניך הנבחר כלא-פעיל ולהסתירו מרשימות פעילות?",
                "אישור הסרת חניך מרשימות פעילות",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question,
                MessageBoxResult.No);

            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            var success = _dbService.TryDeleteStudent(stuId, out var message);
            SetStatus(message);
            if (!success)
            {
                MessageBox.Show(message, "מחיקת חניך", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _selectedStudentId = null;
            LoadStudentsData();
            LoadClassFilter();
            PerformSearch();
            var successMessage = $"החניך {(string.IsNullOrWhiteSpace(studentName) ? stuId : studentName)} הועבר למצב לא-פעיל והוסר מרשימת החניכים הפעילים בהצלחה";
            SetStatus(successMessage);
            MessageBox.Show(successMessage, "מחיקת חניך", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Delete student failed");
            var detailedMessage = $"מחיקת החניך נכשלה: {ex.Message}";
            SetStatus(detailedMessage);
            MessageBox.Show(detailedMessage, "מחיקת חניך", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnEditStudent_OnClick(object sender, RoutedEventArgs e)
    {
        var stuId = (_selectedStudentId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(stuId))
        {
            SetStatus("יש לבחור חניך לפני עריכה.");
            return;
        }

        if (!_isStudentCardEditMode)
        {
            SetStudentCardEditMode(true);
            SetStatus("מצב עריכה פעיל. ניתן לעדכן את פרטי החניך וללחוץ 'שמור שינויים'.");
            return;
        }

        var firstName = (TxtCardFirstName?.Text ?? string.Empty).Trim();
        var lastName = (TxtCardLastName?.Text ?? string.Empty).Trim();
        var classId = NormalizeClassIdForStorage(CmbCardClass?.SelectedValue?.ToString() ?? CmbCardClass?.Text);
        var phone = (TxtCardPhone?.Text ?? string.Empty).Trim();
        var city = (TxtCardCity?.Text ?? string.Empty).Trim();
        var remarks = (TxtCardRemarks?.Text ?? string.Empty).Trim();

        if (string.IsNullOrWhiteSpace(firstName) || string.IsNullOrWhiteSpace(lastName))
        {
            SetStatus("שם פרטי ושם משפחה הם שדות חובה.");
            return;
        }

        try
        {
            var success = _dbService.TryUpdateStudent(stuId, firstName, lastName, classId, phone, city, remarks, out var message);
            SetStatus(message);
            if (!success)
            {
                MessageBox.Show(message, "עדכון חניך", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            UpdateStudentRowCache(stuId, firstName, lastName, classId, phone, city, remarks);
            SetStudentCardEditMode(false);
            if (GridStudents.SelectedItem is DataRowView selectedRow)
            {
                LoadStudentCard(selectedRow);
            }

            RefreshStudentTables(stuId);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Edit student failed");
            SetStatus($"שמירת פרטי החניך נכשלה: {ex.Message}");
        }
    }

    private void BtnCancelEditStudent_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if (GridStudents.SelectedItem is DataRowView selectedRow)
            {
                LoadStudentCard(selectedRow);
                SetStatus("עריכת פרטי החניך בוטלה.");
                return;
            }

            SetStudentCardEditMode(false);
            SetStatus("עריכת פרטי החניך בוטלה.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Cancel edit student failed");
            SetStatus($"ביטול עריכת החניך נכשל: {ex.Message}");
        }
    }

    private void UpdateStudentRowCache(string stuId, string firstName, string lastName, string classId, string phone, string city, string remarks)
    {
        var normalizedStuId = (stuId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedStuId))
        {
            return;
        }

        var classDisplay = classId;
        if (CmbCardClass?.SelectedItem is DataRowView classRow)
        {
            classDisplay = ToClassDisplay(classRow["ClassName"]?.ToString(), classId);
        }

        if (GridStudents.SelectedItem is DataRowView selectedRow)
        {
            if (selectedRow.Row.Table.Columns.Contains("stuFname")) selectedRow["stuFname"] = firstName;
            if (selectedRow.Row.Table.Columns.Contains("stuLname")) selectedRow["stuLname"] = lastName;
            if (selectedRow.Row.Table.Columns.Contains("stuClassID")) selectedRow["stuClassID"] = classId;
            if (selectedRow.Row.Table.Columns.Contains("ClassName")) selectedRow["ClassName"] = classDisplay;
            if (selectedRow.Row.Table.Columns.Contains("stuPhone")) selectedRow["stuPhone"] = phone;
            if (selectedRow.Row.Table.Columns.Contains("stuCity")) selectedRow["stuCity"] = city;
            if (selectedRow.Row.Table.Columns.Contains("stuRemarks")) selectedRow["stuRemarks"] = remarks;
        }

        foreach (DataRow row in _studentsTable.Rows)
        {
            var rowStuId = (row["stuID"]?.ToString() ?? string.Empty).Trim();
            if (!string.Equals(rowStuId, normalizedStuId, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (_studentsTable.Columns.Contains("stuFname")) row["stuFname"] = firstName;
            if (_studentsTable.Columns.Contains("stuLname")) row["stuLname"] = lastName;
            if (_studentsTable.Columns.Contains("stuClassID")) row["stuClassID"] = classId;
            if (_studentsTable.Columns.Contains("ClassName")) row["ClassName"] = classDisplay;
            if (_studentsTable.Columns.Contains("stuPhone")) row["stuPhone"] = phone;
            if (_studentsTable.Columns.Contains("stuCity")) row["stuCity"] = city;
            if (_studentsTable.Columns.Contains("stuRemarks")) row["stuRemarks"] = remarks;
            break;
        }

        GridStudents.Items.Refresh();
    }

    private void BtnAddManualStudent_OnClick(object sender, RoutedEventArgs e)
    {
        var idInput = PromptText("הזן תעודת זהות (9 ספרות)", "הוספת חניך ידנית");
        if (string.IsNullOrWhiteSpace(idInput))
        {
            return;
        }

        var firstName = PromptText("הזן שם פרטי", "הוספת חניך ידנית");
        if (string.IsNullOrWhiteSpace(firstName))
        {
            SetStatus("שם פרטי הוא שדה חובה.");
            return;
        }

        var lastName = PromptText("הזן שם משפחה", "הוספת חניך ידנית");
        if (string.IsNullOrWhiteSpace(lastName))
        {
            SetStatus("שם משפחה הוא שדה חובה.");
            return;
        }

        var classCode = PromptClassSelectionForNewStudent();
        if (string.IsNullOrWhiteSpace(classCode))
        {
            SetStatus("הוספת חניך בוטלה: לא נבחרה כיתה.");
            return;
        }

        try
        {
            var success = _dbService.TryAddManualStudent(idInput, firstName, lastName, classCode, out var message);
            SetStatus(message);
            if (!success)
            {
                MessageBox.Show(message, "הוספת חניך ידנית", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            LoadStudentsData();
            LoadClassFilter();
            PerformSearch(_dbService.NormalizeStudentIdForDisplay(idInput));
            CmbCardClass.SelectedValue = classCode;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Manual student add failed");
            var detailedMessage = $"הוספת חניך ידנית נכשלה: {ex.Message}";
            SetStatus(detailedMessage);
            MessageBox.Show(detailedMessage, "הוספת חניך ידנית", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string PromptText(string prompt, string title)
    {
        var dialog = new InputDialog(prompt, title);
        var owner = Application.Current?.MainWindow;
        if (owner is not null && owner.IsLoaded && owner.IsVisible && !ReferenceEquals(owner, dialog))
        {
            dialog.Owner = owner;
        }
        else
        {
            dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        return dialog.ShowDialog() == true ? dialog.InputText.Trim() : string.Empty;
    }

    private string? PromptClassSelectionForNewStudent()
    {
        var classes = _dbService.GetClasses();
        foreach (DataRow pickerRow in classes.Rows)
        {
            var id = pickerRow["ClassID"]?.ToString()?.Trim() ?? string.Empty;
            if (id.Length is > 0 and < 4 && id.All(char.IsDigit))
            {
                id = id.PadLeft(4, '0');
                pickerRow["ClassID"] = id;
            }

            var displayName = ClassFormatter.ToDisplayName(pickerRow["ClassName"]?.ToString(), id);
            if (!string.IsNullOrEmpty(displayName))
            {
                pickerRow["ClassName"] = displayName;
            }
        }

        if (!classes.AsEnumerable().Any(r => (r["ClassID"]?.ToString()?.Trim()) == ClassFormatter.TeacherCode))
        {
            var teacherRow = classes.NewRow();
            teacherRow["ClassID"] = ClassFormatter.TeacherCode;
            teacherRow["ClassName"] = ClassFormatter.TeacherDisplay;
            classes.Rows.Add(teacherRow);
        }

        if (classes.Rows.Count == 0)
        {
            MessageBox.Show("לא נמצאו כיתות בטבלת Classes.", "הוספת חניך", MessageBoxButton.OK, MessageBoxImage.Warning);
            return null;
        }

        var classPicker = new ComboBox
        {
            Width = 220,
            Height = 32,
            Margin = new Thickness(0, 8, 0, 0),
            DisplayMemberPath = "ClassName",
            SelectedValuePath = "ClassID",
            ItemsSource = classes.DefaultView,
            SelectedIndex = 0
        };

        var okButton = new Button
        {
            Content = "אישור",
            Width = 80,
            Height = 30,
            Margin = new Thickness(0, 0, 8, 0),
            IsDefault = true
        };

        var cancelButton = new Button
        {
            Content = "ביטול",
            Width = 80,
            Height = 30,
            IsCancel = true
        };

        var buttonsPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 14, 0, 0)
        };
        buttonsPanel.Children.Add(okButton);
        buttonsPanel.Children.Add(cancelButton);

        var dialogContent = new StackPanel
        {
            Margin = new Thickness(16)
        };
        dialogContent.Children.Add(new TextBlock
        {
            Text = "בחר כיתה לחניך החדש:",
            FontWeight = FontWeights.SemiBold
        });
        dialogContent.Children.Add(classPicker);
        dialogContent.Children.Add(buttonsPanel);

        var dialog = new Window
        {
            Title = "בחירת כיתה",
            Width = 290,
            Height = 180,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = Application.Current.MainWindow,
            FlowDirection = FlowDirection.RightToLeft,
            Content = dialogContent
        };

        okButton.Click += (_, _) => dialog.DialogResult = true;

        if (dialog.ShowDialog() != true)
        {
            return null;
        }

        var selectedClassId = classPicker.SelectedValue?.ToString();
        if (string.IsNullOrWhiteSpace(selectedClassId) && classPicker.SelectedItem is DataRowView classRow)
        {
            selectedClassId = classRow["ClassID"]?.ToString();
        }

        return NormalizeClassIdForStorage(selectedClassId);
    }

    private void TxtStudentSearch_OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
        {
            return;
        }

        _searchDebounceTimer.Stop();
        PerformSearch();
        e.Handled = true;
    }

    private void ChkExemptCurrentYear_OnCheckedChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressExemptionEvents)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(_selectedStudentId))
        {
            SetStatus("יש לבחור חניך לפני עדכון פטור.");
            _suppressExemptionEvents = true;
            ChkExemptCurrentYear.IsChecked = false;
            _suppressExemptionEvents = false;
            return;
        }

        try
        {
            var studentIdToKeep = _selectedStudentId;
            var shouldExempt = ChkExemptCurrentYear.IsChecked == true;
            var success = _dbService.SetStudentExemptionForCurrentSchoolYear(studentIdToKeep, shouldExempt, out var message);
            SetStatus(message);
            if (!success)
            {
                _suppressExemptionEvents = true;
                ChkExemptCurrentYear.IsChecked = _dbService.IsStudentExemptForCurrentSchoolYear(studentIdToKeep);
                _suppressExemptionEvents = false;
                return;
            }

            LoadStudentsData();
            LoadClassFilter();
            PerformSearch(studentIdToKeep);
            _selectedStudentId = studentIdToKeep;
            RefreshStudentTables(studentIdToKeep);

            Dispatcher.BeginInvoke(new Action(() =>
            {
                var rowToSelect = GridStudents.Items
                    .OfType<DataRowView>()
                    .FirstOrDefault(r => string.Equals(r["stuID"]?.ToString(), studentIdToKeep, StringComparison.OrdinalIgnoreCase));
                if (rowToSelect is null)
                {
                    return;
                }

                if (!ReferenceEquals(GridStudents.SelectedItem, rowToSelect))
                {
                    GridStudents.SelectedItem = rowToSelect;
                }

                GridStudents.ScrollIntoView(rowToSelect);
            }), System.Windows.Threading.DispatcherPriority.Background);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Update exemption failed");
            SetStatus("עדכון סטטוס הפטור נכשל.");
        }
    }

    private void BtnExtendLoan_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if ((sender as FrameworkElement)?.DataContext is not DataRowView row)
            {
                SetStatus("לא נבחרה השאלה להארכה.");
                return;
            }

            var barcode = GetRowText(row, "Barcode");
            if (string.IsNullOrWhiteSpace(barcode))
            {
                SetStatus("לא נמצא ברקוד לעדכון.");
                return;
            }

            DateTime? currentReturnDate = null;
            if (row.Row.Table.Columns.Contains("DateToReturn") && row["DateToReturn"] != DBNull.Value)
            {
                currentReturnDate = Convert.ToDateTime(row["DateToReturn"]);
            }

            var dialog = new ReturnDateDialog(currentReturnDate ?? DateTime.Today)
            {
                Owner = Application.Current.MainWindow
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            var newReturnDate = dialog.SelectedDate;
            var success = _dbService.TryExtendLoan(barcode, newReturnDate, out var message);
            SetStatus(message);
            MessageBox.Show(message, "הארכת השאלה", MessageBoxButton.OK, success ? MessageBoxImage.Information : MessageBoxImage.Warning);

            if (success && !string.IsNullOrWhiteSpace(_selectedStudentId))
            {
                RefreshStudentTables(_selectedStudentId);
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Extend loan from students view failed");
            var message = $"הארכת מועד ההחזרה נכשלה: {ex.Message}";
            SetStatus(message);
            MessageBox.Show(message, "הארכת השאלה", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnScrapLoanCopy_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if ((sender as FrameworkElement)?.DataContext is not DataRowView row)
            {
                SetStatus("לא נבחר עותק לגריטה.");
                return;
            }

            var barcode = GetRowText(row, "Barcode");
            if (string.IsNullOrWhiteSpace(barcode))
            {
                SetStatus("לא נמצא ברקוד לגריטה.");
                return;
            }

            var confirmation = MessageBox.Show(
                $"האם לגרוט את העותק {barcode}?\nפעולה זו תסמן את העותק כגרוט.",
                "גריטת עותק",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (confirmation != MessageBoxResult.Yes)
            {
                return;
            }

            var success = _dbService.TryMarkBookCopyAsScrapped(barcode, out var msg);
            SetStatus(msg);
            MessageBox.Show(msg, "גריטת עותק", MessageBoxButton.OK, success ? MessageBoxImage.Information : MessageBoxImage.Warning);

            if (!string.IsNullOrWhiteSpace(_selectedStudentId))
            {
                RefreshStudentTables(_selectedStudentId);
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Scrap loan copy from students view failed");
            var message = $"גריטת עותק נכשלה: {ex.Message}";
            SetStatus(message);
            MessageBox.Show(message, "גריטת עותק", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ClearStudentCard()
    {
        SetStudentCardEditMode(false);
        TxtCardStuId.Text = string.Empty;
        TxtCardFirstName.Text = string.Empty;
        TxtCardLastName.Text = string.Empty;
        if (CmbCardClass != null)
        {
            CmbCardClass.SelectedIndex = -1;
            CmbCardClass.Text = string.Empty;
        }

        TxtCardPhone.Text = string.Empty;
        TxtCardCity.Text = string.Empty;
        TxtCardRemarks.Text = string.Empty;
        _suppressExemptionEvents = true;
        ChkExemptCurrentYear.IsChecked = false;
        ChkExemptCurrentYear.IsEnabled = false;
        _suppressExemptionEvents = false;
        PaymentStatusBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FEE2E2"));
        TxtPaymentStatusBadge.Text = "✗ טרם שולם לשנה הנוכחית";
        TxtPaymentStatusBadge.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#991B1B"));
        GridActiveLoans.ItemsSource = null;
    }

    private void SetStudentCardEditMode(bool isEditMode)
    {
        _isStudentCardEditMode = isEditMode;

        if (TxtCardFirstName != null) TxtCardFirstName.IsReadOnly = !isEditMode;
        if (TxtCardLastName != null) TxtCardLastName.IsReadOnly = !isEditMode;
        if (TxtCardPhone != null) TxtCardPhone.IsReadOnly = !isEditMode;
        if (TxtCardCity != null) TxtCardCity.IsReadOnly = !isEditMode;
        if (TxtCardRemarks != null) TxtCardRemarks.IsReadOnly = !isEditMode;

        if (TxtCardStuId != null) TxtCardStuId.IsReadOnly = true;
        if (CmbCardClass != null) CmbCardClass.IsEnabled = isEditMode;

        var readOnlyBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F3F4F6"));
        var editableBrush = Brushes.White;
        if (TxtCardFirstName != null) TxtCardFirstName.Background = isEditMode ? editableBrush : readOnlyBrush;
        if (TxtCardLastName != null) TxtCardLastName.Background = isEditMode ? editableBrush : readOnlyBrush;
        if (TxtCardPhone != null) TxtCardPhone.Background = isEditMode ? editableBrush : readOnlyBrush;
        if (TxtCardCity != null) TxtCardCity.Background = isEditMode ? editableBrush : readOnlyBrush;
        if (TxtCardRemarks != null) TxtCardRemarks.Background = isEditMode ? editableBrush : readOnlyBrush;
        if (CmbCardClass != null) CmbCardClass.Background = isEditMode ? editableBrush : readOnlyBrush;

        if (BtnEditStudent != null)
        {
            BtnEditStudent.Content = isEditMode ? "שמור שינויים" : "ערוך";
        }

        if (BtnCancelEditStudent != null)
        {
            BtnCancelEditStudent.IsEnabled = isEditMode;
        }
    }

    private void UpdatePaymentStatusBadge(string stuId)
    {
        string currentYear;
        try
        {
            currentYear = _dbService.GetCurrentSchoolYear();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "UpdatePaymentStatusBadge school year");
            PaymentStatusBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FEF3C7"));
            TxtPaymentStatusBadge.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#92400E"));
            TxtPaymentStatusBadge.Text = "שנת הלימודים לא מוגדרת";
            return;
        }

        var isExempt = _dbService.IsStudentExemptForCurrentSchoolYear(stuId);

        if (isExempt)
        {
            PaymentStatusBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DBEAFE"));
            TxtPaymentStatusBadge.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E40AF"));
            TxtPaymentStatusBadge.Text = $"✓ פטור מתשלום ({currentYear})";
            return;
        }

        if (_dbService.IsPaymentCompletedForCurrentSchoolYear(stuId))
        {
            PaymentStatusBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DCFCE7"));
            TxtPaymentStatusBadge.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#166534"));
            TxtPaymentStatusBadge.Text = $"✓ שולם לשנת {currentYear}";
            return;
        }

        PaymentStatusBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FEE2E2"));
        TxtPaymentStatusBadge.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#991B1B"));
        TxtPaymentStatusBadge.Text = $"✗ טרם שולם לשנת {currentYear}";
    }

    private void SetStatus(string message)
    {
        TxtStatus.Text = message;
    }
}
