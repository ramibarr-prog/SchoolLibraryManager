using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Media;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace SchoolLibraryManager;

public partial class BooksView : UserControl
{
    private static readonly FontFamily BarcodeFontFamily = LoadBarcodeFontFamily();

    private static FontFamily LoadBarcodeFontFamily()
    {
        const string familyName = "Libre Barcode 39";
        var candidates = new[]
        {
            new Uri("pack://application:,,,/Fonts/", UriKind.Absolute),
            new Uri(Path.Combine(AppContext.BaseDirectory, "Fonts") + Path.DirectorySeparatorChar, UriKind.Absolute)
        };

        foreach (var baseUri in candidates)
        {
            try
            {
                var family = new FontFamily(baseUri, "./#" + familyName);
                var typeface = new Typeface(family, FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
                if (typeface.TryGetGlyphTypeface(out _))
                {
                    return family;
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log(ex, $"Barcode font load failed from {baseUri}");
            }
        }

        ErrorLogger.Log(new FileNotFoundException("Barcode font could not be loaded."), "Barcode font fallback to system font");
        return new FontFamily(familyName);
    }
    private const double PrintPageWidth = 793;
    private const double PrintPageHeight = 1122;
    private const int PrintColumns = 3;
    private const int PrintRows = 7;
    private const int PrintLabelsPerPage = PrintColumns * PrintRows;
    private const double PrintCellWidth = 69.9 * 96 / 25.4;
    private const double PrintCellHeight = 38.1 * 96 / 25.4;
    private const double PrintSheetMarginTop = 7.575 * 96 / 25.4;
    private const double PrintSheetMarginLeft = (210 - (3 * 69.9)) / 2 * 96 / 25.4;
    private const double PrintCalibrationOffsetX = 0;
    private const double PrintCalibrationOffsetY = 0;

    private readonly DatabaseService _databaseService;
    private DataTable _booksTable = new();
    private string? _selectedBookId;
    private bool _showScrappedCopies;
    private bool _isEditMode;
    private bool _isAddMode;
    private readonly DispatcherTimer _searchDebounceTimer;
    private int _copiesLoadVersion;
    private readonly ObservableCollection<PrintLabelItem> _labelQueue = [];
    private readonly HashSet<string> _labelQueueLookup = new(StringComparer.OrdinalIgnoreCase);
    private string _schoolName = "ספריית בית הספר";

    public BooksView(DatabaseService databaseService)
    {
        InitializeComponent();
        _databaseService = databaseService;
        DataObject.AddPastingHandler(txtBookYear, IntegerOnly_OnPasting);
        _searchDebounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(300)
        };
        _searchDebounceTimer.Tick += SearchDebounceTimer_Tick;
        Loaded += BooksView_OnLoaded;
    }

    private void SearchDebounceTimer_Tick(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        ApplyLiveFilter();
    }

    private void BooksView_OnLoaded(object sender, RoutedEventArgs e)
    {
        try
        {
            gridLabelQueue.ItemsSource = _labelQueue;
            cmbStartLabelNumber.ItemsSource = Enumerable.Range(0, 21).ToList();
            cmbStartLabelNumber.SelectedIndex = 0;
            _schoolName = _databaseService.GetSchoolName();
            LoadBooks();
            SetEditMode(false);
            txtBookSearch.Focus();
            UpdateLabelPreview(null);
            SetStatus("מסך כותרים ועותקים נטען בהצלחה.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Books view load failed");
            SetStatus("טעינת מסך כותרים ועותקים נכשלה.");
        }
    }

    private static bool IsScrappedCondition(string? condition)
    {
        var c = condition?.Trim();
        return c == "98" || c == "99";
    }

    private async void BtnRestoreScrappedCopy_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button)
        {
            return;
        }

        var barcode = (button.CommandParameter?.ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(barcode) || string.IsNullOrWhiteSpace(_selectedBookId))
        {
            return;
        }

        var confirmation = MessageBox.Show(
            $"האם לשחזר את עותק ברקוד {barcode} מגריטה ולהחזירו למלאי זמין?",
            "אישור שחזור עותק",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question,
            MessageBoxResult.No);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            var result = await Task.Run(() => _databaseService.TryRestoreScrappedCopy(barcode, out var message)
                ? (Success: true, Message: message)
                : (Success: false, Message: message));

            SetStatus(result.Message);
            if (!result.Success)
            {
                return;
            }

            await LoadCopiesByBookIdAsync(_selectedBookId);
            LoadBooks();
            ReselectBook(_selectedBookId);
            MessageBox.Show(string.IsNullOrWhiteSpace(result.Message) ? $"העותק {barcode} שוחזר למלאי בהצלחה." : result.Message, "שחזור עותק", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"Restore copy failed for barcode '{barcode}'");
            SetStatus("שחזור העותק נכשל.");
        }
    }

    private void LoadBooks()
    {
        _booksTable = _databaseService.SearchBooks(string.Empty);
        ApplyLiveFilter();
    }

    private void TxtBookSearch_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        _searchDebounceTimer.Stop();
        _searchDebounceTimer.Start();
    }

    private void ApplyLiveFilter()
    {
        var view = _booksTable.DefaultView;
        var query = (txtBookSearch.Text ?? string.Empty).Trim().Replace("'", "''").Replace("[", "[[]").Replace("*", "[*]").Replace("%", "[%]");
        if (string.IsNullOrWhiteSpace(query))
        {
            view.RowFilter = string.Empty;
        }
        else
        {
            view.RowFilter = $"bkName LIKE '%{query}%' OR bkAuthor LIKE '%{query}%' OR bkID LIKE '%{query}%'";
        }

        gridBooks.ItemsSource = view;
        txtResultsCount.Text = $"נמצאו {view.Count} כותרים";

        if (_isAddMode)
        {
            return;
        }

        if (view.Count > 0)
        {
            if (gridBooks.SelectedItem is not DataRowView)
            {
                gridBooks.SelectedIndex = 0;
            }

            return;
        }

        _selectedBookId = null;
        Interlocked.Increment(ref _copiesLoadVersion);
        btnEditBook.IsEnabled = false;
        ClearBookDetails();
        ClearCopiesGrid();
        txtSelectedBookHeader.Text = "כותר נבחר: -";
        UpdateBorrowersReportButtonState();
        SetEditMode(false);
    }

    private async void GridBooks_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (gridBooks.SelectedItem is not DataRowView row)
        {
            if (_isAddMode)
            {
                return;
            }

            _selectedBookId = null;
            Interlocked.Increment(ref _copiesLoadVersion);
            btnEditBook.IsEnabled = false;
            ClearBookDetails();
            ClearCopiesGrid();
            txtSelectedBookHeader.Text = "כותר נבחר: -";
            UpdateBorrowersReportButtonState();
            SetEditMode(false);
            return;
        }

        _isAddMode = false;
        _selectedBookId = row["bkID"]?.ToString();
        LoadBookDetails(row);
        txtSelectedBookHeader.Text = $"כותר נבחר: {row["bkName"]?.ToString() ?? "-"}";
        SetEditMode(false);
        btnEditBook.IsEnabled = !string.IsNullOrWhiteSpace(_selectedBookId);
        btnScrapSelectedTitle.IsEnabled = !string.IsNullOrWhiteSpace(_selectedBookId);
        UpdateBorrowersReportButtonState();
        await LoadCopiesByBookIdAsync(_selectedBookId);
    }

    private void LoadBookDetails(DataRowView row)
    {
        txtBookTitle.Text = row["bkName"]?.ToString() ?? string.Empty;
        txtBookAuthor.Text = row["bkAuthor"]?.ToString() ?? string.Empty;
        txtBookSubject.Text = row["SubjectName"]?.ToString() ?? string.Empty;
        txtBookPrice.Text = row.Row.Table.Columns.Contains("bkPrice") ? row["bkPrice"]?.ToString() ?? string.Empty : string.Empty;
        SplitPublisherAndYear(row["bkPublish"]?.ToString() ?? string.Empty, out var publisher, out var year);
        txtBookPublisher.Text = publisher;
        txtBookYear.Text = year;
    }

    private void BtnEditBook_OnClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_selectedBookId))
        {
            SetStatus("יש לבחור כותר לפני עריכה.");
            return;
        }

        SetEditMode(true);
        txtBookTitle.Focus();
        txtBookTitle.SelectAll();
    }

    private void BtnCancelEdit_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isAddMode)
        {
            _isAddMode = false;
            LoadBooks();
            if (gridBooks.Items.Count > 0)
            {
                gridBooks.SelectedIndex = 0;
            }

            SetEditMode(false);
            return;
        }

        if (gridBooks.SelectedItem is DataRowView row)
        {
            LoadBookDetails(row);
        }

        SetEditMode(false);
    }

    private void BtnSaveBook_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isAddMode)
        {
            SaveNewBook();
            return;
        }

        if (string.IsNullOrWhiteSpace(_selectedBookId))
        {
            SetStatus("יש לבחור כותר לפני שמירת שינויים.");
            return;
        }

        if (string.IsNullOrWhiteSpace(txtBookTitle.Text))
        {
            SetStatus("חובה להזין כותר.");
            txtBookTitle.Focus();
            return;
        }

        try
        {
            var yearInput = (txtBookYear.Text ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(yearInput))
            {
                var maxYear = DateTime.Now.Year + 1;
                if (yearInput.Length != 4 || !int.TryParse(yearInput, NumberStyles.None, CultureInfo.InvariantCulture, out var publishYear) || publishYear < 1800 || publishYear > maxYear)
                {
                    SetStatus("נא להזין שנת הוצאה תקינה בת 4 ספרות");
                    txtBookYear.Focus();
                    txtBookYear.SelectAll();
                    return;
                }
            }

            var priceInput = (txtBookPrice.Text ?? string.Empty).Trim();
            var normalizedPrice = string.Empty;
            if (!string.IsNullOrWhiteSpace(priceInput))
            {
                if (!decimal.TryParse(priceInput, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var parsedPrice) || parsedPrice < 0)
                {
                    SetStatus("נא להזין מחיר חוקי וחיובי");
                    txtBookPrice.Focus();
                    txtBookPrice.SelectAll();
                    return;
                }

                normalizedPrice = parsedPrice.ToString(CultureInfo.InvariantCulture);
            }

            var success = _databaseService.UpdateBook(
                _selectedBookId,
                txtBookTitle.Text,
                txtBookAuthor.Text,
                txtBookSubject.Text,
                txtBookPublisher.Text,
                yearInput,
                normalizedPrice,
                out var message);

            SetStatus(message);
            if (!success)
            {
                return;
            }

            SystemSounds.Asterisk.Play();
            var currentId = _selectedBookId;
            LoadBooks();
            ReselectBook(currentId);
            SetEditMode(false);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Save edited book failed");
            SetStatus("שמירת השינויים נכשלה.");
        }
    }

    private void BtnAddNewBook_OnClick(object sender, RoutedEventArgs e)
    {
        _isAddMode = true;
        _selectedBookId = null;
        Interlocked.Increment(ref _copiesLoadVersion);
        gridBooks.SelectedItem = null;
        gridBooks.SelectedIndex = -1;
        ClearBookDetails();
        ClearCopiesGrid();
        txtSelectedBookHeader.Text = "הוספת כותר חדש לקטלוג";
        btnEditBook.IsEnabled = false;
        SetEditMode(true);
        txtBookTitle.Focus();
    }

    private void SaveNewBook()
    {
        if (string.IsNullOrWhiteSpace(txtBookTitle.Text))
        {
            SetStatus("חובה להזין כותר.");
            txtBookTitle.Focus();
            return;
        }

        var yearInput = (txtBookYear.Text ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(yearInput))
        {
            var maxYear = DateTime.Now.Year + 1;
            if (yearInput.Length != 4 || !int.TryParse(yearInput, NumberStyles.None, CultureInfo.InvariantCulture, out var publishYear) || publishYear < 1800 || publishYear > maxYear)
            {
                SetStatus("נא להזין שנת הוצאה תקינה בת 4 ספרות");
                txtBookYear.Focus();
                txtBookYear.SelectAll();
                return;
            }
        }

        var subjectId = ResolveSubjectIdByName(txtBookSubject.Text);
        if (!string.IsNullOrWhiteSpace(txtBookSubject.Text) && string.IsNullOrWhiteSpace(subjectId))
        {
            SetStatus("המקצוע שנבחר אינו קיים במערכת.");
            txtBookSubject.Focus();
            txtBookSubject.SelectAll();
            return;
        }

        var publisherValue = (txtBookPublisher.Text ?? string.Empty).Trim();
        var priceInput = (txtBookPrice.Text ?? string.Empty).Trim();
        var normalizedPrice = string.Empty;
        if (!string.IsNullOrWhiteSpace(priceInput))
        {
            if (!decimal.TryParse(priceInput, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var parsedPrice) || parsedPrice < 0)
            {
                SetStatus("נא להזין מחיר חוקי וחיובי");
                txtBookPrice.Focus();
                txtBookPrice.SelectAll();
                return;
            }

            normalizedPrice = parsedPrice.ToString(CultureInfo.InvariantCulture);
        }

        var publishCombined = string.IsNullOrWhiteSpace(yearInput)
            ? publisherValue
            : (string.IsNullOrWhiteSpace(publisherValue) ? yearInput : $"{publisherValue} ({yearInput})");

        var newBookId = GenerateNextBookId();

        var subjectIdValue = 0;
        if (!string.IsNullOrWhiteSpace(subjectId)
            && !int.TryParse(subjectId, NumberStyles.None, CultureInfo.InvariantCulture, out subjectIdValue))
        {
            MessageBox.Show("שגיאה בהמרת מזהה המקצוע. הקלט אינו תקין ולכן הכותר לא נשמר.", "שגיאת קלט", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        try
        {
            var success = _databaseService.TryAddBook(
                newBookId,
                txtBookTitle.Text,
                txtBookAuthor.Text,
                normalizedPrice,
                publishCombined,
                subjectIdValue,
                0,
                out var message);

            if (!success)
            {
                SetStatus(message);
                return;
            }

            _isAddMode = false;
            LoadBooks();
            ReselectBook(newBookId);
            SetEditMode(false);
            SetStatus("הכותר נוסף בהצלחה לקטלוג.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Add new book failed in books view");
            SetStatus("הוספת הכותר נכשלה.");
        }
    }

    private void Price_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        if (sender is not TextBox textBox)
        {
            e.Handled = true;
            return;
        }

        e.Handled = !IsValidPriceInputAfterEdit(textBox, e.Text);
    }

    private void Price_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Space)
        {
            e.Handled = true;
        }
    }

    private void Price_OnPasting(object sender, DataObjectPastingEventArgs e)
    {
        if (sender is not TextBox textBox)
        {
            e.CancelCommand();
            return;
        }

        if (!e.SourceDataObject.GetDataPresent(DataFormats.Text))
        {
            e.CancelCommand();
            return;
        }

        var pastedText = e.SourceDataObject.GetData(DataFormats.Text) as string ?? string.Empty;
        if (!IsValidPriceInputAfterEdit(textBox, pastedText))
        {
            e.CancelCommand();
        }
    }

    private static bool IsValidPriceInputAfterEdit(TextBox textBox, string insertedText)
    {
        if (insertedText.Any(char.IsWhiteSpace) || insertedText.Contains('-', StringComparison.Ordinal))
        {
            return false;
        }

        var current = textBox.Text ?? string.Empty;
        var selectionStart = textBox.SelectionStart;
        var selectionLength = textBox.SelectionLength;
        var next = current.Remove(selectionStart, selectionLength).Insert(selectionStart, insertedText);

        if (string.IsNullOrEmpty(next))
        {
            return true;
        }

        var dotCount = 0;
        foreach (var ch in next)
        {
            if (char.IsDigit(ch))
            {
                continue;
            }

            if (ch == '.')
            {
                dotCount++;
                if (dotCount <= 1)
                {
                    continue;
                }
            }

            return false;
        }

        return true;
    }

    private string GenerateNextBookId()
    {
        var maxId = 0;
        foreach (DataRow row in _booksTable.Rows)
        {
            if (int.TryParse(row["bkID"]?.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > maxId)
            {
                maxId = parsed;
            }
        }

        return (maxId + 1).ToString(CultureInfo.InvariantCulture);
    }

    private string ResolveSubjectIdByName(string? subjectName)
    {
        var cleanSubject = (subjectName ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanSubject))
        {
            return string.Empty;
        }

        var subjects = _databaseService.GetSubjects();
        var match = subjects.AsEnumerable()
            .FirstOrDefault(r => string.Equals(r["SubjectName"]?.ToString()?.Trim(), cleanSubject, StringComparison.OrdinalIgnoreCase));

        return match is null ? string.Empty : match["SubjectID"]?.ToString() ?? string.Empty;
    }

    private void IntegerOnly_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        if (e.Text.Any(ch => !char.IsDigit(ch)))
        {
            e.Handled = true;
        }
    }

    private void IntegerOnly_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Space)
        {
            e.Handled = true;
        }
    }

    private void IntegerOnly_OnPasting(object sender, DataObjectPastingEventArgs e)
    {
        if (sender is not TextBox textBox)
        {
            e.CancelCommand();
            return;
        }

        if (!e.SourceDataObject.GetDataPresent(DataFormats.Text))
        {
            e.CancelCommand();
            return;
        }

        var pastedText = e.SourceDataObject.GetData(DataFormats.Text) as string ?? string.Empty;
        var current = textBox.Text ?? string.Empty;
        var selectionStart = textBox.SelectionStart;
        var selectionLength = textBox.SelectionLength;
        var next = current.Remove(selectionStart, selectionLength).Insert(selectionStart, pastedText);

        if (next.Any(ch => !char.IsDigit(ch)))
        {
            e.CancelCommand();
        }
    }

    private void BtnAddCopyToInventory_OnClick(object sender, RoutedEventArgs e)
    {
        AddCopyToInventory();
    }

    private void TxtNewCopyBarcode_OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
        {
            return;
        }

        AddCopyToInventory();
        e.Handled = true;
    }

    private void AddCopyToInventory()
    {
        if (string.IsNullOrWhiteSpace(_selectedBookId))
        {
            SetStatus("יש לבחור כותר לפני הוספת עותק.");
            return;
        }

        var barcode = (txtNewCopyBarcode.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(barcode) || !barcode.All(char.IsDigit))
        {
            SetStatus("ברקוד עותק חייב להיות מספרי.");
            txtNewCopyBarcode.Focus();
            txtNewCopyBarcode.SelectAll();
            return;
        }

        try
        {
            var success = _databaseService.TryAddBookCopy(_selectedBookId, barcode, string.Empty, out var message);
            SetStatus(message);
            if (!success)
            {
                if (message.Contains("כבר קיים", StringComparison.OrdinalIgnoreCase))
                {
                    SystemSounds.Hand.Play();
                }

                txtNewCopyBarcode.Focus();
                txtNewCopyBarcode.SelectAll();
                return;
            }

            txtNewCopyBarcode.Clear();
            txtNewCopyBarcode.Focus();
            _ = LoadCopiesByBookIdAsync(_selectedBookId);
            LoadBooks();
            ReselectBook(_selectedBookId);
            SystemSounds.Asterisk.Play();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Add inventory copy failed");
            SetStatus("הוספת העותק נכשלה.");
        }
    }

    private async Task LoadCopiesByBookIdAsync(string? bookId)
    {
        var cleanId = (bookId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanId))
        {
            ClearCopiesGrid();
            return;
        }

        var requestVersion = Interlocked.Increment(ref _copiesLoadVersion);
        gridCopies.ItemsSource = null;
        txtCopiesSummary.Text = "טוען עותקים...";

        DataTable copies;
        try
        {
            copies = await Task.Run(() =>
            {
                var all = _databaseService.GetBookCopiesByBookId(cleanId, _showScrappedCopies);
                if (!_showScrappedCopies)
                {
                    return all;
                }

                var scrappedRows = all.AsEnumerable()
                    .Where(r => IsScrappedCondition(r["CopyCondition"]?.ToString()))
                    .ToList();
                return scrappedRows.Count > 0 ? scrappedRows.CopyToDataTable() : all.Clone();
            });
        }
        catch (Exception ex)
        {
            if (requestVersion != Volatile.Read(ref _copiesLoadVersion) || !string.Equals(_selectedBookId, cleanId, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            ErrorLogger.Log(ex, $"Load copies failed for book '{cleanId}'");
            SetStatus("טעינת העותקים נכשלה.");
            ClearCopiesGrid();
            return;
        }

        if (requestVersion != Volatile.Read(ref _copiesLoadVersion) || !string.Equals(_selectedBookId, cleanId, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        await Dispatcher.InvokeAsync(() =>
        {
            gridCopies.ItemsSource = copies.DefaultView;

            var total = copies.Rows.Count;
            var available = copies.AsEnumerable().Count(r =>
                string.Equals(r["CopyReturnStatus"]?.ToString(), "0", StringComparison.OrdinalIgnoreCase) &&
                !IsScrappedCondition(r["CopyCondition"]?.ToString()));
            var borrowed = copies.AsEnumerable().Count(r =>
                string.Equals(r["CopyReturnStatus"]?.ToString(), "1", StringComparison.OrdinalIgnoreCase) &&
                !IsScrappedCondition(r["CopyCondition"]?.ToString()));
            var scrapped = copies.AsEnumerable().Count(r => IsScrappedCondition(r["CopyCondition"]?.ToString()));

            txtCopiesSummary.Text = _showScrappedCopies
                ? $"עותקים עבור כותר זה (סה\"כ מוצג: {total} | זמינים: {available} | מושאלים: {borrowed} | גרוטים: {scrapped})"
                : $"עותקים עבור כותר זה (סה\"כ: {total} | זמינים: {available} | מושאלים: {borrowed})";
        });
    }

    private void ChkShowScrappedCopies_OnCheckedChanged(object sender, RoutedEventArgs e)
    {
        _showScrappedCopies = chkShowScrappedCopies.IsChecked == true;
        _ = LoadCopiesByBookIdAsync(_selectedBookId);
    }

    private async void BtnScrapCopy_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button)
        {
            return;
        }

        var barcode = (button.Tag?.ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(barcode))
        {
            barcode = (button.CommandParameter?.ToString() ?? string.Empty).Trim();
        }

        if (string.IsNullOrWhiteSpace(barcode) && button.DataContext is DataRowView row)
        {
            barcode = row["Barcode"]?.ToString()?.Trim() ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(barcode))
        {
            SetStatus("לא נמצא ברקוד תקין לגריטה.");
            return;
        }

        var selectedBookId = (_selectedBookId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(selectedBookId) && gridBooks.SelectedItem is DataRowView selectedBookRow)
        {
            selectedBookId = selectedBookRow["bkID"]?.ToString()?.Trim() ?? string.Empty;
        }

        var confirmation = MessageBox.Show(
            $"האם לסמן את עותק ברקוד {barcode} כבלאי/גרוט ולהוציאו מהמלאי?",
            "אישור גריטת עותק",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            var result = await Task.Run(() =>
            {
                var success = _databaseService.TryMarkBookCopyAsScrapped(barcode, out var message);
                return (Success: success, Message: message);
            });

            SetStatus(result.Message);
            if (!result.Success)
            {
                MessageBox.Show(result.Message, "גריטת עותק", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            MessageBox.Show(result.Message, "גריטת עותק", MessageBoxButton.OK, MessageBoxImage.Information);
            LoadBooks();

            if (!string.IsNullOrWhiteSpace(selectedBookId))
            {
                _selectedBookId = selectedBookId;
                await LoadCopiesByBookIdAsync(selectedBookId);
                ReselectBook(selectedBookId);
            }
            else
            {
                ClearCopiesGrid();
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"Scrap copy failed for barcode '{barcode}'");
            SetStatus("גריטת העותק נכשלה.");
            MessageBox.Show("גריטת העותק נכשלה.", "גריטת עותק", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnRestoreSelectedCopy_OnClick(object sender, RoutedEventArgs e)
    {
        if (gridCopies.SelectedItem is not DataRowView row)
        {
            SetStatus("יש לבחור עותק לשחזור.");
            return;
        }

        var barcode = row["Barcode"]?.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(barcode))
        {
            SetStatus("לעותק המסומן אין ברקוד תקין.");
            return;
        }

        BtnRestoreScrappedCopy_OnClick(new Button { CommandParameter = barcode }, e);
    }

    private void ClearCopiesGrid()
    {
        gridCopies.ItemsSource = null;
        txtCopiesSummary.Text = "עותקים עבור כותר זה (סה\"כ: 0 | זמינים: 0 | מושאלים: 0)";
    }

    private void SetEditMode(bool enabled)
    {
        _isEditMode = enabled;

        txtBookTitle.IsReadOnly = !enabled;
        txtBookAuthor.IsReadOnly = !enabled;
        txtBookSubject.IsReadOnly = !enabled;
        txtBookPublisher.IsReadOnly = !enabled;
        txtBookYear.IsReadOnly = !enabled;
        txtBookPrice.IsReadOnly = !enabled;

        var brush = enabled ? Brushes.White : (Brush)new BrushConverter().ConvertFromString("#F3F4F6")!;
        txtBookTitle.Background = brush;
        txtBookAuthor.Background = brush;
        txtBookSubject.Background = brush;
        txtBookPublisher.Background = brush;
        txtBookYear.Background = brush;
        txtBookPrice.Background = brush;

        btnEditBook.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        btnSaveBook.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        btnCancelEdit.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        btnScrapSelectedTitle.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        btnSaveBook.Content = _isAddMode ? "שמור כותר חדש 💾" : "שמור שינויים 💾";
        btnCancelEdit.Content = _isAddMode ? "ביטול הוספה ❌" : "ביטול ❌";
    }

    private void ClearBookDetails()
    {
        txtBookTitle.Clear();
        txtBookAuthor.Clear();
        txtBookSubject.Clear();
        txtBookPublisher.Clear();
        txtBookYear.Clear();
        txtBookPrice.Clear();
        btnScrapSelectedTitle.IsEnabled = false;
        UpdateBorrowersReportButtonState();
    }

    private void UpdateBorrowersReportButtonState()
    {
        var reportButton = FindName("btnTitleBorrowersReport") as Button;
        if (reportButton is not null)
        {
            reportButton.IsEnabled = !string.IsNullOrWhiteSpace(_selectedBookId);
        }
    }

    private async void BtnScrapSelectedTitle_OnClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_selectedBookId))
        {
            SetStatus("יש לבחור כותר לפני גריטה.");
            return;
        }

        var confirmation = MessageBox.Show(
            $"האם לגרוט את הכותר הנבחר ואת כל העותקים שלו?\nפעולה זו תסגור השאלות פעילות ותסמן את כל העותקים כבלאי.",
            "אישור גריטת כותר מלא",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            var selectedId = _selectedBookId;
            var result = await Task.Run(() => _databaseService.TryMarkBookTitleAsScrapped(selectedId, out var message)
                ? (Success: true, Message: message)
                : (Success: false, Message: message));

            SetStatus(result.Message);
            if (!result.Success)
            {
                return;
            }

            await LoadCopiesByBookIdAsync(_selectedBookId);
            LoadBooks();
            ReselectBook(_selectedBookId);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"Scrap title failed for bkID '{_selectedBookId}'");
            SetStatus("גריטת הכותר נכשלה.");
        }
    }

    private void BtnAddExistingTitleCopiesToQueue_OnClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_selectedBookId))
        {
            SetStatus("יש לבחור כותר לפני הוספת עותקים לתור המדבקות.");
            return;
        }

        var dt = _databaseService.GetCopiesForBarcodePrintingByBookId(_selectedBookId);
        var added = AddLabelItemsFromTable(dt);
        SetStatus(added > 0 ? $"נוספו {added} עותקים קיימים לתור המדבקות." : "לא נוספו עותקים (ייתכן שכבר קיימים בתור).");
    }

    private void BtnCreateNewCopiesAndQueue_OnClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_selectedBookId))
        {
            SetStatus("יש לבחור כותר לפני יצירת עותקים חדשים.");
            return;
        }

        if (!int.TryParse((txtNewCopiesToCreate.Text ?? string.Empty).Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var quantity) || quantity <= 0)
        {
            SetStatus("יש להזין כמות עותקים חוקית ליצירה.");
            txtNewCopiesToCreate.Focus();
            txtNewCopiesToCreate.SelectAll();
            return;
        }

        var success = _databaseService.TryCreateCopiesForBook(_selectedBookId, quantity, out var createdBarcodes, out var message);
        if (!success)
        {
            SetStatus(message);
            return;
        }

        var createdDt = _databaseService.GetCopiesForBarcodePrinting(createdBarcodes);
        var added = AddLabelItemsFromTable(createdDt);
        txtNewCopiesToCreate.Clear();
        _ = LoadCopiesByBookIdAsync(_selectedBookId);
        LoadBooks();
        ReselectBook(_selectedBookId);
        SetStatus($"{message} נוספו {added} מדבקות לתור.");
    }

    private void BtnAddSingleBarcodeToQueue_OnClick(object sender, RoutedEventArgs e)
    {
        var barcode = (txtSingleBarcodeToQueue.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(barcode))
        {
            SetStatus("יש להזין ברקוד להוספה.");
            return;
        }

        var dt = _databaseService.GetCopiesForBarcodePrinting([barcode]);
        if (dt.Rows.Count > 0 && _labelQueueLookup.Contains(barcode))
        {
            SetStatus($"הברקוד {barcode} כבר נמצא בתור ההדפסה.");
            return;
        }

        var added = AddLabelItemsFromTable(dt);
        if (added == 0)
        {
            SetStatus($"הברקוד {barcode} לא נמצא במלאי הספרייה.");
            return;
        }

        txtSingleBarcodeToQueue.Clear();
        SetStatus($"הברקוד {barcode} נוסף לתור המדבקות.");
    }

    private void BtnAddSelectedCopyToLabelQueue_OnClick(object sender, RoutedEventArgs e)
    {
        if (gridCopies.SelectedItem is not DataRowView row)
        {
            SetStatus("יש לסמן עותק לפני הוספה לתור המדבקות.");
            return;
        }

        var barcode = row["Barcode"]?.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(barcode))
        {
            SetStatus("העותק המסומן אינו כולל ברקוד תקין.");
            return;
        }

        var dt = _databaseService.GetCopiesForBarcodePrinting([barcode]);
        var added = AddLabelItemsFromTable(dt);
        SetStatus(added > 0 ? "העותק המסומן נוסף לתור המדבקות." : "העותק כבר קיים בתור המדבקות.");
    }

    private void BtnTitleBorrowersReport_OnClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_selectedBookId))
        {
            SetStatus("יש לבחור כותר לפני הפקת דוח משאילים.");
            return;
        }

        try
        {
            var borrowers = _databaseService.GetActiveBorrowersForBook(_selectedBookId);
            if (borrowers.Rows.Count == 0)
            {
                SetStatus("אין כרגע משאילים פעילים לכותר זה.");
                return;
            }

            var selectedBookTitle = gridBooks.SelectedItem is DataRowView row
                ? row["bkName"]?.ToString()?.Trim() ?? string.Empty
                : string.Empty;

            var document = BuildTitleBorrowersReportDocument(selectedBookTitle, _selectedBookId, borrowers);
            var printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                printDialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, "Active Borrowers By Title");
                SetStatus("דוח משאילים לכותר נשלח להדפסה.");
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"Title borrowers report failed for bkID '{_selectedBookId}'");
            SetStatus("הפקת דוח משאילים לכותר נכשלה.");
        }
    }

    private void GridLabelQueue_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateLabelPreview(gridLabelQueue.SelectedItem as PrintLabelItem ?? _labelQueue.FirstOrDefault());
    }

    private void BtnRemoveLabelQueueItem_OnClick(object sender, RoutedEventArgs e)
    {
        var rawBarcode = (sender as FrameworkElement)?.Tag?.ToString()
            ?? (sender as Button)?.CommandParameter?.ToString()
            ?? string.Empty;

        var barcode = rawBarcode.Trim();
        if (string.IsNullOrWhiteSpace(barcode))
        {
            return;
        }

        var removedItems = _labelQueue
            .Where(x => string.Equals((x.Barcode ?? string.Empty).Trim(), barcode, StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var queueItem in removedItems)
        {
            _labelQueue.Remove(queueItem);
        }

        _labelQueueLookup.Remove(barcode);

        UpdateLabelPreview(gridLabelQueue.SelectedItem as PrintLabelItem ?? _labelQueue.FirstOrDefault());
        SetStatus($"הברקוד {barcode} הוסר מתור המדבקות.");
    }

    private void BtnClearLabelQueue_OnClick(object sender, RoutedEventArgs e)
    {
        _labelQueue.Clear();
        _labelQueueLookup.Clear();
        gridLabelQueue.SelectedItem = null;
        UpdateLabelPreview(null);
        SetStatus("תור המדבקות נוקה.");
    }

    private void BtnPrintLabelQueue_OnClick(object sender, RoutedEventArgs e)
    {
        if (_labelQueue.Count == 0)
        {
            SetStatus("תור המדבקות ריק.");
            return;
        }

        var startLabel = cmbStartLabelNumber.SelectedItem is int labelNumber ? labelNumber : 0;
        var skip = Math.Max(0, startLabel);

        var printDialog = new PrintDialog();
        if (printDialog.ShowDialog() != true)
        {
            return;
        }

        var guardButton = sender as System.Windows.Controls.Button;
        if (guardButton is not null)
        {
            guardButton.IsEnabled = false;
        }

        try
        {
            var pages = BuildPrintPages(_labelQueue.ToList(), skip);
            for (var i = 0; i < pages.Count; i++)
            {
                printDialog.PrintVisual(pages[i], $"Title Barcode Labels Page {i + 1}");
            }

            SetStatus($"נשלחו להדפסה {pages.Count} עמודים ({_labelQueue.Count} מדבקות, התחלה ממדבקה {startLabel}).");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Print label queue failed");
            SetStatus("הדפסת המדבקות נכשלה.");
        }
        finally
        {
            if (guardButton is not null)
            {
                guardButton.IsEnabled = true;
            }
        }
    }

    private int AddLabelItemsFromTable(DataTable dt)
    {
        var added = 0;
        foreach (DataRow row in dt.Rows)
        {
            var barcode = row.Table.Columns.Contains("Barcode") ? row["Barcode"]?.ToString()?.Trim() : null;
            if (string.IsNullOrWhiteSpace(barcode) || !_labelQueueLookup.Add(barcode))
            {
                continue;
            }

            _labelQueue.Add(new PrintLabelItem
            {
                Barcode = barcode,
                BookCode = row.Table.Columns.Contains("bkID") ? row["bkID"]?.ToString()?.Trim() ?? string.Empty : string.Empty,
                BookTitle = row.Table.Columns.Contains("bkName") ? row["bkName"]?.ToString()?.Trim() ?? string.Empty : string.Empty
            });

            added++;
        }

        if (gridLabelQueue.SelectedItem is null && _labelQueue.Count > 0)
        {
            gridLabelQueue.SelectedIndex = 0;
        }

        UpdateLabelPreview(gridLabelQueue.SelectedItem as PrintLabelItem ?? _labelQueue.FirstOrDefault());
        return added;
    }

    private void UpdateLabelPreview(PrintLabelItem? item)
    {
    }

    private FlowDocument BuildTitleBorrowersReportDocument(string? bookTitle, string? bookId, DataTable borrowers)
    {
        var title = string.IsNullOrWhiteSpace(bookTitle) ? "ללא כותרת" : bookTitle.Trim();
        var code = (bookId ?? string.Empty).Trim();

        var document = new FlowDocument
        {
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 12,
            PagePadding = new Thickness(40),
            FlowDirection = FlowDirection.RightToLeft
        };

        document.Blocks.Add(new Paragraph(new Run("דוח משאילים לכותר מסוים"))
        {
            FontSize = 18,
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(0, 0, 0, 8)
        });

        document.Blocks.Add(new Paragraph(new Run($"כותר: {title} | מזהה כותר: {code} | תאריך הפקה: {DateTime.Now:dd/MM/yyyy HH:mm}"))
        {
            Margin = new Thickness(0, 0, 0, 12)
        });

        var table = new Table();
        table.Columns.Add(new TableColumn { Width = new GridLength(110) });
        table.Columns.Add(new TableColumn { Width = new GridLength(180) });
        table.Columns.Add(new TableColumn { Width = new GridLength(120) });
        table.Columns.Add(new TableColumn { Width = new GridLength(140) });

        var rowGroup = new TableRowGroup();
        table.RowGroups.Add(rowGroup);

        var header = new TableRow();
        header.Cells.Add(new TableCell(new Paragraph(new Run("ברקוד עותק"))) { FontWeight = FontWeights.Bold });
        header.Cells.Add(new TableCell(new Paragraph(new Run("שם חניך"))) { FontWeight = FontWeights.Bold });
        header.Cells.Add(new TableCell(new Paragraph(new Run("כיתה"))) { FontWeight = FontWeights.Bold });
        header.Cells.Add(new TableCell(new Paragraph(new Run("תאריך השאלה"))) { FontWeight = FontWeights.Bold });
        rowGroup.Rows.Add(header);

        foreach (DataRow row in borrowers.Rows)
        {
            var barcode = row.Table.Columns.Contains("Barcode") ? row["Barcode"]?.ToString()?.Trim() ?? string.Empty : string.Empty;
            var studentName = row.Table.Columns.Contains("StudentName") ? row["StudentName"]?.ToString()?.Trim() ?? string.Empty : string.Empty;
            var className = row.Table.Columns.Contains("ClassName") ? row["ClassName"]?.ToString()?.Trim() ?? string.Empty : string.Empty;

            var loanDateText = string.Empty;
            if (row.Table.Columns.Contains("LoanDate") && row["LoanDate"] != DBNull.Value)
            {
                loanDateText = Convert.ToDateTime(row["LoanDate"], CultureInfo.CurrentCulture).ToString("dd/MM/yyyy", CultureInfo.CurrentCulture);
            }

            var dataRow = new TableRow();
            dataRow.Cells.Add(new TableCell(new Paragraph(new Run(string.IsNullOrWhiteSpace(barcode) ? "-" : barcode))));
            dataRow.Cells.Add(new TableCell(new Paragraph(new Run(string.IsNullOrWhiteSpace(studentName) ? "-" : studentName))));
            dataRow.Cells.Add(new TableCell(new Paragraph(new Run(string.IsNullOrWhiteSpace(className) ? "-" : className))));
            dataRow.Cells.Add(new TableCell(new Paragraph(new Run(string.IsNullOrWhiteSpace(loanDateText) ? "-" : loanDateText))));
            rowGroup.Rows.Add(dataRow);
        }

        document.Blocks.Add(table);
        SchoolBranding.AddLetterhead(document);
        return document;
    }

    private List<DrawingVisual> BuildPrintPages(IReadOnlyList<PrintLabelItem> items, int skip)
    {
        var totalSlots = skip + items.Count;
        var pageCount = (int)Math.Ceiling(totalSlots / (double)PrintLabelsPerPage);
        var pages = new List<DrawingVisual>(Math.Max(1, pageCount));

        for (var pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            var pageVisual = new DrawingVisual();
            using var context = pageVisual.RenderOpen();
            context.DrawRectangle(Brushes.White, null, new Rect(0, 0, PrintPageWidth, PrintPageHeight));

            for (var slotInPage = 0; slotInPage < PrintLabelsPerPage; slotInPage++)
            {
                var absoluteSlot = (pageIndex * PrintLabelsPerPage) + slotInPage;
                if (absoluteSlot < skip)
                {
                    continue;
                }

                var labelIndex = absoluteSlot - skip;
                if (labelIndex < 0 || labelIndex >= items.Count)
                {
                    continue;
                }

                var row = slotInPage / PrintColumns;
                var col = slotInPage % PrintColumns;
                var cellRect = new Rect(PrintSheetMarginLeft + PrintCalibrationOffsetX + (col * PrintCellWidth), PrintSheetMarginTop + PrintCalibrationOffsetY + (row * PrintCellHeight), PrintCellWidth, PrintCellHeight);
                DrawLabel(context, cellRect, items[labelIndex]);
            }

            pages.Add(pageVisual);
        }

        return pages;
    }

    private void DrawLabel(DrawingContext context, Rect cellRect, PrintLabelItem item)
    {
        var labelRect = new Rect(cellRect.X + 4, cellRect.Y + 4, cellRect.Width - 8, cellRect.Height - 8);
        context.DrawRectangle(Brushes.White, new Pen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1D5DB")), 1), labelRect);

        DrawCenteredText(context, string.IsNullOrWhiteSpace(_schoolName) ? "ספריית בית הספר" : _schoolName, 10.5, FontWeights.Bold, new Rect(labelRect.X + 6, labelRect.Y + 3, labelRect.Width - 12, 14));

        var title = string.IsNullOrWhiteSpace(item.BookTitle) ? "-" : item.BookTitle;
        DrawCenteredText(context, title, 9.5, FontWeights.SemiBold, new Rect(labelRect.X + 6, labelRect.Y + 17, labelRect.Width - 12, 28));

        var barcodeRect = new Rect(labelRect.X + 8, labelRect.Y + 42, labelRect.Width - 16, 54);
        DrawCenteredText(context, WrapBarcodeForCode39(item.Barcode), 34, FontWeights.Normal, barcodeRect, BarcodeFontFamily, FlowDirection.LeftToRight);

        context.DrawRectangle(Brushes.White, null, new Rect(labelRect.X + 2, labelRect.Bottom - 28, labelRect.Width - 4, 26));
        DrawCenteredText(context, item.Barcode, 15, FontWeights.Bold, new Rect(labelRect.X + 6, labelRect.Bottom - 27, labelRect.Width - 12, 24), new FontFamily("Segoe UI"), FlowDirection.LeftToRight);
    }

    private static void DrawCenteredText(DrawingContext context, string? text, double fontSize, FontWeight weight, Rect target, FontFamily? fontFamily = null, FlowDirection flowDirection = FlowDirection.RightToLeft)
    {
        var value = string.IsNullOrWhiteSpace(text) ? "-" : text.Trim();
        var formatted = new FormattedText(
            value,
            CultureInfo.CurrentCulture,
            flowDirection,
            new Typeface(fontFamily ?? new FontFamily("Segoe UI"), FontStyles.Normal, weight, FontStretches.Normal),
            fontSize,
            Brushes.Black,
            1.0)
        {
            MaxTextWidth = target.Width,
            MaxTextHeight = target.Height,
            TextAlignment = TextAlignment.Center,
            Trimming = TextTrimming.CharacterEllipsis
        };

        var y = target.Y + Math.Max(0, (target.Height - formatted.Height) / 2);
        context.DrawText(formatted, new Point(target.X, y));
    }

    private static string WrapBarcodeForCode39(string? rawBarcode)
    {
        var value = (rawBarcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            value = "000000000";
        }

        var sanitized = new string(value
            .ToUpperInvariant()
            .Where(ch => char.IsLetterOrDigit(ch) || ch is '-' or '.' or ' ' or '$' or '/' or '+' or '%')
            .ToArray());

        if (string.IsNullOrWhiteSpace(sanitized))
        {
            sanitized = "000000000";
        }

        return $"*{sanitized}*";
    }

    private sealed class PrintLabelItem
    {
        public string Barcode { get; init; } = string.Empty;
        public string BookTitle { get; init; } = string.Empty;
        public string BookCode { get; init; } = string.Empty;
    }

    private void ReselectBook(string? bookId)
    {
        if (string.IsNullOrWhiteSpace(bookId))
        {
            return;
        }

        for (var i = 0; i < gridBooks.Items.Count; i++)
        {
            if (gridBooks.Items[i] is DataRowView row && string.Equals(row["bkID"]?.ToString(), bookId, StringComparison.OrdinalIgnoreCase))
            {
                gridBooks.SelectedIndex = i;
                gridBooks.ScrollIntoView(gridBooks.SelectedItem);
                return;
            }
        }
    }

    private void SetStatus(string message)
    {
        txtStatus.Text = message;
    }

    private static void SplitPublisherAndYear(string publishRaw, out string publisher, out string year)
    {
        publisher = publishRaw.Trim();
        year = string.Empty;
        if (string.IsNullOrWhiteSpace(publisher))
        {
            return;
        }

        var open = publisher.LastIndexOf('(');
        var close = publisher.LastIndexOf(')');
        if (open >= 0 && close > open)
        {
            var candidate = publisher[(open + 1)..close].Trim();
            if (candidate.All(char.IsDigit) && candidate.Length is 2 or 4)
            {
                year = candidate;
                publisher = publisher[..open].Trim();
            }
        }
    }
}
