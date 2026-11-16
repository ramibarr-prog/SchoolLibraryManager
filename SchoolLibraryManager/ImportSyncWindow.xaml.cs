using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;
using MiniExcelLibs;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace SchoolLibraryManager;

public partial class ImportSyncWindow : Window
{
    private readonly DatabaseService _dbService;
    private readonly List<DatabaseService.StudentSyncRow> _studentSyncRows = [];
    private readonly List<string> _paymentIds = [];
    private DataTable _brokenLoansTable = new();
    private readonly Dictionary<string, string> _classNameToCode = new(StringComparer.OrdinalIgnoreCase);

    public ImportSyncWindow(DatabaseService dbService)
    {
        InitializeComponent();
        _dbService = dbService;
        AddMaintenanceHint();
        Loaded += ImportSyncWindow_OnLoaded;
        ChkMarkMissingAsGraduates.Checked += ChkMarkMissingAsGraduates_OnChanged;
        ChkMarkMissingAsGraduates.Unchecked += ChkMarkMissingAsGraduates_OnChanged;
    }

    private void AddMaintenanceHint()
    {
        if (RestoreDatabasePanel.Parent is not System.Windows.Controls.Panel container)
        {
            return;
        }

        var hint = new System.Windows.Controls.TextBlock
        {
            Text = "מומלץ לבצע גיבוי/דחיסה קלה לפני ייבוא רשומות מרובות",
            Foreground = System.Windows.Media.Brushes.DarkOrange,
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 6, 0, 0),
            TextWrapping = TextWrapping.Wrap,
            FlowDirection = FlowDirection.RightToLeft
        };
        container.Children.Insert(0, hint);
    }

    private void ChkMarkMissingAsGraduates_OnChanged(object sender, RoutedEventArgs e)
    {
        if (_studentSyncRows.Count == 0)
        {
            return;
        }

        try
        {
            var preview = _dbService.SyncStudents(
                _studentSyncRows,
                ChkMarkMissingAsGraduates.IsChecked == true,
                dryRun: true);
            TxtSyncGraduates.Text = preview.MarkedGraduates.ToString("N0");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Refresh graduates preview failed");
        }
    }

    private void ImportSyncWindow_OnLoaded(object sender, RoutedEventArgs e)
    {
        try
        {
            LoadClassMappings();
            LoadDatabaseHealthCheck();
            SetStatus("חלון ייבוא וסנכרון מוכן.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Import sync window load failed");
            SetStatus("טעינת החלון נכשלה.");
        }
    }

    private void LoadClassMappings()
    {
        _classNameToCode.Clear();
        var classes = _dbService.GetClasses();
        foreach (DataRow row in classes.Rows)
        {
            var classId = row["ClassID"]?.ToString()?.Trim() ?? string.Empty;
            var className = row["ClassName"]?.ToString()?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(classId) && !string.IsNullOrWhiteSpace(className))
            {
                _classNameToCode[className] = classId;
            }
        }
    }

    private void BtnBrowseStudentSyncFile_OnClick(object sender, RoutedEventArgs e)
    {
        var path = PickImportFile();
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        TxtStudentSyncFile.Text = path;
    }

    private void BtnAnalyzeStudentSync_OnClick(object sender, RoutedEventArgs e)
    {
        if (!File.Exists(TxtStudentSyncFile.Text))
        {
            SetStatus("יש לבחור קובץ חניכים תקין.");
            return;
        }

        try
        {
            _studentSyncRows.Clear();
            var rows = ReadStudentRosterRows(TxtStudentSyncFile.Text);
            var parsedRows = ParseStudentSyncRows(rows).ToList();
            _studentSyncRows.AddRange(parsedRows);

            GridStudentSyncPreview.ItemsSource = _studentSyncRows;

            var preview = _dbService.SyncStudents(
                _studentSyncRows,
                ChkMarkMissingAsGraduates.IsChecked == true,
                dryRun: true);

            TxtSyncTotalStudents.Text = preview.TotalInFile.ToString("N0");
            TxtSyncNewStudents.Text = preview.NewStudents.ToString("N0");
            TxtSyncUpdatedStudents.Text = preview.UpdatedStudents.ToString("N0");
            TxtSyncGraduates.Text = preview.MarkedGraduates.ToString("N0");

            SetStatus($"נותחו {_studentSyncRows.Count} חניכים מהקובץ.");
        }
        catch (Exception ex)
        {
            if (ex is InvalidOperationException)
            {
                MessageBox.Show(ex.Message, "שגיאת ייבוא", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            ErrorLogger.Log(ex, "Analyze student sync file failed");
            SetStatus("ניתוח קובץ חניכים נכשל.");
        }
    }

    private void BtnExecuteStudentSync_OnClick(object sender, RoutedEventArgs e)
    {
        if (_studentSyncRows.Count == 0)
        {
            SetStatus("יש לנתח קובץ חניכים לפני ביצוע סנכרון.");
            return;
        }

        if (ChkMarkMissingAsGraduates.IsChecked == true)
        {
            DatabaseService.StudentSyncResult archivePreview;
            try
            {
                archivePreview = _dbService.SyncStudents(_studentSyncRows, true, dryRun: true);
            }
            catch (Exception ex)
            {
                ErrorLogger.Log(ex, "Archive candidates preview failed");
                SetStatus("בדיקת המועמדים לארכיון נכשלה. הסנכרון בוטל.");
                return;
            }

            var archivePercent = archivePreview.ActiveStudentsBeforeSync > 0
                ? archivePreview.MarkedGraduates * 100.0 / archivePreview.ActiveStudentsBeforeSync
                : 100.0;
            var overThreshold = archivePercent > 10.0;

            var archiveMessage =
                $"בסה\"כ {archivePreview.MarkedGraduates} חניכים לא מופיעים בקובץ ויועברו לארכיון.{Environment.NewLine}" +
                $"מתוכם {archivePreview.GraduatesWithActiveLoans} עדיין מחזיקים ספרים פעילים בספרייה.{Environment.NewLine}" +
                $"({archivePercent:0.#}% מהמצבת הפעילה){Environment.NewLine}{Environment.NewLine}";
            if (overThreshold)
            {
                archiveMessage += $"אזהרה: הפעולה מעבירה יותר מ-10% מהמצבת לארכיון. בדוק שקובץ האקסל שלם ותקין.{Environment.NewLine}{Environment.NewLine}";
            }

            archiveMessage += "האם להמשיך בייבוא?";

            var archiveConfirm = MessageBox.Show(
                archiveMessage,
                "אישור העברה לארכיון",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);
            if (archiveConfirm != MessageBoxResult.Yes)
            {
                SetStatus("הסנכרון בוטל על ידי המשתמש.");
                return;
            }

            if (overThreshold)
            {
                var strongConfirm = MessageBox.Show(
                    $"אישור סופי: להעביר {archivePreview.MarkedGraduates} חניכים ({archivePercent:0.#}% מהמצבת) לארכיון?",
                    "אישור מיוחד",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                    MessageBoxResult.No);
                if (strongConfirm != MessageBoxResult.Yes)
                {
                    SetStatus("הסנכרון בוטל על ידי המשתמש.");
                    return;
                }
            }
        }

        if (string.IsNullOrEmpty(_dbService.CreateBackup("PreSync")))
        {
            MessageBox.Show("יצירת הגיבוי לפני הסנכרון נכשלה, ולכן הפעולה בוטלה. לא בוצע שינוי במסד הנתונים.", "שגיאת גיבוי", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("הפעולה בוטלה: הגיבוי נכשל.");
            return;
        }

        var guardButton = sender as System.Windows.Controls.Button;
        if (guardButton is not null)
        {
            guardButton.IsEnabled = false;
        }

        try
        {
            var result = _dbService.SyncStudents(_studentSyncRows, ChkMarkMissingAsGraduates.IsChecked == true);
            SetStatus($"סנכרון הושלם: נוספו {result.NewStudents}, עודכנו {result.UpdatedStudents}, סומנו כבוגרים {result.MarkedGraduates}.");
            BtnAnalyzeStudentSync_OnClick(sender, e);
            MessageBox.Show(
                $"ייבוא מצבת חניכים הסתיים בהצלחה!{Environment.NewLine}נוספו: {result.NewStudents}{Environment.NewLine}עודכנו: {result.UpdatedStudents}{Environment.NewLine}הועברו לבוגרים: {result.MarkedGraduates}",
                "הצלחה",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Execute student sync failed");
            SetStatus("ביצוע סנכרון חניכים נכשל.");
            MessageBox.Show($"סנכרון החניכים נכשל:{Environment.NewLine}{ex.Message}", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (guardButton is not null)
            {
                guardButton.IsEnabled = true;
            }
        }
    }

    private void BtnBrowsePaymentFile_OnClick(object sender, RoutedEventArgs e)
    {
        var path = PickImportFile();
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        TxtPaymentImportFile.Text = path;
    }

    private void BtnPreviewPaymentImport_OnClick(object sender, RoutedEventArgs e)
    {
        if (!File.Exists(TxtPaymentImportFile.Text))
        {
            SetStatus("יש לבחור קובץ משלמים תקין.");
            return;
        }

        try
        {
            _paymentIds.Clear();
            var rows = ReadPaymentFileRows(TxtPaymentImportFile.Text);
            _paymentIds.AddRange(ExtractStudentIds(rows));
            _paymentStudentDetails.Clear();
            foreach (var pair in ExtractStudentDetails(rows))
            {
                _paymentStudentDetails[pair.Key] = pair.Value;
            }

            var preview = _dbService.PreviewPaymentImport(_paymentIds);
            TxtPayReadyToMark.Text = preview.ReadyToMarkPaid.ToString("N0");
            TxtPayAlreadyPaid.Text = preview.AlreadyPaidCount.ToString("N0");
            TxtPayExempt.Text = preview.ExemptCount.ToString("N0");
            TxtPayUnknown.Text = preview.UnknownStudentsCount.ToString("N0");

            GridUnknownPaymentIds.ItemsSource = preview.UnknownIds
                .Select(id => new
                {
                    UnknownId = _paymentStudentDetails.TryGetValue(id, out var d) && !string.IsNullOrWhiteSpace(d.FirstName + d.LastName)
                        ? $"{id} - {d.FirstName} {d.LastName}".Trim()
                        : id
                })
                .ToList();

            SetStatus($"בוצעה בדיקת קובץ משלמים: {preview.TotalInFile} תעודות זהות ייחודיות.");
        }
        catch (Exception ex)
        {
            if (ex is InvalidOperationException)
            {
                MessageBox.Show(ex.Message, "שגיאת ייבוא", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            ErrorLogger.Log(ex, "Preview payment import failed");
            SetStatus("בדיקת קובץ תשלומים נכשלה.");
        }
    }

    private static string GetExpectedSchoolYear(DateTime date)
    {
        var reference = date.Month >= 9 ? new DateTime(date.Year, 12, 31) : date;
        var hebrewYear = new System.Globalization.HebrewCalendar().GetYear(reference) % 1000;

        var letters = new System.Text.StringBuilder();
        var values = new[] { 400, 300, 200, 100, 90, 80, 70, 60, 50, 40, 30, 20, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1 };
        var chars = new[] { 'ת', 'ש', 'ר', 'ק', 'צ', 'פ', 'ע', 'ס', 'נ', 'מ', 'ל', 'כ', 'י', 'ט', 'ח', 'ז', 'ו', 'ה', 'ד', 'ג', 'ב', 'א' };
        var remaining = hebrewYear;
        for (var i = 0; i < values.Length && remaining > 0; i++)
        {
            if (remaining == 15) { letters.Append("טו"); break; }
            if (remaining == 16) { letters.Append("טז"); break; }
            while (remaining >= values[i])
            {
                letters.Append(chars[i]);
                remaining -= values[i];
            }
        }

        return DatabaseService.NormalizeSchoolYear(letters.ToString());
    }

    private void BtnExecutePaymentImport_OnClick(object sender, RoutedEventArgs e)
    {
        string currentYear;
        try
        {
            currentYear = _dbService.GetCurrentSchoolYear();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Payment import school year check");
            MessageBox.Show("שנת הלימודים הנוכחית אינה מוגדרת במערכת. לא בוצע ייבוא.", "שגיאת ייבוא", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var expectedYear = GetExpectedSchoolYear(DateTime.Now);
        if (!string.Equals(expectedYear, DatabaseService.NormalizeSchoolYear(currentYear), StringComparison.Ordinal))
        {
            var warning = MessageBox.Show(
                $"שים לב! התאריך של היום מצביע על כך שאנחנו בשנת הלימודים {expectedYear}, אך המערכת עדיין מוגדרת על שנת {currentYear}.{Environment.NewLine}האם אתה בטוח שברצונך לייבא את התשלומים לתוך שנת הלימודים הישנה?{Environment.NewLine}(מומלץ לבטל, לגשת ללשונית 'תקינות מסד' ולפתוח שנה חדשה).",
                "אזהרת שנת לימודים",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (warning != MessageBoxResult.Yes)
            {
                return;
            }
        }

        BtnExecutePaymentImportCore(sender, e);
    }

    private void BtnExecutePaymentImportCore(object sender, RoutedEventArgs e)
    {
        if (_paymentIds.Count == 0)
        {
            SetStatus("יש לבצע בדיקה לקובץ תשלומים לפני האישור.");
            return;
        }

        if (string.IsNullOrEmpty(_dbService.CreateBackup("PreSync")))
        {
            MessageBox.Show("יצירת הגיבוי לפני הייבוא נכשלה, ולכן הפעולה בוטלה. לא בוצע שינוי במסד הנתונים.", "שגיאת גיבוי", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("הפעולה בוטלה: הגיבוי נכשל.");
            return;
        }

        var guardButton = sender as System.Windows.Controls.Button;
        if (guardButton is not null)
        {
            guardButton.IsEnabled = false;
        }

        try
        {
            var result = _dbService.ExecutePaymentImport(_paymentIds, _paymentStudentDetails);
            SetStatus($"ייבוא תשלומים הושלם: סומנו {result.NewlyMarkedPaid} כשילמו, דולגו {result.Skipped}.");
            BtnPreviewPaymentImport_OnClick(sender, e);
            MessageBox.Show(
                $"ייבוא קובץ משלמים הסתיים בהצלחה!{Environment.NewLine}עודכנו כשילמו: {result.NewlyMarkedPaid}{Environment.NewLine}דולגו (כולל הגנת פטורים): {result.Skipped}",
                "הצלחה",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Execute payment import failed");
            SetStatus("אישור ייבוא תשלומים נכשל.");
            MessageBox.Show($"ייבוא התשלומים נכשל:{Environment.NewLine}{ex.Message}", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (guardButton is not null)
            {
                guardButton.IsEnabled = true;
            }
        }
    }

    private static string? PickImportFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "בחר קובץ ייבוא",
            Filter = "Excel/CSV (*.xlsx;*.csv)|*.xlsx;*.csv|Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv",
            CheckFileExists = true
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    private static List<Dictionary<string, string>> ReadFileRows(string path)
    {
        var extension = Path.GetExtension(path).ToLowerInvariant();
        if (extension is not (".csv" or ".xlsx"))
        {
            throw new InvalidOperationException("נתמכים רק קבצי xlsx ו-csv. יש לשמור קובץ xls בפורמט xlsx.");
        }

        return extension == ".csv" ? ReadCsvRows(path) : ReadExcelRows(path);
    }

    private static List<Dictionary<string, string>> ReadExcelRows(string path)
    {
        var result = new List<Dictionary<string, string>>();

        foreach (var row in MiniExcel.Query(path, useHeaderRow: true))
        {
            if (row is not IDictionary<string, object> dict)
            {
                continue;
            }

            var mapped = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in dict)
            {
                mapped[CleanHeader(pair.Key)] = pair.Value?.ToString()?.Trim() ?? string.Empty;
            }

            result.Add(mapped);
        }

        return result;
    }

    private static System.Text.Encoding DetectCsvEncoding(string path)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var bytes = File.ReadAllBytes(path);
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return new System.Text.UTF8Encoding(true);
        }

        if (bytes.Length >= 2 && ((bytes[0] == 0xFF && bytes[1] == 0xFE) || (bytes[0] == 0xFE && bytes[1] == 0xFF)))
        {
            return bytes[0] == 0xFF ? System.Text.Encoding.Unicode : System.Text.Encoding.BigEndianUnicode;
        }

        try
        {
            _ = new System.Text.UTF8Encoding(false, true).GetString(bytes);
            return new System.Text.UTF8Encoding(false);
        }
        catch (System.Text.DecoderFallbackException)
        {
            return System.Text.Encoding.GetEncoding(1255);
        }
    }

    private static List<Dictionary<string, string>> ReadCsvRows(string path)
    {
        var result = new List<Dictionary<string, string>>();

        using var parser = new TextFieldParser(path, DetectCsvEncoding(path));
        parser.TextFieldType = FieldType.Delimited;
        parser.SetDelimiters(",");
        parser.HasFieldsEnclosedInQuotes = true;

        if (parser.EndOfData)
        {
            return result;
        }

        var headers = parser.ReadFields() ?? [];
        while (!parser.EndOfData)
        {
            var fields = parser.ReadFields() ?? [];
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (var i = 0; i < headers.Length; i++)
            {
                var key = headers[i] is null ? $"Column{i + 1}" : CleanHeader(headers[i]);
                var value = i < fields.Length ? fields[i]?.Trim() ?? string.Empty : string.Empty;
                row[key] = value;
            }

            result.Add(row);
        }

        return result;
    }

    private IEnumerable<DatabaseService.StudentSyncRow> ParseStudentSyncRows(List<Dictionary<string, string>> rows)
    {
        if (rows.Count == 0)
        {
            yield break;
        }

        var headers = rows[0].Keys.ToList();
        var idColumn = RosterIdHeaders.FirstOrDefault(h => headers.Contains(h, StringComparer.OrdinalIgnoreCase));
        if (idColumn is null)
        {
            ValidateRequiredHeaders(rows, [IdHeader]);
        }

        var hasSeparateNames = headers.Contains("שם פרטי", StringComparer.OrdinalIgnoreCase)
            && headers.Contains("שם משפחה", StringComparer.OrdinalIgnoreCase);
        if (!hasSeparateNames)
        {
            ValidateRequiredHeaders(rows, [NameHeader]);
        }

        foreach (var row in rows)
        {
            var stuId = NormalizeStudentId(GetValue(row, idColumn));
            if (string.IsNullOrWhiteSpace(stuId))
            {
                continue;
            }

            string firstName;
            string lastName;
            if (hasSeparateNames)
            {
                firstName = GetValue(row, "שם פרטי");
                lastName = GetValue(row, "שם משפחה");
            }
            else
            {
                var parts = GetValue(row, NameHeader).Split(' ', StringSplitOptions.RemoveEmptyEntries);
                lastName = parts.Length > 0 ? parts[0] : string.Empty;
                firstName = parts.Length > 1 ? string.Join(' ', parts.Skip(1)) : string.Empty;
            }

            var gradeText = StripQuotes(GetValue(row, ClassHeader));
            if (string.IsNullOrWhiteSpace(gradeText))
            {
                gradeText = StripQuotes(GetValue(row, SheetNameKey)).Split(' ', StringSplitOptions.RemoveEmptyEntries).LastOrDefault() ?? string.Empty;
            }

            var parallelText = GetValue(row, "מקבילה");
            var internalName = GetValue(row, "שם כיתה");

            string classCode;
            string className;
            if (GradeLetterToNumber.TryGetValue(gradeText, out var grade) && int.TryParse(parallelText.Split('.')[0], out var parallel) && parallel > 0)
            {
                classCode = ClassFormatter.BuildClassCode(grade, parallel);
                var label = $"{gradeText}{parallel}";
                className = string.IsNullOrWhiteSpace(internalName) ? label : $"{internalName} ({label})";
            }
            else if (headers.Contains(NameHeader, StringComparer.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(gradeText) && !GradeLetterToNumber.ContainsKey(gradeText))
            {
                classCode = NormalizeClassCode(GetValue(row, ClassHeader));
                className = string.Empty;
            }
            else
            {
                classCode = "0000";
                className = "אחר";
            }

            yield return new DatabaseService.StudentSyncRow(stuId, firstName, lastName, classCode, string.Empty, string.Empty)
            {
                ClassName = className
            };
        }
    }

    private const string SheetNameKey = "__sheet";
    private static readonly string[] RosterIdHeaders = ["ת.ז", "ת.ז.", "ת.ז. תלמיד", "תעודת זהות"];
    private static readonly Dictionary<string, int> GradeLetterToNumber = new(StringComparer.Ordinal)
    {
        ["ז"] = 7, ["ח"] = 8, ["ט"] = 9, ["י"] = 10, ["יא"] = 11, ["יב"] = 12
    };

    private static string StripQuotes(string value)
    {
        return (value ?? string.Empty).Replace("'", string.Empty).Replace("\"", string.Empty).Replace("\u05F3", string.Empty).Replace("\u05F4", string.Empty).Trim();
    }

    private static List<Dictionary<string, string>> ReadStudentRosterRows(string path)
    {
        if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return ReadFileRows(path);
        }

        var result = new List<Dictionary<string, string>>();
        foreach (var sheet in MiniExcel.GetSheetNames(path))
        {
            var raw = new List<Dictionary<string, string>>();
            foreach (var row in MiniExcel.Query(path, useHeaderRow: false, sheetName: sheet))
            {
                if (row is not IDictionary<string, object> dict)
                {
                    continue;
                }

                raw.Add(dict.ToDictionary(p => p.Key, p => p.Value?.ToString()?.Trim() ?? string.Empty));
            }

            var headerIndex = -1;
            for (var i = 0; i < Math.Min(raw.Count, 20); i++)
            {
                if (raw[i].Values.Select(CleanHeader).Any(h => RosterIdHeaders.Contains(h, StringComparer.OrdinalIgnoreCase)))
                {
                    headerIndex = i;
                    break;
                }
            }

            if (headerIndex < 0)
            {
                continue;
            }

            var columnNames = raw[headerIndex].ToDictionary(p => p.Key, p => CleanHeader(p.Value));
            for (var i = headerIndex + 1; i < raw.Count; i++)
            {
                var mapped = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var pair in columnNames)
                {
                    if (!string.IsNullOrWhiteSpace(pair.Value))
                    {
                        mapped[pair.Value] = raw[i].TryGetValue(pair.Key, out var value) ? value : string.Empty;
                    }
                }

                var idColumn = RosterIdHeaders.FirstOrDefault(h => mapped.ContainsKey(h));
                if (idColumn is null || string.IsNullOrWhiteSpace(NormalizeStudentId(mapped[idColumn])))
                {
                    continue;
                }

                mapped[SheetNameKey] = sheet;
                result.Add(mapped);
            }
        }

        if (result.Count == 0)
        {
            throw new InvalidOperationException("הייבוא נעצר. לא נמצאו שורות חניכים עם ת.ז תקינה בקובץ.");
        }

        return result;
    }

    private const string IdHeader = "ת.ז. תלמיד";
    private const string NameHeader = "שם תלמיד";
    private const string ClassHeader = "כיתה";
    private const string PayStatusHeader = "סטטוס תשלום";
    private const string PayIdHeaderAlt = "תעודת זהות";
    private readonly Dictionary<string, (string FirstName, string LastName, string City, string Phone)> _paymentStudentDetails = new(StringComparer.OrdinalIgnoreCase);

    private static Dictionary<string, (string FirstName, string LastName, string City, string Phone)> ExtractStudentDetails(List<Dictionary<string, string>> rows)
    {
        var details = new Dictionary<string, (string FirstName, string LastName, string City, string Phone)>(StringComparer.OrdinalIgnoreCase);
        if (rows.Count == 0)
        {
            return details;
        }

        var idHeader = FindPaymentIdHeader(rows[0].Keys);
        foreach (var row in rows)
        {
            if (IsTotalRow(row))
            {
                continue;
            }

            var id = NormalizeStudentId(GetValue(row, idHeader));
            if (string.IsNullOrWhiteSpace(id))
            {
                continue;
            }

            details.TryAdd(id, (GetValue(row, "שם פרטי"), GetValue(row, "שם משפחה"), GetValue(row, "יישוב"), GetValue(row, "טלפון")));
        }

        return details;
    }
    private const string PayBalanceHeader = "יתרה לתלמיד";
    private const string PayTotalPaidHeader = "סה\"כ שולם";

    private static string FindPaymentIdHeader(IEnumerable<string> headers)
    {
        var list = headers.ToList();
        foreach (var candidate in new[] { PayIdHeaderAlt, IdHeader, "ת.ז" })
        {
            var match = list.FirstOrDefault(h => string.Equals(h, candidate, StringComparison.OrdinalIgnoreCase));
            if (match is not null)
            {
                return match;
            }
        }

        return PayIdHeaderAlt;
    }

    private static bool IsTotalRow(Dictionary<string, string> row)
    {
        return row.Values.Any(v => v.Contains("סה\"כ", StringComparison.Ordinal) || v.Contains("סה״כ", StringComparison.Ordinal));
    }

    private static decimal ParseAmount(string value)
    {
        var cleaned = new string((value ?? string.Empty).Where(c => char.IsDigit(c) || c is '.' or '-').ToArray());
        return decimal.TryParse(cleaned, System.Globalization.NumberStyles.Number | System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out var amount) ? amount : 0m;
    }

    private static List<Dictionary<string, string>> ReadPaymentFileRows(string path)
    {
        var extension = Path.GetExtension(path).ToLowerInvariant();
        if (extension is not (".csv" or ".xlsx"))
        {
            throw new InvalidOperationException("נתמכים רק קבצי xlsx ו-csv. יש לשמור קובץ xls בפורמט xlsx.");
        }

        var raw = new List<Dictionary<string, string>>();
        if (extension == ".csv")
        {
            using var parser = new TextFieldParser(path, DetectCsvEncoding(path));
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            parser.HasFieldsEnclosedInQuotes = true;
            while (!parser.EndOfData)
            {
                var fields = parser.ReadFields() ?? [];
                var line = new Dictionary<string, string>();
                for (var i = 0; i < fields.Length; i++)
                {
                    line[i.ToString()] = fields[i]?.Trim() ?? string.Empty;
                }

                raw.Add(line);
            }
        }
        else
        {
            foreach (var row in MiniExcel.Query(path, useHeaderRow: false))
            {
                if (row is not IDictionary<string, object> dict)
                {
                    continue;
                }

                raw.Add(dict.ToDictionary(p => p.Key, p => p.Value?.ToString()?.Trim() ?? string.Empty));
            }
        }

        var headerIndex = -1;
        for (var i = 0; i < Math.Min(raw.Count, 20); i++)
        {
            if (raw[i].Values.Select(CleanHeader).Any(h => string.Equals(h, PayIdHeaderAlt, StringComparison.OrdinalIgnoreCase)
                                                          || string.Equals(h, IdHeader, StringComparison.OrdinalIgnoreCase)))
            {
                headerIndex = i;
                break;
            }
        }

        if (headerIndex < 0)
        {
            throw new InvalidOperationException("הייבוא נעצר. לא נמצאה שורת כותרות עם העמודה 'תעודת זהות' בקובץ.");
        }

        var columnNames = raw[headerIndex].ToDictionary(p => p.Key, p => CleanHeader(p.Value));
        var result = new List<Dictionary<string, string>>();
        for (var i = headerIndex + 1; i < raw.Count; i++)
        {
            var mapped = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in columnNames)
            {
                if (string.IsNullOrWhiteSpace(pair.Value))
                {
                    continue;
                }

                mapped[pair.Value] = raw[i].TryGetValue(pair.Key, out var value) ? value : string.Empty;
            }

            result.Add(mapped);
        }

        if (result.Count == 0)
        {
            result.Add(new Dictionary<string, string>(columnNames.Values.Where(v => v.Length > 0).Distinct().ToDictionary(v => v, _ => string.Empty), StringComparer.OrdinalIgnoreCase));
        }

        return result;
    }

    private static string CleanHeader(string? header)
    {
        if (string.IsNullOrEmpty(header))
        {
            return string.Empty;
        }

        var chars = header
            .Where(c => char.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.Format)
            .Select(c => char.IsWhiteSpace(c) ? ' ' : c)
            .ToArray();
        return new string(chars).Trim();
    }

    private static void ValidateRequiredHeaders(List<Dictionary<string, string>> rows, string[] required)
    {
        var headers = rows.Count > 0 ? rows[0].Keys.ToList() : [];
        var missing = required.Where(h => !headers.Contains(h, StringComparer.OrdinalIgnoreCase)).ToList();
        if (missing.Count > 0)
        {
            throw new InvalidOperationException("הייבוא נעצר. חסרה כותרת חובה בקובץ: " + string.Join(", ", missing.Select(h => $"\"{h}\""))
                + Environment.NewLine + "עמודות שנקראו מהקובץ: " + string.Join(" | ", headers.Select(h => $"[{h}]")));
        }
    }

    private static List<string> ExtractStudentIds(List<Dictionary<string, string>> rows)
    {
        if (rows.Count == 0)
        {
            return [];
        }

        var headerKeys = rows[0].Keys;
        var idHeader = FindPaymentIdHeader(headerKeys);
        var hasTotalPaid = headerKeys.Contains(PayTotalPaidHeader, StringComparer.OrdinalIgnoreCase);
        var hasStatus = headerKeys.Contains(PayStatusHeader, StringComparer.OrdinalIgnoreCase);
        var hasBalance = headerKeys.Contains(PayBalanceHeader, StringComparer.OrdinalIgnoreCase);

        if (hasTotalPaid)
        {
            ValidateRequiredHeaders(rows, [idHeader, PayTotalPaidHeader]);
            return rows
                .Where(r => !IsTotalRow(r) && ParseAmount(GetValue(r, PayTotalPaidHeader)) > 0)
                .Select(r => NormalizeStudentId(GetValue(r, idHeader)))
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        if (hasStatus)
        {
            ValidateRequiredHeaders(rows, [idHeader, PayStatusHeader]);
        }
        else if (hasBalance)
        {
            ValidateRequiredHeaders(rows, [idHeader, PayBalanceHeader]);
        }
        else
        {
            ValidateRequiredHeaders(rows, [idHeader, PayStatusHeader + " / " + PayBalanceHeader]);
        }

        return rows
            .Where(r => !IsTotalRow(r) && (hasStatus
                ? !GetValue(r, PayStatusHeader).Contains("לא")
                : IsBalanceSettled(GetValue(r, PayBalanceHeader))))
            .Select(r => NormalizeStudentId(GetValue(r, idHeader)))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsBalanceSettled(string value)
    {
        var cleaned = new string(value.Where(c => char.IsDigit(c) || c is '.' or ',' or '-').ToArray()).Replace(",", string.Empty);
        if (cleaned.Length == 0)
        {
            return false;
        }

        return decimal.TryParse(cleaned, System.Globalization.NumberStyles.Number | System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out var balance) && balance <= 0;
    }

    private string NormalizeClassCode(string value)
    {
        var raw = (value ?? string.Empty).Replace("'", "").Replace("\"", "").Replace("\u05F3", "").Replace("\u05F4", "").Trim();
        if (string.IsNullOrWhiteSpace(raw))
        {
            return "0000";
        }

        return raw;
    }

    private static string NormalizeStudentId(string value)
    {
        var raw = (value ?? string.Empty).Trim();
        if (raw.EndsWith(".0", StringComparison.Ordinal))
        {
            raw = raw[..^2];
        }

        var digits = new string(raw.Where(char.IsDigit).ToArray());
        if (string.IsNullOrWhiteSpace(digits))
        {
            return string.Empty;
        }

        if (digits.Length > 9)
        {
            digits = digits[^9..];
        }

        return digits.PadLeft(9, '0');
    }

    private static string GetValue(Dictionary<string, string> row, string? key)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            return string.Empty;
        }

        return row.TryGetValue(key, out var value) ? value?.Trim() ?? string.Empty : string.Empty;
    }

    private static string? FindColumn(IEnumerable<string> headers, IEnumerable<string> candidates)
    {
        var normalizedHeaders = headers
            .Select(h => new { Original = h, Normalized = NormalizeHeader(h) })
            .ToList();

        foreach (var candidate in candidates)
        {
            var normalizedCandidate = NormalizeHeader(candidate);
            var match = normalizedHeaders.FirstOrDefault(h => h.Normalized.Contains(normalizedCandidate));
            if (match != null)
            {
                return match.Original;
            }
        }

        return null;
    }

    private static string NormalizeHeader(string value)
    {
        return new string((value ?? string.Empty)
            .Where(ch => !char.IsWhiteSpace(ch) && ch != '_' && ch != '-' && ch != '.' && ch != '"' && ch != '\'')
            .ToArray())
            .ToLowerInvariant();
    }

    private static void SetMaintenanceBusy(object sender, bool busy)
    {
        if (sender is FrameworkElement { Parent: UIElement panel })
        {
            panel.IsEnabled = !busy;
        }

        Mouse.OverrideCursor = busy ? Cursors.Wait : null;
    }

    private async Task RunLoanIntegrityMaintenanceAsync(bool applyFix, object sender)
    {
        SetMaintenanceBusy(sender, true);
        SetStatus("מבצע פעולה, אנא המתן...");
        try
        {
            string message = string.Empty;
            var report = await Task.Run(() => _dbService.RepairOrphanActiveLoans(applyFix, out message));
            var summary = applyFix
                ? $"תיקון הושלם | נסרקו: {report.ActiveLoansScanned:N0} | חריגות: {report.OrphanLoansFound:N0} | נסגרו: {report.LoansClosed:N0} | עותקים שאופסו: {report.CopiesReset:N0}"
                : $"סריקה הושלמה | נסרקו: {report.ActiveLoansScanned:N0} | חריגות שנמצאו: {report.OrphanLoansFound:N0}";

            SetStatus($"{summary} | {message}");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "RunLoanIntegrityMaintenance failed");
            SetStatus("סריקת/תיקון שלמות נתונים נכשלה.");
        }
        finally
        {
            SetMaintenanceBusy(sender, false);
        }
    }

    private async void BtnCompactDatabase_OnClick(object sender, RoutedEventArgs e)
    {
        var confirmation = MessageBox.Show(
            "האם לבצע דחיסה ומיטוב למסד הנתונים כעת?",
            "דחיסה ומיטוב מסד נתונים",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question,
            MessageBoxResult.No);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        SetMaintenanceBusy(sender, true);
        (Application.Current.MainWindow as MainWindow)?.SetMaintenanceMode(true);
        SetStatus("מבצע דחיסה ומיטוב, אנא המתן...");
        try
        {
            var result = await Task.Run(() => _dbService.CompactAndRepairDatabase());
            SetStatus(result.Message);
            MessageBox.Show(result.Message,
                "דחיסה ומיטוב מסד נתונים",
                MessageBoxButton.OK,
                result.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Compact database action failed");
            SetStatus("דחיסה ומיטוב מסד הנתונים נכשלו.");
            MessageBox.Show("דחיסה ומיטוב מסד הנתונים נכשלו.", "דחיסה ומיטוב מסד נתונים", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            (Application.Current.MainWindow as MainWindow)?.SetMaintenanceMode(false);
            SetMaintenanceBusy(sender, false);
        }
    }

    private void BtnRolloverSchoolYear_OnClick(object sender, RoutedEventArgs e)
    {
        string currentYear;
        try
        {
            currentYear = _dbService.GetCurrentSchoolYear();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Rollover: read current school year");
            currentYear = "(לא מוגדרת)";
        }

        var inputDialog = new InputDialog(
            $"הכנס את שנת התשלומים החדשה (לדוגמה: תשפח){Environment.NewLine}השנה הנוכחית: {currentYear}{Environment.NewLine}{Environment.NewLine}שים לב: הפעולה רלוונטית לתשלומים בלבד ואינה משנה נתונים אחרים.",
            "פתח שנת תשלומים חדשה");

        if (IsLoaded && IsVisible)
        {
            inputDialog.Owner = this;
        }
        else
        {
            inputDialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        if (inputDialog.ShowDialog() != true)
        {
            return;
        }

        var input = inputDialog.InputText;

        if (string.IsNullOrWhiteSpace(input))
        {
            return;
        }

        var newYear = DatabaseService.NormalizeSchoolYear(input);
        if (string.Equals(newYear, DatabaseService.NormalizeSchoolYear(currentYear), StringComparison.Ordinal))
        {
            MessageBox.Show("שנת הלימודים החדשה זהה לשנה הנוכחית. לא בוצע שינוי.", "פתח שנת לימודים חדשה", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var confirmation = MessageBox.Show(
            $"לעבור משנת תשלומים {currentYear} לשנת תשלומים {newYear}?\n\nהפעולה משפיעה על התשלומים בלבד: כל החניכים יופיעו כ\"טרם שילמו\" עד לייבוא תשלומי השנה החדשה, ופטורים יש לסמן מחדש.\nחניכים, ספרים, השאלות ושאר הנתונים לא ישתנו.\nמומלץ לבצע את הפעולה רק לפני ייבוא קובץ התשלומים של השנה החדשה.\nלפני השינוי ייווצר גיבוי אוטומטי.",
            "אישור מעבר שנת תשלומים",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        if (string.IsNullOrEmpty(_dbService.CreateBackup("PreRollover")))
        {
            const string backupFailed = "הגיבוי לפני המעבר נכשל, ולכן שנת הלימודים לא שונתה.";
            SetStatus(backupFailed);
            MessageBox.Show(backupFailed, "פתח שנת לימודים חדשה", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var success = _dbService.RolloverToNewSchoolYear(newYear, out var message);
        SetStatus(message);
        MessageBox.Show(message, "פתח שנת לימודים חדשה", MessageBoxButton.OK, success ? MessageBoxImage.Information : MessageBoxImage.Error);
    }

    private void SetStatus(string message)
    {
        TxtStatus.Text = message;
    }

    private sealed class BackupFileItem
    {
        public string FullPath { get; init; } = string.Empty;
        public string FileName { get; init; } = string.Empty;
        public DateTime LastWriteTime { get; init; }
        public string DateText { get; init; } = string.Empty;
        public string TypeText { get; init; } = string.Empty;
        public string SizeText { get; init; } = string.Empty;
        public string Kind { get; init; } = string.Empty;
    }

    private MessageBoxResult ShowRestoreMessage(string text, string caption, MessageBoxButton button, MessageBoxImage image, MessageBoxResult defaultResult = MessageBoxResult.OK)
    {
        const MessageBoxOptions rtl = MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading;
        Window? owner = Window.GetWindow(this) ?? Application.Current?.MainWindow;
        if (owner != null && owner.IsLoaded)
        {
            return MessageBox.Show(owner, text, caption, button, image, defaultResult, rtl);
        }

        return MessageBox.Show(text, caption, button, image, defaultResult, rtl);
    }

    private static DateTime GetBackupTimestamp(FileInfo info)
    {
        var name = System.IO.Path.GetFileNameWithoutExtension(info.Name);
        if (name.Length >= 15 && DateTime.TryParseExact(name[^15..], "yyyyMMdd_HHmmss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var parsed))
        {
            return parsed;
        }

        return info.CreationTime;
    }

    private static string GetBackupKind(string fileName)
    {
        if (fileName.Contains("_Emergency_PreRestore_", StringComparison.OrdinalIgnoreCase)) return "Emergency_PreRestore";
        if (fileName.Contains("_PreCompact_", StringComparison.OrdinalIgnoreCase)) return "PreCompact";
        if (fileName.Contains("_Shutdown_", StringComparison.OrdinalIgnoreCase)) return "Shutdown";
        if (fileName.Contains("_Manual_", StringComparison.OrdinalIgnoreCase)) return "Manual";
        if (fileName.Contains("_Auto_", StringComparison.OrdinalIgnoreCase)) return "Auto";
        return string.Empty;
    }

    private static string GetBackupKindText(string kind) => kind switch
    {
        "Auto" => "🟢 גיבוי אוטומטי (הפעלה)",
        "Manual" => "🟣 גיבוי ידני יזום",
        "Shutdown" => "🔵 גיבוי סגירת מערכת",
        "PreCompact" => "🟠 גיבוי לפני דחיסה",
        "Emergency_PreRestore" => "🔴 גיבוי חירום (לפני שחזור)",
        _ => "⚪ קובץ גיבוי"
    };

    private static string GetRelativeAge(DateTime date)
    {
        var days = (DateTime.Today - date.Date).Days;
        return days <= 0 ? "היום" : days == 1 ? "אתמול" : $"לפני {days} ימים";
    }

    private static string FormatFileSize(long bytes)
        => bytes >= 1024 * 1024
            ? $"{bytes / (1024d * 1024d):0.00} MB"
            : $"{bytes / 1024d:0.00} KB";

    private void LoadBackupsList()
    {
        try
        {
            var filterIndex = CmbBackupFilter.SelectedIndex;
            var items = new List<BackupFileItem>();
            if (Directory.Exists(_dbService.BackupFolder))
            {
                foreach (var file in Directory.EnumerateFiles(_dbService.BackupFolder))
                {
                    var ext = System.IO.Path.GetExtension(file);
                    if (!ext.Equals(".mdb", StringComparison.OrdinalIgnoreCase) && !ext.Equals(".accdb", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var info = new FileInfo(file);
                    var kind = GetBackupKind(info.Name);
                    var backupTime = GetBackupTimestamp(info);
                    var include = filterIndex switch
                    {
                        1 => kind == "Manual",
                        2 => kind is "Auto" or "Shutdown",
                        3 => kind is "PreCompact" or "Emergency_PreRestore",
                        _ => true
                    };
                    if (!include)
                    {
                        continue;
                    }

                    items.Add(new BackupFileItem
                    {
                        FullPath = info.FullName,
                        FileName = info.Name,
                        LastWriteTime = backupTime,
                        DateText = $"{backupTime:dd/MM/yyyy HH:mm:ss} ({GetRelativeAge(backupTime)})",
                        TypeText = GetBackupKindText(kind),
                        SizeText = FormatFileSize(info.Length),
                        Kind = kind
                    });
                }
            }

            DgAvailableBackups.ItemsSource = items.OrderByDescending(x => x.LastWriteTime).ToList();
            BtnConfirmRestore.IsEnabled = false;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "LoadBackupsList failed");
            SetStatus("טעינת רשימת הגיבויים נכשלה.");
        }
    }

    private void BtnRestoreDatabase_OnClick(object sender, RoutedEventArgs e)
    {
        if (RestoreDatabasePanel.Visibility == Visibility.Visible)
        {
            RestoreDatabasePanel.Visibility = Visibility.Collapsed;
            return;
        }

        RestoreDatabasePanel.Visibility = Visibility.Visible;
        LoadBackupsList();
    }

    private void CmbBackupFilter_OnSelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (IsLoaded && RestoreDatabasePanel.Visibility == Visibility.Visible)
        {
            LoadBackupsList();
        }
    }

    private void BtnRefreshBackups_OnClick(object sender, RoutedEventArgs e) => LoadBackupsList();

    private void DgAvailableBackups_OnSelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        BtnConfirmRestore.IsEnabled = DgAvailableBackups.SelectedItem is BackupFileItem;
    }

    private void BtnCancelRestore_OnClick(object sender, RoutedEventArgs e)
    {
        RestoreDatabasePanel.Visibility = Visibility.Collapsed;
    }

    private async void BtnBrowseExternalBackup_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "בחירת קובץ גיבוי",
            Filter = "Access database (*.mdb;*.accdb)|*.mdb;*.accdb",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var info = new FileInfo(dialog.FileName);
        await RunRestoreAsync(info.FullName, GetBackupKindText(GetBackupKind(info.Name)), GetBackupTimestamp(info));
    }

    private async void BtnConfirmRestore_OnClick(object sender, RoutedEventArgs e)
    {
        if (DgAvailableBackups.SelectedItem is not BackupFileItem item)
        {
            return;
        }

        await RunRestoreAsync(item.FullPath, item.TypeText, item.LastWriteTime);
    }

    private async Task RunRestoreAsync(string path, string typeText, DateTime backupDate)
    {
        const string caption = "שחזור מסד נתונים";
        const string failedStatus = "שחזור מסד הנתונים נכשל.";
        var restarting = false;

        try
        {
            var info = new FileInfo(path);
            if (!info.Exists || info.Length <= 0)
            {
                ShowRestoreMessage("קובץ הגיבוי לא נמצא או שהוא ריק.", caption, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var confirm = ShowRestoreMessage(
                $"האם לשחזר את מסד הנתונים מהגיבוי הבא?\n\nתאריך: {backupDate:dd/MM/yyyy HH:mm:ss}\nסוג: {typeText}\nקובץ: {info.Name}\n\nהנתונים הנוכחיים יוחלפו. גיבוי חירום ייווצר אוטומטית לפני השחזור.",
                caption, MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            App.IsRestoring = true;
            (Application.Current.MainWindow as MainWindow)?.SetMaintenanceMode(true);
            RestoreDatabasePanel.IsEnabled = false;
            BtnRestoreDatabase.IsEnabled = false;
            Mouse.OverrideCursor = Cursors.Wait;
            SetStatus("משחזר מסד נתונים, אנא המתן...");

            var result = await Task.Run(() => _dbService.RestoreDatabase(path));

            Mouse.OverrideCursor = null;

            if (!result.Success)
            {
                ShowRestoreMessage(result.Message, caption, MessageBoxButton.OK, MessageBoxImage.Error);
                SetStatus(failedStatus);
                return;
            }

            ShowRestoreMessage("השחזור הושלם בהצלחה. המערכת תופעל מחדש כעת.", caption, MessageBoxButton.OK, MessageBoxImage.Information);
            RestoreDatabasePanel.Visibility = Visibility.Collapsed;

            restarting = true;
            var relaunched = false;
            try
            {
                var exePath = Environment.ProcessPath;
                if (!string.IsNullOrWhiteSpace(exePath))
                {
                    var startInfo = new System.Diagnostics.ProcessStartInfo(exePath) { UseShellExecute = true };
                    foreach (var arg in Environment.GetCommandLineArgs().Skip(1).Where(a => !string.Equals(a, "--restarted", StringComparison.OrdinalIgnoreCase)))
                    {
                        startInfo.ArgumentList.Add(arg);
                    }

                    startInfo.ArgumentList.Add("--restarted");
                    System.Diagnostics.Process.Start(startInfo);
                    relaunched = true;
                }
            }
            catch (Exception relaunchEx)
            {
                ErrorLogger.Log(relaunchEx, "Relaunch after restore failed");
            }

            if (!relaunched)
            {
                ShowRestoreMessage("השחזור הושלם בהצלחה, אך לא ניתן היה להפעיל את התוכנה מחדש אוטומטית. התוכנה תיסגר, יש להפעיל אותה ידנית.", caption, MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            Application.Current.Shutdown();
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Restore database from UI failed");
            Mouse.OverrideCursor = null;
            ShowRestoreMessage($"אירעה שגיאה בעת שחזור מסד הנתונים: {ex.Message}", caption, MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus(failedStatus);
        }
        finally
        {
            Mouse.OverrideCursor = null;
            if (!restarting)
            {
                App.IsRestoring = false;
                (Application.Current.MainWindow as MainWindow)?.SetMaintenanceMode(false);
                RestoreDatabasePanel.IsEnabled = true;
                BtnRestoreDatabase.IsEnabled = true;
            }
        }
    }

    private void BtnBackupNow_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var backupPath = _dbService.CreateBackup("Manual");
            if (string.IsNullOrEmpty(backupPath))
            {
                MessageBox.Show("יצירת הגיבוי נכשלה. פרטים ביומן השגיאות.", "גיבוי מסד נתונים", MessageBoxButton.OK, MessageBoxImage.Error);
                SetStatus("גיבוי מסד הנתונים נכשל.");
                return;
            }

            var pathMessage = $"הגיבוי נשמר בהצלחה:\n{System.IO.Path.GetFileName(backupPath)}";

            MessageBox.Show(pathMessage, "גיבוי מסד נתונים", MessageBoxButton.OK, MessageBoxImage.Information);
            SetStatus("גיבוי מסד הנתונים בוצע בהצלחה.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Manual backup from reports failed");
            MessageBox.Show("אירעה שגיאה בעת ביצוע גיבוי מסד הנתונים.", "גיבוי מסד נתונים", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("גיבוי מסד הנתונים נכשל.");
        }
    }

    private async void BtnRunHealthCheck_OnClick(object sender, RoutedEventArgs e)
    {
        BtnRunHealthCheck.IsEnabled = false;
        BtnBackupNow.IsEnabled = false;
        Mouse.OverrideCursor = Cursors.Wait;
        SetStatus("מבצע בדיקת שלמות מסד, אנא המתן...");

        try
        {
            var report = await Task.Run(() => _dbService.RunDatabaseHealthCheck());
            AnomalousLoansSection.Visibility = Visibility.Visible;
            LoadDatabaseHealthCheck(report);
            SetStatus(_brokenLoansTable.Rows.Count == 0
                ? "בדיקת שלמות המסד הושלמה: מסד הנתונים תקין."
                : $"בדיקת שלמות המסד הושלמה: נמצאו {_brokenLoansTable.Rows.Count} השאלות חריגות הדורשות טיפול.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Run database health check failed");
            SetStatus("הרצת בדיקת שלמות נכשלה: " + ex.Message);
        }
        finally
        {
            Mouse.OverrideCursor = null;
            BtnRunHealthCheck.IsEnabled = true;
            BtnBackupNow.IsEnabled = true;
        }
    }

    private void LoadDatabaseHealthCheck(DatabaseService.DbHealthReport? existingReport = null)
    {
        var report = existingReport ?? _dbService.RunDatabaseHealthCheck();

        TxtHealthTotalBooks.Text = report.TotalBooks.ToString("N0");
        TxtHealthTotalCopies.Text = report.TotalCopies.ToString("N0");
        TxtHealthTotalStudents.Text = report.TotalStudents.ToString("N0");
        TxtHealthActiveLoans.Text = report.ActiveTransactions.ToString("N0");
        TxtHealthHistoryRecords.Text = report.HistoryTransactions.ToString("N0");

        var brokenForAlerts = _dbService.GetBrokenActiveLoans();
        var scrappedCount = brokenForAlerts.Select("AutoClean = 1 AND MissingBarcode = 0").Length;
        var alertTexts = new List<string>(report.Warnings);
        if (scrappedCount > 0)
        {
            alertTexts.Add($"נמצאו {scrappedCount} השאלות פעילות של עותקים גרוטים (שאריות סלקוד).");
        }

        if (alertTexts.Count == 0 && brokenForAlerts.Rows.Count > 0)
        {
            alertTexts.Add($"נמצאו {brokenForAlerts.Rows.Count} השאלות חריגות הדורשות טיפול.");
        }

        var warningItems = new List<HealthMessageItem>();
        if (alertTexts.Count == 0)
        {
            warningItems.Add(new HealthMessageItem(
                "מסד הנתונים תקין לחלוטין וללא שגיאות שלמות!",
                new SolidColorBrush((Color)ColorConverter.ConvertFromString("#16A34A"))));
        }
        else
        {
            foreach (var warning in alertTexts)
            {
                var color = warning.Contains("כפולים", StringComparison.OrdinalIgnoreCase) ||
                            warning.Contains("אינו קיים", StringComparison.OrdinalIgnoreCase)
                    ? "#DC2626"
                    : "#C2410C";

                warningItems.Add(new HealthMessageItem(
                    warning,
                    new SolidColorBrush((Color)ColorConverter.ConvertFromString(color))));
            }
        }

        LstHealthWarnings.ItemsSource = warningItems;

        LoadBrokenActiveLoans();
    }

    private void LoadBrokenActiveLoans()
    {
        _brokenLoansTable = _dbService.GetBrokenActiveLoans();
        var hasRows = _brokenLoansTable.Rows.Count > 0;

        GridBrokenLoans.ItemsSource = _brokenLoansTable.DefaultView;
        GridBrokenLoans.Visibility = hasRows ? Visibility.Visible : Visibility.Collapsed;
        TxtBrokenLoansOk.Visibility = hasRows ? Visibility.Collapsed : Visibility.Visible;
        BtnAutoCleanMissingBarcodeLoans.Visibility = AnomalousLoansSection.Visibility == Visibility.Visible && _brokenLoansTable.Select("AutoClean = 1").Length > 0 ? Visibility.Visible : Visibility.Collapsed;
        BtnAutoCleanMissingBarcodeLoans.IsEnabled = true;
    }

    private void BtnResolveBrokenLoan_OnClick(object sender, RoutedEventArgs e)
    {
        var row = (sender as System.Windows.Controls.Button)?.CommandParameter as DataRowView
                  ?? (sender as FrameworkElement)?.DataContext as DataRowView;
        if (row is null)
        {
            SetStatus("לא נבחרה רשומת השאלה לטיפול.");
            return;
        }

        var barcode = row.Row.Table.Columns.Contains("Barcode") ? row["Barcode"]?.ToString()?.Trim() ?? string.Empty : string.Empty;
        if (string.IsNullOrWhiteSpace(barcode))
        {
            SetStatus("לא ניתן לזהות ברקוד להשאלה החריגה.");
            return;
        }

        var studentName = row.Row.Table.Columns.Contains("StudentName") ? row["StudentName"]?.ToString()?.Trim() ?? string.Empty : string.Empty;

        var confirm = ShowRestoreMessage(
            $"האם לנקות / לסגור את ההשאלה החריגה עבור ברקוד {barcode}?",
            "אישור ניקוי השאלה",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
        GridBrokenLoans.IsEnabled = false;
        BtnAutoCleanMissingBarcodeLoans.IsEnabled = false;
        SetStatus("מעבד נתונים ומנקה השאלות מול מסד הנתונים...");
        try
        {
            var success = _dbService.CleanOrCloseOrphanedLoan(barcode, out var message);
            SetStatus(message);
            if (!success)
            {
                ShowRestoreMessage(message, "ניקוי השאלה", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            LoadDatabaseHealthCheck();
            var who = string.IsNullOrWhiteSpace(studentName) ? string.Empty : $" - {studentName}";
            var done = $"ההשאלה החריגה (ברקוד {barcode}{who}) נסגרה ונוקתה ממסד הנתונים בהצלחה.";
            SetStatus(done);
            ShowRestoreMessage(done, "ניקוי השאלה", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"Clean orphaned loan failed for {barcode}");
            SetStatus("ניקוי ההשאלה נכשל.");
        }
        finally
        {
            System.Windows.Input.Mouse.OverrideCursor = null;
            GridBrokenLoans.IsEnabled = true;
            BtnAutoCleanMissingBarcodeLoans.IsEnabled = true;
        }
    }

    private async void BtnAutoCleanMissingBarcodeLoans_OnClick(object sender, RoutedEventArgs e)
    {
        var confirm = ShowRestoreMessage(
        "לנקות אוטומטית את כל ההשאלות החריגות (ברקוד שאינו קיים במלאי ועותקים גרוטים רשומים כמושאלים)? השאלות של חניך שאינו קיים יישארו לטיפול ידני.",
            "ניקוי אוטומטי",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
        BtnAutoCleanMissingBarcodeLoans.IsEnabled = false;
        GridBrokenLoans.IsEnabled = false;
        SetStatus("מעבד נתונים ומנקה השאלות מול מסד הנתונים...");
        try
        {
            int cleaned = 0;
            string message = string.Empty;
            var ok = await Task.Run(() => _dbService.CleanAllAnomalousLoans(out cleaned, out message));
            LoadDatabaseHealthCheck(await Task.Run(() => _dbService.RunDatabaseHealthCheck()));
            SetStatus(message);
            ShowRestoreMessage(ok ? $"נוקו ונסגרו {cleaned} השאלות חריגות (ברקוד חסר / עותק גרוט) בהצלחה." : message,
                "ניקוי אוטומטי", MessageBoxButton.OK, ok ? MessageBoxImage.Information : MessageBoxImage.Error);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Auto clean missing barcode loans failed");
            SetStatus("הניקוי האוטומטי נכשל.");
        }
        finally
        {
            System.Windows.Input.Mouse.OverrideCursor = null;
            GridBrokenLoans.IsEnabled = true;
            BtnAutoCleanMissingBarcodeLoans.IsEnabled = true;
        }
    }

    private sealed record HealthMessageItem(string Message, Brush Foreground);
}
