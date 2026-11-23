using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace SchoolLibraryManager;

public class DatabaseService
{
    private readonly string _databasePath;
    private readonly string _backupFolder;
    private readonly string _connectionString;
    private static readonly Dictionary<string, string> studentNames = new(StringComparer.OrdinalIgnoreCase);

    public readonly record struct InventoryOverview(
        int TotalTitles,
        int TotalCopies,
        int ActiveLoans,
        int AvailableCopies);

    public readonly record struct DatabaseDiagnostics(string DatabasePath, string StudentTableName, string StudentIdColumnName, int StudentCount);
    public readonly record struct DbHealthReport(
        int TotalStudents,
        int TotalBooks,
        int TotalCopies,
        int ActiveLoans,
        int OverdueLoans,
        int BrokenLoanLinks,
        int MissingBooks,
        int MissingStudents)
    {
        public int ActiveTransactions => ActiveLoans;
        public int HistoryTransactions => OverdueLoans;
        public List<string> Warnings { get; init; } = [];
        public List<string> SampleBarcodes { get; init; } = [];

        public int ExtraMetric1 { get; init; }
        public int ExtraMetric2 { get; init; }
        public int ExtraMetric3 { get; init; }
        public int ExtraMetric4 { get; init; }

        public DbHealthReport(
            int totalStudents,
            int totalBooks,
            int totalCopies,
            int activeLoans,
            int overdueLoans,
            int brokenLoanLinks,
            int missingBooks,
            int missingStudents,
            int extraMetric1,
            int extraMetric2,
            int extraMetric3,
            int extraMetric4)
            : this(totalStudents, totalBooks, totalCopies, activeLoans, overdueLoans, brokenLoanLinks, missingBooks, missingStudents)
        {
            ExtraMetric1 = extraMetric1;
            ExtraMetric2 = extraMetric2;
            ExtraMetric3 = extraMetric3;
            ExtraMetric4 = extraMetric4;
        }
    }

    public readonly record struct LoanMaintenanceReport(
        int ActiveLoansScanned,
        int OrphanLoansFound,
        int LoansClosed,
        int CopiesReset,
        List<string> SampleBarcodes);

    public bool UpdateBook(string bkId, string bkName, string bkAuthor, string subjectName, string publisher, string year, string price, out string message)
    {
        var publishCombined = (publisher ?? string.Empty).Trim();
        var yearValue = (year ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(yearValue))
        {
            publishCombined = string.IsNullOrWhiteSpace(publishCombined)
                ? yearValue
                : $"{publishCombined} ({yearValue})";
        }

        var subjectId = ResolveSubjectIdByName(subjectName);
        return UpdateBook(bkId, bkName, bkAuthor, price, publishCombined, subjectId, 0, out message);
    }

    public sealed class StudentSyncRow
    {
        public StudentSyncRow()
        {
        }

        public StudentSyncRow(string stuId, string firstName, string lastName, string classCode, string? phone, string? city)
        {
            StuId = stuId ?? string.Empty;
            FirstName = firstName ?? string.Empty;
            LastName = lastName ?? string.Empty;
            ClassCode = classCode ?? string.Empty;
            Phone = phone ?? string.Empty;
            City = city ?? string.Empty;
        }

        public string StuId { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string ClassCode { get; set; } = string.Empty;
        public string ClassName { get; set; } = string.Empty;
        public string Phone { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
    }

    public readonly record struct StudentSyncResult(
        int TotalInFile,
        int NewStudents,
        int UpdatedStudents,
        int MarkedGraduates,
        int GraduatesWithActiveLoans = 0,
        int ActiveStudentsBeforeSync = 0);

    public DatabaseService(string? databasePath = null)
    {
        var baseDirectory = AppContext.BaseDirectory;
        var dataDirectory = Path.Combine(baseDirectory, "Data");

        _databasePath = ResolveDatabasePath(databasePath, baseDirectory, dataDirectory);
        _backupFolder = Path.Combine(baseDirectory, "Backups");

        const string provider = "Provider=Microsoft.ACE.OLEDB.12.0;";
        _connectionString = $"{provider}Data Source=\"{_databasePath}\";Persist Security Info=False;";
    }

    public sealed record AnnualLossSummary(
            string SchoolYear,
            DateTime From,
            DateTime ToExclusive,
            int UnreturnedCount,
            decimal TotalCost,
            int BooksWithoutPrice,
            DataTable Details);

        /// <summary>
        /// Annual loss report: active loans (History = '0') whose loan date falls inside the current school year,
        /// with the estimated replacement cost taken from Book.bkPrice.
        /// </summary>
        public AnnualLossSummary GetAnnualLossSummary(string schoolYear)
        {
            schoolYear = (schoolYear ?? string.Empty).Trim();

            var (from, toExclusive) = GetSchoolYearDateRange(schoolYear);

            const string sql =
                "SELECT t.Barcode AS Barcode, b.bkName AS BookTitle, b.bkPrice AS PriceText, " +
                "t.DateOfTrans AS LoanDate, t.DateToReturn AS DueDate, t.stuID AS StuID, " +
                "s.stuFname AS FirstName, s.stuLname AS LastName " +
                "FROM (([Transaction] AS t " +
                "LEFT JOIN [BookCopy] AS c ON t.Barcode = c.Barcode) " +
                "LEFT JOIN [Book] AS b ON t.bkID = b.bkID) " +
                "LEFT JOIN [Student] AS s ON t.stuID = s.stuID " +
                "WHERE t.History = '0' AND t.DateOfTrans >= ? AND t.DateOfTrans < ? " +
                "ORDER BY t.DateOfTrans";

            var raw = new DataTable();
            using (var connection = CreateFreshPaymentReadConnection())
            using (var cmd = new OleDbCommand(sql, connection))
            {
                cmd.Parameters.Add("@from", OleDbType.Date).Value = from;
                cmd.Parameters.Add("@to", OleDbType.Date).Value = toExclusive;
                using var adapter = new OleDbDataAdapter(cmd);
                adapter.Fill(raw);
            }

            var details = new DataTable();
            details.Columns.Add("Barcode", typeof(string));
            details.Columns.Add("BookTitle", typeof(string));
            details.Columns.Add("StudentName", typeof(string));
            details.Columns.Add("LoanDate", typeof(string));
            details.Columns.Add("DueDate", typeof(string));
            details.Columns.Add("Price", typeof(decimal));

            decimal total = 0m;
            var withoutPrice = 0;
            foreach (DataRow row in raw.Rows)
            {
                var hasPrice = TryParseLossPrice(row["PriceText"], out var price);
                if (hasPrice)
                {
                    total += price;
                }
                else
                {
                    withoutPrice++;
                }

                var name = $"{row["FirstName"]} {row["LastName"]}".Trim();
                details.Rows.Add(
                    Convert.ToString(row["Barcode"], CultureInfo.InvariantCulture)?.Trim() ?? string.Empty,
                    Convert.ToString(row["BookTitle"], CultureInfo.InvariantCulture)?.Trim() ?? string.Empty,
                    name,
                    FormatLossDate(row["LoanDate"]),
                    FormatLossDate(row["DueDate"]),
                    hasPrice ? price : 0m);
            }

            return new AnnualLossSummary(schoolYear, from, toExclusive, details.Rows.Count, total, withoutPrice, details);
        }

        private static string FormatLossDate(object? value)
            => value is DateTime d ? d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : string.Empty;

        private static bool TryParseLossPrice(object? value, out decimal price)
        {
            price = 0m;
            var text = Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim();
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            var cleaned = new string(text.Where(ch => char.IsDigit(ch) || ch == '.' || ch == ',').ToArray());
            if (cleaned.Length == 0)
            {
                return false;
            }

            if (cleaned.Contains(',') && !cleaned.Contains('.'))
            {
                cleaned = cleaned.Replace(',', '.');
            }
            else
            {
                cleaned = cleaned.Replace(",", string.Empty);
            }

            return decimal.TryParse(cleaned, NumberStyles.Number, CultureInfo.InvariantCulture, out price) && price > 0m;
        }

        /// <summary>
        /// Maps a Hebrew school year (e.g. "תשפז") to a Gregorian range starting June 21st of the year the
        /// Hebrew year begins, up to and including June 20th of the next year (exclusive upper bound: June 21st).
        /// Falls back to the current accounting year (June 21st cut-off).
        /// </summary>
        private static (DateTime From, DateTime ToExclusive) GetSchoolYearDateRange(string schoolYear)
        {
            if (TryParseHebrewYear(schoolYear, out var hebrewYear))
            {
                var startYear = hebrewYear - 3761;
                return (new DateTime(startYear, 6, 21), new DateTime(startYear + 1, 6, 21));
            }

            var today = DateTime.Today;
            var cutOffThisYear = new DateTime(today.Year, 6, 21);
            var year = today >= cutOffThisYear ? today.Year : today.Year - 1;
            return (new DateTime(year, 6, 21), new DateTime(year + 1, 6, 21));
        }

        internal static bool TryParseHebrewYear(string? text, out int hebrewYear)
        {
            hebrewYear = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var sum = 0;
            foreach (var ch in text)
            {
                sum += ch switch
                {
                    'א' => 1, 'ב' => 2, 'ג' => 3, 'ד' => 4, 'ה' => 5, 'ו' => 6, 'ז' => 7, 'ח' => 8, 'ט' => 9,
                    'י' => 10, 'כ' or 'ך' => 20, 'ל' => 30, 'מ' or 'ם' => 40, 'נ' or 'ן' => 50, 'ס' => 60,
                    'ע' => 70, 'פ' or 'ף' => 80, 'צ' or 'ץ' => 90, 'ק' => 100, 'ר' => 200, 'ש' => 300, 'ת' => 400,
                    _ => 0
                };
            }

            var candidate = 5000 + sum;
            if (candidate is < 5700 or > 5900)
            {
                return false;
            }

            hebrewYear = candidate;
            return true;
        }

        /// <summary>Returns Hebrew school-year labels (e.g. "תשפז") around the given year, newest first.</summary>
        public static List<string> GetSchoolYearChoices(string? currentYear, int yearsBack = 8)
        {
            var center = TryParseHebrewYear(currentYear, out var parsed)
                ? parsed
                : (DateTime.Today.Month >= 6 ? DateTime.Today.Year : DateTime.Today.Year - 1) + 3761;

            const int firstSupportedYear = 5784; // תשפד
            var result = new List<string>();
            for (var year = center + 1; year >= Math.Max(center - yearsBack, firstSupportedYear); year--)
            {
                result.Add(ToHebrewYearLetters(year));
            }

            return result;
        }

        private static string ToHebrewYearLetters(int hebrewYear)
        {
            var n = hebrewYear % 1000;
            var sb = new System.Text.StringBuilder();
            string[] hundreds = ["", "ק", "ר", "ש", "ת", "תק", "תר", "תש", "תת", "תתק"];
            string[] tens = ["", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ"];
            string[] units = ["", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט"];

            sb.Append(hundreds[n / 100]);
            var rest = n % 100;
            if (rest == 15)
            {
                sb.Append("טו");
            }
            else if (rest == 16)
            {
                sb.Append("טז");
            }
            else
            {
                sb.Append(tens[rest / 10]);
                sb.Append(units[rest % 10]);
            }

            return sb.ToString();
        }

        public DataTable GetInventoryBySubject()
    {
        var books = new DataTable();
        var copies = new DataTable();

        using (var connection = CreateFreshPaymentReadConnection())
        {
            using (var cmd = new OleDbCommand("SELECT bk.bkID AS BookID, s.SubjectName AS SubjectName FROM [Book] AS bk LEFT JOIN [Subject] AS s ON bk.bkSubjectID = s.SubjectID", connection))
            using (var adapter = new OleDbDataAdapter(cmd))
            {
                adapter.Fill(books);
            }

            using (var cmd = new OleDbCommand("SELECT bkID AS BookID FROM [BookCopy] WHERE Trim(CopyCondition & '') IN ('1','2')", connection))
            using (var adapter = new OleDbDataAdapter(cmd))
            {
                adapter.Fill(copies);
            }
        }

        var copiesPerBook = copies.AsEnumerable()
            .GroupBy(r => Convert.ToString(r["BookID"], CultureInfo.InvariantCulture)?.Trim() ?? string.Empty)
            .ToDictionary(g => g.Key, g => g.Count());

        var result = new DataTable();
        result.Columns.Add("SubjectName", typeof(string));
        result.Columns.Add("TitlesCount", typeof(int));
        result.Columns.Add("CopiesCount", typeof(int));

        foreach (var group in books.AsEnumerable()
            .GroupBy(r => string.IsNullOrWhiteSpace(Convert.ToString(r["SubjectName"])) ? "ללא תחום" : Convert.ToString(r["SubjectName"])!.Trim())
            .OrderBy(g => g.Key))
        {
            var copyCount = group.Sum(r => copiesPerBook.TryGetValue(Convert.ToString(r["BookID"], CultureInfo.InvariantCulture)?.Trim() ?? string.Empty, out var c) ? c : 0);
            result.Rows.Add(group.Key, group.Count(), copyCount);
        }

        return result;
    }

    public InventoryOverview GetInventoryOverview()
    {
        using var connection = CreateFreshPaymentReadConnection();
        var totalTitles = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [Book]", connection).ExecuteScalar() ?? 0, CultureInfo.InvariantCulture);
        var totalCopies = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [BookCopy] WHERE Trim(CopyCondition & '') IN ('1','2')", connection).ExecuteScalar() ?? 0, CultureInfo.InvariantCulture);
        var activeLoans = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [Transaction] WHERE History = '0'", connection).ExecuteScalar() ?? 0, CultureInfo.InvariantCulture);
        var activeCopyLoans = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [Transaction] AS t INNER JOIN [BookCopy] AS bc ON t.Barcode = bc.Barcode WHERE t.History = '0' AND Trim(bc.CopyCondition & '') IN ('1','2')", connection).ExecuteScalar() ?? 0, CultureInfo.InvariantCulture);
        var availableCopies = Math.Max(0, totalCopies - activeCopyLoans);

        return new InventoryOverview(totalTitles, totalCopies, activeLoans, availableCopies);
    }

    public DataTable GetUnpaidStudentsReport(string? classFilter = null)
    {
        var all = SearchStudents(string.Empty, classFilter);
        var unpaid = all.Clone();
        foreach (DataRow row in all.Rows)
        {
            if (string.Equals(row["PaymentStatus"]?.ToString()?.Trim(), "לא שולם", StringComparison.Ordinal))
            {
                unpaid.ImportRow(row);
            }
        }

        return unpaid;
    }

    public DataTable GetClasses()
    {
        var result = new DataTable();

        using var connection = CreateFreshPaymentReadConnection();

        var hasClasses = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null)?
            .AsEnumerable()
            .Any(r => string.Equals(r["TABLE_NAME"]?.ToString()?.Trim(), "Classes", StringComparison.OrdinalIgnoreCase)) == true;

        if (!hasClasses)
        {
            result.Columns.Add("ClassID", typeof(string));
            result.Columns.Add("ClassName", typeof(string));
            return result;
        }

        using var cmd = new OleDbCommand("SELECT c.ClassID, c.ClassName FROM [Classes] AS c WHERE c.ClassID IN ('0000', '0001', '9999') OR c.ClassID IN (SELECT DISTINCT stuClassID FROM [Student] WHERE stuClassID IS NOT NULL) ORDER BY c.ClassName", connection);
        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(result);

        if (result.Rows.Count == 0)
        {
            using var fallbackCmd = new OleDbCommand("SELECT DISTINCT stuClassID AS ClassID, stuClassID AS ClassName FROM [Student] WHERE stuClassID IS NOT NULL AND stuClassID <> '9999' ORDER BY stuClassID", connection);
            using var fallbackAdapter = new OleDbDataAdapter(fallbackCmd);
            fallbackAdapter.Fill(result);
        }

        return result;
    }

    public bool IsStudentExemptForCurrentSchoolYear(string stuId)
    {
        var normalized = (stuId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return false;
        }

        if (!TryGetCurrentSchoolYear(out var currentYear))
        {
            return false;
        }

        using var connection = CreateFreshPaymentReadConnection();
        var idCandidates = BuildStudentIdCandidates(normalized);
        if (idCandidates.Length == 0)
        {
            return false;
        }

        var idFilter = string.Join(" OR ", idCandidates.Select(_ => "stuID = ?"));
        var sql = $@"SELECT Amount, ReciptNum, [Year]
                     FROM [Payment]
                     WHERE ({idFilter})";

        using var cmd = new OleDbCommand(sql, connection);
        foreach (var candidate in idCandidates)
        {
            cmd.Parameters.AddWithValue("@p", candidate);
        }

        using var reader = cmd.ExecuteReader();
        while (reader != null && reader.Read())
        {
            var yearValue = reader["Year"] == DBNull.Value ? string.Empty : reader["Year"]?.ToString()?.Trim() ?? string.Empty;
            if (!IsCurrentSchoolYearValue(yearValue, currentYear))
            {
                continue;
            }

            var amount = ReadAmountValue(reader["Amount"]);
            var reciptNum = reader["ReciptNum"] == DBNull.Value ? string.Empty : reader["ReciptNum"]?.ToString()?.Trim() ?? string.Empty;
            if (IsExemptionPaymentRow(amount, reciptNum))
            {
                return true;
            }
        }

        return false;
    }

    public bool CanDeleteStudent(string stuId)
    {
        return CanDeleteStudent(stuId, out _);
    }

    public bool CanDeleteStudent(string stuId, out string reason)
    {
        reason = string.Empty;
        var normalized = (stuId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            reason = "חניך לא תקין.";
            return false;
        }

        using var connection = CreateOpenConnection();
        using var cmd = new OleDbCommand("SELECT COUNT(*) FROM [Transaction] WHERE stuID = ? AND History = '0'", connection);
        cmd.Parameters.AddWithValue("@p1", normalized);
        var activeLoans = Convert.ToInt32(cmd.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture);
        if (activeLoans > 0)
        {
            reason = "לא ניתן למחוק חניך עם השאלות פעילות.";
            return false;
        }

        return true;
    }

    public bool TryDeleteStudent(string stuId, out string message)
    {
        var normalized = (stuId ?? string.Empty).Trim();
        if (!CanDeleteStudent(normalized, out var reason))
        {
            message = reason;
            return false;
        }

        using var connection = CreateOpenConnection();
        using var cmd = new OleDbCommand("UPDATE [Student] SET stuClassID = '9999' WHERE stuID = ?", connection);
        cmd.Parameters.AddWithValue("@p1", normalized);
        var affected = cmd.ExecuteNonQuery();
        message = affected > 0 ? "החניך סומן כלא פעיל בהצלחה." : "לא נמצא חניך למחיקה.";
        return affected > 0;
    }

    public bool SetStudentExemptionForCurrentSchoolYear(string stuId, bool shouldExempt)
    {
        return SetStudentExemptionForCurrentSchoolYear(stuId, shouldExempt, out _);
    }

    public bool SetStudentExemptionForCurrentSchoolYear(string stuId, bool shouldExempt, out string message)
    {
        var normalized = (stuId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            message = "חניך לא תקין.";
            return false;
        }

        if (!TryGetCurrentSchoolYear(out var currentYear))
        {
            message = "שנת הלימודים הנוכחית אינה מוגדרת או שלא ניתן לקרוא אותה. לא בוצע שינוי.";
            return false;
        }

        var idCandidates = BuildStudentIdCandidates(normalized);
        if (idCandidates.Length == 0)
        {
            message = "חניך לא תקין.";
            return false;
        }

        var shouldReleasePaymentPool = false;
        try
        {
            using var connection = CreateOpenConnection();

            if (shouldExempt)
            {
                var existsFilter = string.Join(" OR ", idCandidates.Select(_ => "stuID = ?"));
                var existsSql = $@"SELECT Amount, ReciptNum, [Year]
                                   FROM [Payment]
                                   WHERE ({existsFilter})";

                var hasExemptionRow = false;
                decimal realPaymentAmount = 0m;
                var realPaymentReceipt = string.Empty;
                using (var existsCmd = new OleDbCommand(existsSql, connection))
                {
                    foreach (var candidate in idCandidates)
                    {
                        existsCmd.Parameters.AddWithValue("@p", candidate);
                    }

                    using var existsReader = existsCmd.ExecuteReader();
                    while (existsReader != null && existsReader.Read())
                    {
                        var rowYear = existsReader["Year"] == DBNull.Value ? string.Empty : existsReader["Year"]?.ToString() ?? string.Empty;
                        if (!IsCurrentSchoolYearValue(rowYear, currentYear))
                        {
                            continue;
                        }

                        var rowAmount = ReadAmountValue(existsReader["Amount"]);
                        var rowReceipt = existsReader["ReciptNum"] == DBNull.Value ? string.Empty : existsReader["ReciptNum"]?.ToString()?.Trim() ?? string.Empty;
                        if (IsExemptionPaymentRow(rowAmount, rowReceipt))
                        {
                            hasExemptionRow = true;
                        }
                        else if (rowAmount > 0)
                        {
                            realPaymentAmount += rowAmount;
                            realPaymentReceipt = string.IsNullOrWhiteSpace(realPaymentReceipt) ? rowReceipt : realPaymentReceipt;
                        }
                    }
                }

                if (hasExemptionRow)
                {
                    message = "החניך כבר מסומן כפטור לשנה זו.";
                    return true;
                }

                if (realPaymentAmount > 0)
                {
                    message = $"לחניך קיים תשלום לשנה זו (סכום: {realPaymentAmount.ToString("0.##", CultureInfo.InvariantCulture)}, קבלה: {realPaymentReceipt}). לא ניתן לסמן פטור ללא מחיקת התשלום בנפרד.";
                    return false;
                }

                var basePayKey = $"P_{normalized}_{currentYear}";
                var payKeyId = basePayKey;
                var suffix = 0;
                while (true)
                {
                    using var keyCheckCmd = new OleDbCommand("SELECT COUNT(*) FROM [Payment] WHERE PayKeyID = ?", connection);
                    keyCheckCmd.Parameters.AddWithValue("@pPayKey", payKeyId);
                    var keyExists = Convert.ToInt32(keyCheckCmd.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture) > 0;
                    if (!keyExists)
                    {
                        break;
                    }

                    suffix++;
                    payKeyId = $"{basePayKey}_{suffix:D2}";
                }

                using var insertCmd = new OleDbCommand(
                    "INSERT INTO [Payment] (PayKeyID, stuID, Amount, ReciptNum, [Year]) VALUES (?, ?, ?, ?, ?)",
                    connection);
                insertCmd.Parameters.AddWithValue("@pPayKey", payKeyId);
                insertCmd.Parameters.AddWithValue("@pStuId", normalized);
                insertCmd.Parameters.AddWithValue("@pAmount", 0);
                insertCmd.Parameters.AddWithValue("@pReciptNum", "פטור מאושר");
                insertCmd.Parameters.AddWithValue("@pYear", currentYear);

                var inserted = insertCmd.ExecuteNonQuery();
                shouldReleasePaymentPool = inserted > 0;
                message = inserted > 0 ? "הפטור עודכן בהצלחה." : "הפטור לא עודכן.";
                return inserted > 0;
            }

            var deleteFilter = string.Join(" OR ", idCandidates.Select(_ => "stuID = ?"));
            var yearVariants = GetSchoolYearVariants(currentYear);
            var yearFilter = string.Join(", ", yearVariants.Select(_ => "?"));
            using var deleteCmd = new OleDbCommand(
                $"DELETE FROM [Payment] WHERE ({deleteFilter}) AND Amount = 0 AND ReciptNum = 'פטור מאושר' AND [Year] IN ({yearFilter})", connection);
            foreach (var candidate in idCandidates)
            {
                deleteCmd.Parameters.AddWithValue("@pStuId", candidate);
            }

            foreach (var variant in yearVariants)
            {
                deleteCmd.Parameters.AddWithValue("@pYear", variant);
            }
            var deleted = deleteCmd.ExecuteNonQuery();
            shouldReleasePaymentPool = deleted > 0;
            message = "הפטור בוטל בהצלחה.";
            return true;
        }
        catch (OleDbException ex)
        {
            ErrorLogger.Log(ex, $"SetStudentExemptionForCurrentSchoolYear failed (stuId={normalized}, year={currentYear}, shouldExempt={shouldExempt}): {ex}");
            message = $"עדכון פטור נכשל: {ex.Message}";
            return false;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"SetStudentExemptionForCurrentSchoolYear failed (stuId={normalized}, year={currentYear}, shouldExempt={shouldExempt}): {ex}");
            message = $"עדכון פטור נכשל: {ex.Message}";
            return false;
        }
        finally
        {
            if (shouldReleasePaymentPool)
            {
                OleDbConnection.ReleaseObjectPool();
            }
        }
    }

    public bool TryUpdateStudent(string stuId, string firstName, string lastName, string classId, string? phone, string? city, string? remarks, out string message)
    {
        var normalizedStuId = (stuId ?? string.Empty).Trim();
        var normalizedFirstName = (firstName ?? string.Empty).Trim();
        var normalizedLastName = (lastName ?? string.Empty).Trim();
        var normalizedClassId = NormalizeClassCode(classId);

        if (string.IsNullOrWhiteSpace(normalizedStuId))
        {
            message = "חניך לא תקין.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(normalizedFirstName) || string.IsNullOrWhiteSpace(normalizedLastName))
        {
            message = "שם פרטי ושם משפחה הם שדות חובה.";
            return false;
        }

        try
        {
            if (string.Equals(normalizedClassId, "9999", StringComparison.Ordinal)
                && !CanDeleteStudent(normalizedStuId))
            {
                message = "לא ניתן להעביר חניך לסטטוס לא פעיל בעודו מחזיק ספרים.";
                return false;
            }

            using var connection = CreateOpenConnection();
            using var cmd = new OleDbCommand(
                "UPDATE [Student] SET stuFname = ?, stuLname = ?, stuClassID = ?, stuPhone = ?, stuCity = ?, stuRemarks = ? WHERE stuID = ?",
                connection);
            cmd.Parameters.AddWithValue("@pFName", normalizedFirstName);
            cmd.Parameters.AddWithValue("@pLName", normalizedLastName);
            cmd.Parameters.AddWithValue("@pClass", normalizedClassId);
            cmd.Parameters.AddWithValue("@pPhone", (phone ?? string.Empty).Trim());
            cmd.Parameters.AddWithValue("@pCity", (city ?? string.Empty).Trim());
            cmd.Parameters.AddWithValue("@pRemarks", (remarks ?? string.Empty).Trim());
            cmd.Parameters.AddWithValue("@pStuId", normalizedStuId);

            var affected = cmd.ExecuteNonQuery();
            message = affected > 0 ? "פרטי החניך נשמרו בהצלחה." : "לא נמצא חניך לעדכון.";
            return affected > 0;
        }
        catch (OleDbException ex)
        {
            ErrorLogger.Log(ex, "TryUpdateStudent failed");
            message = $"שמירת פרטי החניך נכשלה: {ex.Message}";
            return false;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "TryUpdateStudent failed");
            message = $"שמירת פרטי החניך נכשלה: {ex.Message}";
            return false;
        }
    }

    private const string OtherClassCode = "0000";
    private const string OtherClassDisplay = "--אחר--";

    private static bool IsOtherClassValue(string? value)
    {
        var normalized = (value ?? string.Empty).Trim();
        return string.IsNullOrWhiteSpace(normalized)
            || string.Equals(normalized, "0", StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, OtherClassCode, StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, OtherClassDisplay, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeClassCode(string? classCode)
    {
        var normalized = (classCode ?? string.Empty).Trim();
        return IsOtherClassValue(normalized) ? OtherClassCode : normalized;
    }

    private static string ResolveDatabasePath(string? providedPath, string baseDirectory, string dataDirectory)
    {
        if (!string.IsNullOrWhiteSpace(providedPath))
        {
            var fullProvided = Path.GetFullPath(providedPath);
            if (File.Exists(fullProvided)) return fullProvided;
            throw new FileNotFoundException("DB_NOT_FOUND", fullProvided);
        }

        var candidates = new[]
        {
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LibraryDB.mdb"),
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "LibraryDB.mdb")
        };

        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate)) return candidate;
        }

        throw new FileNotFoundException("DB_NOT_FOUND");
    }

    private static readonly string[] RequiredTables = { "Student", "Transaction", "History", "BookCopy", "Book" };

    private static Dictionary<string, long>? TryGetTableRowCounts(string path, out string error)
    {
        error = string.Empty;
        try
        {
            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{Path.GetFullPath(path)}\";Mode=Read;Persist Security Info=False;";
            using var connection = new OleDbConnection(connectionString);
            connection.Open();
            var counts = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            foreach (var table in RequiredTables)
            {
                using var cmd = new OleDbCommand($"SELECT COUNT(*) FROM [{table}]", connection);
                counts[table] = Convert.ToInt64(cmd.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture);
            }

            return counts;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"TryGetTableRowCounts({path})");
            error = ex.Message;
            return null;
        }
    }

    public static bool IsBackupOrTempDatabaseFile(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path ?? string.Empty);
        string[] markers = { "_Auto_", "_Manual_", "_Shutdown_", "_PreCompact_", "_PreUpdate_", "_PreSync_", "_PreRollover_", "_compact_", "_precompact_", "_restore_" };
        if (markers.Any(m => name.Contains(m, StringComparison.OrdinalIgnoreCase))) return true;

        var directory = Path.GetFileName(Path.GetDirectoryName(Path.GetFullPath(path ?? string.Empty)) ?? string.Empty);
        return string.Equals(directory, "Backups", StringComparison.OrdinalIgnoreCase);
    }

    public static bool TryValidateDatabaseFile(string path, out string error)
    {
        error = string.Empty;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            error = "הקובץ לא קיים.";
            return false;
        }

        if (IsBackupOrTempDatabaseFile(path))
        {
            error = "הקובץ נראה כקובץ גיבוי או קובץ זמני. יש לבחור את קובץ המסד הראשי.";
            return false;
        }

        try
        {
            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{Path.GetFullPath(path)}\";Mode=Read;Persist Security Info=False;";
            using var connection = new OleDbConnection(connectionString);
            connection.Open();
            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object?[] { null, null, null, "TABLE" });
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (schema != null)
            {
                foreach (DataRow row in schema.Rows)
                {
                    names.Add(row["TABLE_NAME"]?.ToString() ?? string.Empty);
                }
            }

            var missing = RequiredTables.Where(t => !names.Contains(t)).ToList();
            if (missing.Count > 0)
            {
                error = "בקובץ חסרות הטבלאות: " + string.Join(", ", missing);
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"TryValidateDatabaseFile({path})");
            error = "לא ניתן לפתוח את הקובץ כמסד נתונים: " + ex.Message;
            return false;
        }
    }

    public string DatabasePath => _databasePath;

    public string GetSchoolName()
    {
        const string fallback = "ספריית בית הספר";
        try
        {
            using var connection = CreateOpenConnection();
            using var cmd = new OleDbCommand("SELECT TOP 1 SchoolName FROM [School]", connection);
            var value = cmd.ExecuteScalar()?.ToString()?.Trim();
            return string.IsNullOrWhiteSpace(value) ? fallback : value;
        }
        catch
        {
            return fallback;
        }
    }

    public OleDbConnection CreateOpenConnection()
    {
        var connection = new OleDbConnection(_connectionString);
        try
        {
            connection.Open();
            return connection;
        }
        catch (Exception ex)
        {
            connection.Dispose();
            TryNotifyDatabaseAccessError(ex);
            throw;
        }
    }

    private static DateTime _lastDbAccessErrorNotice = DateTime.MinValue;

    public static bool IsDatabaseAccessError(Exception? ex)
    {
        for (var current = ex; current is not null; current = current.InnerException)
        {
            if (current is OleDbException oleEx && oleEx.ErrorCode == unchecked((int)0x80004005))
            {
                return true;
            }
        }

        return false;
    }

    public static bool TryNotifyDatabaseAccessError(Exception? ex)
    {
        if (!IsDatabaseAccessError(ex))
        {
            return false;
        }

        try
        {
            ErrorLogger.Log(ex!, "Database access error (locked/read-only)");

            if ((DateTime.UtcNow - _lastDbAccessErrorNotice).TotalSeconds < 3)
            {
                return true;
            }

            _lastDbAccessErrorNotice = DateTime.UtcNow;
            const string text = "שגיאה בגישה למסד הנתונים. ייתכן שהקובץ מוגדר לקריאה בלבד, שהוא פתוח בתוכנה אחרת (כגון Access), או שהוא ננעל באופן זמני (למשל על ידי אנטי וירוס או סנכרון ענן).";
            var dispatcher = System.Windows.Application.Current?.Dispatcher;
            if (dispatcher is null || dispatcher.CheckAccess())
            {
                System.Windows.MessageBox.Show(text, "שגיאה", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            else
            {
                dispatcher.Invoke(() => System.Windows.MessageBox.Show(text, "שגיאה", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error));
            }
        }
        catch
        {
        }

        return true;
    }

    private OleDbConnection CreateFreshPaymentReadConnection()
    {
        var builder = new OleDbConnectionStringBuilder(_connectionString)
        {
            ["OLE DB Services"] = -4
        };

        var connection = new OleDbConnection(builder.ConnectionString);
        try
        {
            connection.Open();

            using var barrierCommand = new OleDbCommand("SELECT COUNT(*) FROM [Payment] WHERE 1=0", connection);
            barrierCommand.ExecuteScalar();

            return connection;
        }
        catch
        {
            connection.Dispose();
            throw;
        }
    }

    public void Initialize()
    {
        CreateBackup("Auto");
        EnsureSystemSettings();
    }

    public string CreateBackup(string backupPrefix)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_databasePath) || !File.Exists(_databasePath))
            {
                return string.Empty;
            }

            Directory.CreateDirectory(_backupFolder);

            var prefix = string.IsNullOrWhiteSpace(backupPrefix) ? "Manual" : backupPrefix.Trim();
            var name = Path.GetFileNameWithoutExtension(_databasePath);
            var ext = Path.GetExtension(_databasePath);
            var destination = Path.Combine(_backupFolder, $"{name}_{prefix}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");

            var partial = Path.Combine(_backupFolder, "~partial_" + Path.GetFileName(destination));

            OleDbConnection.ReleaseObjectPool();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            try
            {
                File.Copy(_databasePath, partial, true);

                var sourceCounts = TryGetTableRowCounts(_databasePath, out var sourceError);
                var backupCounts = TryGetTableRowCounts(partial, out var backupError);
                if (sourceCounts is null || backupCounts is null)
                {
                    return BackupFailed(backupPrefix, $"לא ניתן לאמת את קובץ הגיבוי. {sourceError} {backupError}".Trim());
                }

                foreach (var table in RequiredTables)
                {
                    if (!sourceCounts.TryGetValue(table, out var before) || !backupCounts.TryGetValue(table, out var after) || before != after)
                    {
                        return BackupFailed(backupPrefix, $"אימות הגיבוי נכשל (מספר השורות בטבלה {table} שונה).");
                    }
                }

                File.Move(partial, destination, true);
            }
            finally
            {
                try { if (File.Exists(partial)) File.Delete(partial); } catch { }
            }

            File.SetCreationTime(destination, DateTime.Now);
            File.SetLastWriteTime(destination, DateTime.Now);

            CleanupOldBackups(destination);
            ClearBackupFailureFlag();
            return destination;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"CreateBackup({backupPrefix}) failed");
            return BackupFailed(backupPrefix, ex.Message);
        }
    }

    private static string BackupFailureFlagPath => Path.Combine(AppContext.BaseDirectory, "BackupFailed.flag");

    private static string BackupFailed(string backupPrefix, string reason)
    {
        ErrorLogger.Log($"Backup failed ({backupPrefix}): {reason}", "CreateBackup");
        try
        {
            File.WriteAllText(BackupFailureFlagPath, $"{DateTime.Now:yyyy-MM-dd HH:mm} | {backupPrefix} | {reason}");
        }
        catch
        {
        }

        return string.Empty;
    }

    private static void ClearBackupFailureFlag()
    {
        try { if (File.Exists(BackupFailureFlagPath)) File.Delete(BackupFailureFlagPath); } catch { }
    }

    public static string? ConsumeBackupFailureWarning()
    {
        try
        {
            return File.Exists(BackupFailureFlagPath) ? File.ReadAllText(BackupFailureFlagPath) : null;
        }
        catch
        {
            return null;
        }
    }

    private static bool IsAutomaticBackupFile(string file)
    {
        var extension = Path.GetExtension(file);
        if (!extension.Equals(".mdb", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".accdb", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var fileName = Path.GetFileName(file);
        return fileName.Contains("_Auto_", StringComparison.OrdinalIgnoreCase)
            || fileName.Contains("_Shutdown_", StringComparison.OrdinalIgnoreCase)
            || fileName.Contains("_PreCompact_", StringComparison.OrdinalIgnoreCase)
            || fileName.Contains("_PreUpdate_", StringComparison.OrdinalIgnoreCase);
    }

    private static DateTime GetBackupTimestamp(string file)
    {
        var nameWithoutExtension = Path.GetFileNameWithoutExtension(file);
        if (nameWithoutExtension.Length >= 15
            && DateTime.TryParseExact(
                nameWithoutExtension[^15..],
                "yyyyMMdd_HHmmss",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out var parsed))
        {
            return parsed;
        }

        return File.GetLastWriteTime(file);
    }

    private void CleanupOldBackups(string keepPath)
    {
        try
        {
            const int keepFloor = 20;
            var now = DateTime.Now;

            var automaticBackups = Directory.EnumerateFiles(_backupFolder)
                .Where(IsAutomaticBackupFile)
                .Select(file => (File: file, Date: GetBackupTimestamp(file)))
                .OrderByDescending(item => item.Date)
                .ToList();

            var keptDays = new HashSet<DateTime>();
            var keptMonths = new HashSet<(int Year, int Month)>();

            foreach (var item in automaticBackups.Skip(keepFloor))
            {
                if (string.Equals(item.File, keepPath, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var age = now - item.Date;
                bool keep;
                if (age <= TimeSpan.FromHours(48))
                {
                    keep = true;
                }
                else if (age <= TimeSpan.FromDays(14))
                {
                    keep = keptDays.Add(item.Date.Date);
                }
                else if (age <= TimeSpan.FromDays(365))
                {
                    keep = keptMonths.Add((item.Date.Year, item.Date.Month));
                }
                else
                {
                    keep = false;
                }

                if (keep)
                {
                    continue;
                }

                try
                {
                    File.Delete(item.File);
                }
                catch (Exception ex)
                {
                    ErrorLogger.Log(ex, $"Backup cleanup failed for {item.File}");
                }
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "CleanupOldBackups failed");
        }
    }

    public string BackupFolder => _backupFolder;

    public (bool Success, string Message) RestoreDatabase(string backupPath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(backupPath) || !File.Exists(backupPath))
            {
                return (false, "קובץ הגיבוי שנבחר לא נמצא.");
            }

            if (new FileInfo(backupPath).Length <= 0)
            {
                return (false, "קובץ הגיבוי שנבחר ריק.");
            }

            if (string.Equals(Path.GetFullPath(backupPath), Path.GetFullPath(_databasePath), StringComparison.OrdinalIgnoreCase))
            {
                return (false, "לא ניתן לשחזר את קובץ מסד הנתונים הפעיל על עצמו.");
            }

            var emergencyPath = CreateBackup("Emergency_PreRestore");
            if (string.IsNullOrEmpty(emergencyPath))
            {
                return (false, "יצירת גיבוי חירום לפני השחזור נכשלה. השחזור בוטל.");
            }

            OleDbConnection.ReleaseObjectPool();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            var name = Path.GetFileNameWithoutExtension(_databasePath);
            var ext = Path.GetExtension(_databasePath);
            var tempPath = Path.Combine(Path.GetDirectoryName(_databasePath) ?? AppDomain.CurrentDomain.BaseDirectory, $"{name}_restore_{Guid.NewGuid():N}{ext}");
            try
            {
                File.Copy(backupPath, tempPath, overwrite: true);
                File.Move(tempPath, _databasePath, overwrite: true);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }

            return (true, "השחזור הושלם בהצלחה.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "RestoreDatabase failed");
            return (false, $"שחזור מסד הנתונים נכשל: {ex.Message}");
        }
    }

    public LoanMaintenanceReport RepairOrphanActiveLoans(bool applyFix, out string message)
    {
        var sampleBarcodes = new List<string>();
        var scanned = 0;
        var orphanFound = 0;
        var loansClosed = 0;
        var copiesReset = 0;

        using var connection = CreateOpenConnection();
        using var tx = applyFix ? connection.BeginTransaction() : null;

        try
        {
            var studentIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var studentsCmd = new OleDbCommand("SELECT stuID FROM [Student]", connection, tx))
            using (var reader = studentsCmd.ExecuteReader())
            {
                while (reader is not null && reader.Read())
                {
                    var stuId = reader["stuID"]?.ToString()?.Trim() ?? string.Empty;
                    var normalizedStuId = NormalizeStuId(stuId);
                    if (!string.IsNullOrWhiteSpace(stuId))
                    {
                        studentIds.Add(stuId);
                    }

                    if (!string.IsNullOrWhiteSpace(normalizedStuId))
                    {
                        studentIds.Add(normalizedStuId);
                    }
                }
            }

            var copiesByBarcode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var copiesCmd = new OleDbCommand("SELECT Barcode, CopyCondition FROM [BookCopy]", connection, tx))
            using (var reader = copiesCmd.ExecuteReader())
            {
                while (reader is not null && reader.Read())
                {
                    var barcode = reader["Barcode"]?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(barcode))
                    {
                        continue;
                    }

                    var copyCondition = reader["CopyCondition"]?.ToString()?.Trim() ?? string.Empty;
                    copiesByBarcode[barcode] = copyCondition;
                }
            }

            var processedBarcodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var historyIndexTracker = 0;
            using (var loansCmd = new OleDbCommand("SELECT Barcode, stuID FROM [Transaction] WHERE History = '0'", connection, tx))
            using (var reader = loansCmd.ExecuteReader())
            {
                while (reader is not null && reader.Read())
                {
                    scanned++;
                    var barcode = reader["Barcode"]?.ToString()?.Trim() ?? string.Empty;
                    var stuId = reader["stuID"]?.ToString()?.Trim() ?? string.Empty;
                    var normalizedStuId = NormalizeStuId(stuId);

                    var copyCondition = string.Empty;
                    var hasCopy = !string.IsNullOrWhiteSpace(barcode) && copiesByBarcode.TryGetValue(barcode, out copyCondition);
                    var isScrappedCopy = hasCopy && (copyCondition == "98" || copyCondition == "99");
                    var hasStudent = !string.IsNullOrWhiteSpace(stuId)
                                     && (studentIds.Contains(stuId)
                                         || (!string.IsNullOrWhiteSpace(normalizedStuId) && studentIds.Contains(normalizedStuId)));

                    var isOrphan = !hasCopy;
                    if (!isOrphan)
                    {
                        continue;
                    }

                    orphanFound++;
                    if (!string.IsNullOrWhiteSpace(barcode) && sampleBarcodes.Count < 15)
                    {
                        sampleBarcodes.Add(barcode);
                    }

                    if (!applyFix)
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(barcode) && processedBarcodes.Add(barcode))
                    {
                        CloseActiveLoanWithHistory(connection, tx!, barcode,
                            "ניקוי השאלה יתומה אוטומטי", ref loansClosed, ref historyIndexTracker);

                        if (hasCopy && !isScrappedCopy)
                        {
                            using var resetCopyCmd = new OleDbCommand("UPDATE [BookCopy] SET CopyReturnStatus = '0' WHERE Barcode = ?", connection, tx);
                            resetCopyCmd.Parameters.AddWithValue("@p1", barcode);
                            copiesReset += resetCopyCmd.ExecuteNonQuery();
                        }
                    }
                }
            }

            tx?.Commit();
            message = applyFix
                ? $"תחזוקת השאלות הושלמה: נסרקו {scanned}, אותרו {orphanFound}, נסגרו {loansClosed}, ואופסו {copiesReset} עותקים."
                : $"בדיקה הושלמה: נסרקו {scanned}, אותרו {orphanFound} רשומות השאלה חריגות.";

            return new LoanMaintenanceReport(scanned, orphanFound, loansClosed, copiesReset, sampleBarcodes);
        }
        catch (Exception ex)
        {
            try { tx?.Rollback(); } catch { }
            ErrorLogger.Log(ex, "RepairOrphanActiveLoans failed");
            message = $"תחזוקת השאלות נכשלה: {ex.Message}";
            return new LoanMaintenanceReport(scanned, orphanFound, loansClosed, copiesReset, sampleBarcodes);
        }
    }

    /// <summary>
    /// Silent background "Compact &amp; Repair" of the Access (MDB) database, preventing file bloat and white-space fragmentation.
    /// </summary>
    /// <remarks>
    /// Triggers: runs automatically when the main window closes (MainWindow_Closing), only if at least 30 days have passed since
    /// the last compaction, or the file has grown significantly (at least 20 MB and 1.5x) since the last compaction.
    /// Otherwise it returns immediately with no measurable cost.
    /// Safety: the whole routine is wrapped in a try-catch; failures are logged and never block shutdown or disturb the librarian.
    /// The compact itself creates a pre-compact backup and swaps the file safely.
    /// </remarks>
    public void TryAutoCompactOnShutdown()
    {
        System.Diagnostics.Debug.WriteLine("=== SHUTDOWN COMPACT ROUTINE REACHED ===");

        try
        {
            if (string.IsNullOrWhiteSpace(_databasePath) || !File.Exists(_databasePath))
            {
                return;
            }

            var (lastCompact, baselineBytes) = LoadCompactState();
            var hasLast = lastCompact != DateTime.MinValue;
            var sizeBytes = new FileInfo(_databasePath).Length;

            const long minGrowthBytes = 20L * 1024 * 1024;
            const double growthRatio = 1.5;

            var timeDue = !hasLast || DateTime.Now - lastCompact >= TimeSpan.FromDays(30);
            var sizeDue = baselineBytes > 0
                          && sizeBytes - baselineBytes >= minGrowthBytes
                          && sizeBytes >= baselineBytes * growthRatio;

            if (!timeDue && !sizeDue)
            {
                return;
            }

            System.Diagnostics.Debug.WriteLine($"[AutoCompact] Triggered (time={timeDue}, size={sizeDue}); size={sizeBytes:N0} bytes, baseline={baselineBytes:N0}.");

            OleDbConnection.ReleaseObjectPool();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            var result = CompactAndRepairDatabase();
            if (result.Success)
            {
                var newSize = new FileInfo(_databasePath).Length;
                SaveCompactState(DateTime.Now, newSize);
                System.Diagnostics.Debug.WriteLine($"[AutoCompact] Done: {sizeBytes:N0} -> {newSize:N0} bytes.");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[AutoCompact] Failed: {result.Message}");
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "TryAutoCompactOnShutdown failed");
        }
    }

    private sealed class CompactState
    {
        public string? LastCompact { get; set; }
        public long LastCompactSizeBytes { get; set; }
    }

    private static string CompactStatePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "compact_state.json");

    private static (DateTime LastCompact, long SizeBytes) LoadCompactState()
    {
        try
        {
            if (!File.Exists(CompactStatePath))
            {
                return (DateTime.MinValue, 0);
            }

            var state = JsonSerializer.Deserialize<CompactState>(File.ReadAllText(CompactStatePath));
            if (state is not null
                && DateTime.TryParse(state.LastCompact, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var last))
            {
                return (last, state.LastCompactSizeBytes);
            }
        }
        catch
        {
        }

        return (DateTime.MinValue, 0);
    }

    private static void SaveCompactState(DateTime lastCompact, long sizeBytes)
    {
        try
        {
            var state = new CompactState
            {
                LastCompact = lastCompact.ToString("o", CultureInfo.InvariantCulture),
                LastCompactSizeBytes = sizeBytes
            };
            File.WriteAllText(CompactStatePath, JsonSerializer.Serialize(state));
        }
        catch
        {
        }
    }

    public (bool Success, string Message) CompactAndRepairDatabase()
    {
        if (string.IsNullOrWhiteSpace(_databasePath) || !File.Exists(_databasePath))
        {
            return (false, "קובץ מסד הנתונים לא נמצא.");
        }

        OleDbConnection.ReleaseObjectPool();
        GC.Collect();
        GC.WaitForPendingFinalizers();

        if (string.IsNullOrEmpty(CreateBackup("PreCompact")))
        {
            return (false, "לפני דחיסה נדרש גיבוי תקין, אך יצירת הגיבוי נכשלה.");
        }

        var extension = Path.GetExtension(_databasePath);
        var dbDirectory = Path.GetDirectoryName(_databasePath) ?? AppDomain.CurrentDomain.BaseDirectory;
        var tempPath = Path.Combine(dbDirectory, $"{Path.GetFileNameWithoutExtension(_databasePath)}_compact_{Guid.NewGuid():N}{extension}");
        var swapPath = Path.Combine(dbDirectory, $"{Path.GetFileNameWithoutExtension(_databasePath)}_precompact_{Guid.NewGuid():N}{extension}");

        object? daoEngineObject = null;
        try
        {
            var daoType = Type.GetTypeFromProgID("DAO.DBEngine.120")
                          ?? Type.GetTypeFromProgID("DAO.DBEngine.36");
            if (daoType is null)
            {
                return (false, "רכיב DAO.DBEngine לא זמין במערכת זו. יש להתקין Microsoft Access Database Engine מתאים.");
            }

            daoEngineObject = Activator.CreateInstance(daoType);
            if (daoEngineObject is null)
            {
                return (false, "לא ניתן היה לאתחל את רכיב DAO.DBEngine לביצוע דחיסה.");
            }

            daoType.InvokeMember(
                "CompactDatabase",
                BindingFlags.InvokeMethod,
                binder: null,
                target: daoEngineObject,
                args: new object[] { _databasePath, tempPath });

            if (!File.Exists(tempPath))
            {
                return (false, "קובץ מסד הנתונים הדחוס לא נוצר.");
            }

            OleDbConnection.ReleaseObjectPool();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            var sourceCounts = TryGetTableRowCounts(_databasePath, out var sourceError);
            var compactCounts = TryGetTableRowCounts(tempPath, out var compactError);
            if (sourceCounts is null || compactCounts is null)
            {
                return (false, $"אימות הקובץ הדחוס נכשל, המסד המקורי לא שונה. {sourceError}{compactError}".Trim());
            }

            foreach (var table in RequiredTables)
            {
                if (!sourceCounts.TryGetValue(table, out var before) || !compactCounts.TryGetValue(table, out var after) || before != after)
                {
                    return (false, $"אימות הקובץ הדחוס נכשל (מספר השורות בטבלה {table} שונה). המסד המקורי לא שונה.");
                }
            }

            File.Replace(tempPath, _databasePath, swapPath);
            if (File.Exists(swapPath))
            {
                File.Delete(swapPath);
            }

            return (true, "דחיסה ומיטוב מסד הנתונים הושלמו בהצלחה.");
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "CompactAndRepairDatabase failed");
            return (false, $"דחיסה ומיטוב נכשלו: {ex.Message}");
        }
        finally
        {
            if (daoEngineObject is not null)
            {
                try { Marshal.FinalReleaseComObject(daoEngineObject); } catch { }
            }

            try
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
            catch
            {
            }

            try
            {
                if (File.Exists(swapPath))
                {
                    File.Delete(swapPath);
                }
            }
            catch
            {
            }
        }
    }

    public DatabaseDiagnostics GetDatabaseDiagnostics()
    {
        using var connection = CreateOpenConnection();
        using var countCommand = new OleDbCommand("SELECT COUNT(*) FROM [Student]", connection);
        var studentCount = Convert.ToInt32(countCommand.ExecuteScalar(), CultureInfo.InvariantCulture);

        return new DatabaseDiagnostics(_databasePath, "Student", "stuID", studentCount);
    }

    private string? _currentSchoolYearCache;
    private readonly object _schoolYearLock = new();

    public string GetCurrentSchoolYear()
    {
        lock (_schoolYearLock)
        {
            if (_currentSchoolYearCache is not null)
            {
                return _currentSchoolYearCache;
            }

            using var connection = CreateOpenConnection();
            using var command = new OleDbCommand("SELECT TOP 1 SettingValue FROM SystemSettings WHERE SettingKey = 'CurrentSchoolYear'", connection);
            var year = NormalizeSchoolYear(command.ExecuteScalar()?.ToString());
            if (string.IsNullOrEmpty(year))
            {
                throw new InvalidOperationException("שנת הלימודים הנוכחית אינה מוגדרת במסד הנתונים.");
            }

            _currentSchoolYearCache = year;
            return year;
        }
    }

    public bool RolloverToNewSchoolYear(string newYear, out string message)
    {
        var normalized = NormalizeSchoolYear(newYear);
        if (!System.Text.RegularExpressions.Regex.IsMatch(normalized, "^תש[א-ת]{2}$"))
        {
            message = "פורמט שנה לא תקין. יש להזין שנה עברית, לדוגמה: תשפח.";
            return false;
        }

        try
        {
            lock (_schoolYearLock)
            {
                using var connection = CreateOpenConnection();
                int affected;
                using (var update = new OleDbCommand("UPDATE SystemSettings SET SettingValue = ? WHERE SettingKey = 'CurrentSchoolYear'", connection))
                {
                    update.Parameters.AddWithValue("@p1", normalized);
                    affected = update.ExecuteNonQuery();
                }

                if (affected == 0)
                {
                    using var insert = new OleDbCommand("INSERT INTO SystemSettings (SettingKey, SettingValue) VALUES ('CurrentSchoolYear', ?)", connection);
                    insert.Parameters.AddWithValue("@p1", normalized);
                    insert.ExecuteNonQuery();
                }

                _currentSchoolYearCache = null;
            }

            message = $"שנת הלימודים עודכנה ל-{normalized}.";
            return true;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"RolloverToNewSchoolYear({normalized}) failed");
            message = "עדכון שנת הלימודים נכשל: " + ex.Message;
            return false;
        }
    }

    private bool TryGetCurrentSchoolYear(out string year)
    {
        try
        {
            year = GetCurrentSchoolYear();
            return true;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "GetCurrentSchoolYear failed");
            year = string.Empty;
            return false;
        }
    }

    public static string NormalizeSchoolYear(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return new string(value.Where(c => c != '"' && c != '\'' && c != '״' && c != '׳' && c != '”' && c != '“' && c != '’' && c != '‘' && !char.IsWhiteSpace(c)).ToArray());
    }

    private static string[] GetSchoolYearVariants(string year)
    {
        var normalized = NormalizeSchoolYear(year);
        if (normalized.Length < 2)
        {
            return new[] { normalized };
        }

        var head = normalized[..^1];
        var last = normalized[^1];
        return new[] { normalized, head + '"' + last, head + '״' + last };
    }

    public bool IsPaymentCompletedForCurrentSchoolYear(string stuId)
    {
        var normalizedStuId = (stuId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedStuId)) return false;

        if (!TryGetCurrentSchoolYear(out var currentYear)) return false;
        var idCandidates = BuildStudentIdCandidates(normalizedStuId);
        if (idCandidates.Length == 0) return false;

        using var connection = CreateFreshPaymentReadConnection();
        var idFilter = string.Join(" OR ", idCandidates.Select(_ => "stuID = ?"));

        var sql = $"SELECT Amount, ReciptNum, [Year] FROM [Payment] WHERE ({idFilter})";

        using var command = new OleDbCommand(sql, connection);
        foreach (var candidate in idCandidates)
        {
            command.Parameters.AddWithValue("@p", candidate);
        }

        using var reader = command.ExecuteReader();
        while (reader != null && reader.Read())
        {
            var yearValue = reader["Year"] == DBNull.Value ? string.Empty : reader["Year"]?.ToString()?.Trim() ?? string.Empty;
            if (!IsCurrentSchoolYearValue(yearValue, currentYear))
            {
                continue;
            }

            var amount = ReadAmountValue(reader["Amount"]);
            var reciptNum = reader["ReciptNum"] == DBNull.Value ? string.Empty : reader["ReciptNum"]?.ToString()?.Trim() ?? string.Empty;

            if (IsExemptionPaymentRow(amount, reciptNum) || amount > 0)
            {
                return true;
            }
        }

        return false;
    }

    public Student? FindStudent(string input)
    {
        var normalized = (input ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized)) return null;

        using var connection = CreateOpenConnection();
        const string activeOnlyFilter = "([stuClassID] IS NULL OR CStr([stuClassID]) <> '9999')";

        var idCandidates = BuildStudentIdCandidates(normalized);
        if (idCandidates.Length > 0)
        {
            var idFilter = string.Join(" OR ", idCandidates.Select(_ => "[stuID] = ?"));
            using var byId = new OleDbCommand($"SELECT TOP 1 * FROM [Student] WHERE ({idFilter}) AND {activeOnlyFilter}", connection);
            foreach (var candidate in idCandidates) byId.Parameters.AddWithValue("@p", candidate);

            using var reader = byId.ExecuteReader();
            if (reader is not null && reader.Read()) return WithClassName(connection, MapStudent(reader));
        }

        var parts = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 2)
        {
            using var byFullName = new OleDbCommand(
                $"SELECT TOP 2 * FROM [Student] WHERE [stuFname] LIKE ? AND [stuLname] LIKE ? AND {activeOnlyFilter}", connection);
            byFullName.Parameters.AddWithValue("@p1", $"%{parts[0]}%");
            byFullName.Parameters.AddWithValue("@p2", $"%{parts[1]}%");

            using var nameReader = byFullName.ExecuteReader();
            if (nameReader is not null && nameReader.Read())
            {
                var single = WithClassName(connection, MapStudent(nameReader));
                return nameReader.Read() ? null : single;
            }
        }

        using var bySingleName = new OleDbCommand(
            $"SELECT TOP 2 * FROM [Student] WHERE ([stuFname] LIKE ? OR [stuLname] LIKE ?) AND {activeOnlyFilter}", connection);
        bySingleName.Parameters.AddWithValue("@p1", $"%{normalized}%");
        bySingleName.Parameters.AddWithValue("@p2", $"%{normalized}%");

        using var singleReader = bySingleName.ExecuteReader();
        if (singleReader is null || !singleReader.Read()) return null;
        var onlyMatch = WithClassName(connection, MapStudent(singleReader));
        return singleReader.Read() ? null : onlyMatch;
    }

    public DataTable GetActiveLoansByStudent(string stuId)
    {
        var result = new DataTable();
        var normalizedStuId = (stuId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedStuId)) return result;

        var idCandidates = BuildStudentIdCandidates(normalizedStuId);
        if (idCandidates.Length == 0) return result;

        using var connection = CreateOpenConnection();
        var idFilter = string.Join(" OR ", idCandidates.Select(_ => "t.stuID = ?"));

        var sql = $@"SELECT t.Barcode,
                IIF(b.bkName IS NULL, 'ללא כותרת', b.bkName) AS Title,
                t.DateOfTrans AS DateBorrowed,
                t.DateToReturn,
                IIF(t.DateToReturn IS NOT NULL AND t.DateToReturn < Date(), 1, 0) AS IsOverdue
        FROM ([Transaction] t
        LEFT JOIN [BookCopy] bc ON t.Barcode = bc.Barcode)
        LEFT JOIN [Book] b ON bc.bkID = b.bkID
        WHERE ({idFilter}) AND t.History = '0'
        ORDER BY t.DateOfTrans DESC";

        using var command = new OleDbCommand(sql, connection);
        foreach (var candidate in idCandidates)
        {
            command.Parameters.AddWithValue("@p", candidate);
        }

        using var adapter = new OleDbDataAdapter(command);
        adapter.Fill(result);
        return result;
    }

    public bool TryBorrowCopy(string stuId, string copyBarcode, out string message)
    {
        var normalizedStuId = (stuId ?? string.Empty).Trim();
        var normalizedBarcode = (copyBarcode ?? string.Empty).Trim();

        if (string.IsNullOrWhiteSpace(normalizedStuId)) { message = "חניך לא תקין."; return false; }
        if (string.IsNullOrWhiteSpace(normalizedBarcode)) { message = "יש להזין ברקוד עותק."; return false; }

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();

        try
        {
            var stuCandidates = BuildStudentIdCandidates(normalizedStuId);
            var stuFilter = string.Join(" OR ", stuCandidates.Select(_ => "[stuID] = ?"));
            string? resolvedStuId;
            using (var studentCommand = new OleDbCommand($"SELECT TOP 1 [stuID] FROM [Student] WHERE {stuFilter}", connection, tx))
            {
                foreach (var candidate in stuCandidates) studentCommand.Parameters.AddWithValue("@p", candidate);
                resolvedStuId = studentCommand.ExecuteScalar()?.ToString()?.Trim();
            }

            if (string.IsNullOrWhiteSpace(resolvedStuId))
            {
                tx.Rollback();
                message = "חניך לא נמצא במערכת.";
                return false;
            }

            string? bkId = null;
            using (var copyCmd = new OleDbCommand("SELECT TOP 1 bkID FROM [BookCopy] WHERE Barcode = ?", connection, tx))
            {
                copyCmd.Parameters.AddWithValue("@p1", normalizedBarcode);
                var bkObj = copyCmd.ExecuteScalar();
                if (bkObj == null)
                {
                    tx.Rollback();
                    message = $"עותק עם ברקוד {normalizedBarcode} אינו קיים במאגר.";
                    return false;
                }
                bkId = bkObj.ToString();
            }

            using (var activeCheck = new OleDbCommand("SELECT COUNT(*) FROM [Transaction] WHERE Barcode = ? AND History = '0'", connection, tx))
            {
                activeCheck.Parameters.AddWithValue("@p1", normalizedBarcode);
                if (Convert.ToInt32(activeCheck.ExecuteScalar() ?? 0) > 0)
                {
                    message = BuildAlreadyBorrowedMessage(GetActiveLoanBorrowerId(connection, tx, normalizedBarcode), resolvedStuId);
                    tx.Rollback();
                    return false;
                }
            }

            bool recordExists;
            using (var existsCmd = new OleDbCommand("SELECT COUNT(*) FROM [Transaction] WHERE Barcode = ?", connection, tx))
            {
                existsCmd.Parameters.AddWithValue("@p1", normalizedBarcode);
                recordExists = Convert.ToInt32(existsCmd.ExecuteScalar() ?? 0) > 0;
            }

            var defaultReturnDate = GetEndOfSchoolYearReturnDate();
            if (recordExists)
            {
                using var updateLoan = new OleDbCommand(
                    "UPDATE [Transaction] SET stuID = ?, bkID = ?, DateOfTrans = ?, DateToReturn = ?, History = '0' WHERE Barcode = ?", connection, tx);
                updateLoan.Parameters.AddWithValue("@p1", resolvedStuId);
                updateLoan.Parameters.AddWithValue("@p2", bkId ?? (object)DBNull.Value);
                updateLoan.Parameters.Add("@p3", OleDbType.Date).Value = DateTime.Now;
                updateLoan.Parameters.Add("@p4", OleDbType.Date).Value = defaultReturnDate;
                updateLoan.Parameters.AddWithValue("@p5", normalizedBarcode);
                updateLoan.ExecuteNonQuery();
            }
            else
            {
                using var insertLoan = new OleDbCommand(
                    "INSERT INTO [Transaction] (Barcode, stuID, bkID, DateOfTrans, DateToReturn, History) VALUES (?, ?, ?, ?, ?, '0')", connection, tx);
                insertLoan.Parameters.AddWithValue("@p1", normalizedBarcode);
                insertLoan.Parameters.AddWithValue("@p2", resolvedStuId);
                insertLoan.Parameters.AddWithValue("@p3", bkId ?? (object)DBNull.Value);
                insertLoan.Parameters.Add("@p4", OleDbType.Date).Value = DateTime.Now;
                insertLoan.Parameters.Add("@p5", OleDbType.Date).Value = defaultReturnDate;
                insertLoan.ExecuteNonQuery();
            }


            using (var updateCopy = new OleDbCommand("UPDATE [BookCopy] SET CopyReturnStatus = '1' WHERE Barcode = ?", connection, tx))
            {
                updateCopy.Parameters.AddWithValue("@p1", normalizedBarcode);
                updateCopy.ExecuteNonQuery();
            }

            tx.Commit();
            message = "ההשאלה בוצעה בהצלחה!";
            return true;
        }
        catch (Exception ex)
        {
            try { tx.Rollback(); } catch { }
            ErrorLogger.Log(ex, $"Borrow failed: {copyBarcode}");
            message = $"שגיאת ביצוע: {ex.Message}";
            return false;
        }
    }

    public bool TryExtendLoan(string barcode, DateTime newReturnDate, out string message)
    {
        var normalizedBarcode = (barcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedBarcode))
        {
            message = "יש להזין ברקוד תקין להארכת ההשאלה.";
            return false;
        }

        try
        {
            using var connection = CreateOpenConnection();

            using (var loanDateCmd = new OleDbCommand("SELECT DateOfTrans FROM [Transaction] WHERE Barcode = ? AND History = '0'", connection))
            {
                loanDateCmd.Parameters.AddWithValue("@p1", normalizedBarcode);
                var loanDateValue = loanDateCmd.ExecuteScalar();
                if (loanDateValue is not null && loanDateValue != DBNull.Value)
                {
                    var loanDate = Convert.ToDateTime(loanDateValue).Date;
                    if (newReturnDate.Date < loanDate)
                    {
                        message = $"תאריך ההחזרה לא יכול להיות מוקדם מתאריך ההשאלה ({loanDate:dd/MM/yyyy}).";
                        return false;
                    }
                }
            }

            if (newReturnDate.Date < DateTime.Today)
            {
                message = "לא ניתן לקבוע תאריך החזרה שכבר עבר.";
                return false;
            }

            if (newReturnDate.Date > DateTime.Today.AddYears(2))
            {
                message = "תאריך ההחזרה רחוק מדי. הטווח המותר הוא עד שנתיים מהיום.";
                return false;
            }

            using var cmd = new OleDbCommand("UPDATE [Transaction] SET DateToReturn = ? WHERE Barcode = ? AND History = '0'", connection);
            cmd.Parameters.Add("@p1", OleDbType.Date).Value = newReturnDate.Date;
            cmd.Parameters.AddWithValue("@p2", normalizedBarcode);

            var affectedRows = cmd.ExecuteNonQuery();
            if (affectedRows <= 0)
            {
                message = "לא נמצאה השאלה פעילה עבור הברקוד שנבחר.";
                return false;
            }

            message = $"מועד ההחזרה עודכן ל-{newReturnDate:dd/MM/yyyy}.";
            return true;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"Extend loan failed: {normalizedBarcode}");
            message = $"הארכת מועד ההחזרה נכשלה: {ex.Message}";
            return false;
        }
    }

    public enum ReturnOutcome { Ok, NotBorrowed, Error, ScrappedOrLost }

    public bool TryReturnCopy(string copyBarcode, out string message)
    {
        return TryReturnCopyWithBorrowerInfo(copyBarcode, out message, out _) == ReturnOutcome.Ok;
    }

    public bool IsCopyMarkedScrappedOrLost(string copyBarcode)
    {
        var barcode = (copyBarcode ?? string.Empty).Trim();
        if (barcode.Length == 0)
        {
            return false;
        }

        try
        {
            using var connection = CreateOpenConnection();
            using var cmd = new OleDbCommand("SELECT TOP 1 CopyCondition FROM [BookCopy] WHERE Trim(CStr(Barcode)) = ?", connection);
            cmd.Parameters.Add("@p1", OleDbType.VarWChar).Value = barcode;
            using var reader = cmd.ExecuteReader();
            if (reader == null || !reader.Read())
            {
                return false;
            }

            var condition = reader["CopyCondition"]?.ToString()?.Trim() ?? string.Empty;
            return condition == "98" || condition == "99";
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"IsCopyMarkedScrappedOrLost failed: {barcode}");
            return false;
        }
    }

    private static void RestoreCopyAsUsed(OleDbConnection connection, OleDbTransaction tx, string barcode)
    {
        using var cmd = new OleDbCommand("UPDATE [BookCopy] SET CopyCondition = '2' WHERE Barcode = ? AND Trim(CopyCondition & '') IN ('98','99')", connection, tx);
        cmd.Parameters.AddWithValue("@p1", barcode);
        cmd.ExecuteNonQuery();
    }

    public ReturnOutcome TryReturnCopyWithBorrowerInfo(string copyBarcode, out string message, out Student? borrowerInfo, bool restoreScrappedAsUsed = false)
    {
        borrowerInfo = null;
        var normalizedBarcode = (copyBarcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedBarcode)) { message = "יש להזין ברקוד עותק."; return ReturnOutcome.Error; }

        if (IsCopyMarkedScrappedOrLost(normalizedBarcode))
        {
            message = "לא ניתן להחזיר עותק גרוט או אבוד. יש לשחזר אותו תחילה במסך ניהול כותרים ועותקים.";
            return ReturnOutcome.ScrappedOrLost;
        }

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();

        try
        {
            string borrowerStuId = string.Empty;
            string bkId = string.Empty;
            DateTime? dateOfTrans = null;

            if (restoreScrappedAsUsed)
            {
                RestoreCopyAsUsed(connection, tx, normalizedBarcode);
            }

            using (var findCmd = new OleDbCommand("SELECT stuID, bkID, DateOfTrans FROM [Transaction] WHERE Barcode = ? AND History = '0'", connection, tx))
            {
                findCmd.Parameters.AddWithValue("@p1", normalizedBarcode);
                using var reader = findCmd.ExecuteReader();
                if (reader == null || !reader.Read())
                {
                    reader?.Close();
                    if (restoreScrappedAsUsed)
                    {
                        tx.Commit();
                        message = "העותק שוחזר למלאי הפעיל כעותק משומש, אך לא נמצאה השאלה פעילה עבורו.";
                        return ReturnOutcome.NotBorrowed;
                    }

                    tx.Rollback();
                    if (IsCopyMarkedScrappedOrLost(normalizedBarcode))
                    {
                        message = "לא ניתן להחזיר עותק גרוט או אבוד. יש לשחזר אותו תחילה במסך ניהול כותרים ועותקים.";
                        return ReturnOutcome.ScrappedOrLost;
                    }

                    message = "לא נמצאה השאלה פעילה עבור עותק זה.";
                    return ReturnOutcome.NotBorrowed;
                }
                borrowerStuId = reader["stuID"]?.ToString()?.Trim() ?? string.Empty;
                bkId = reader["bkID"]?.ToString()?.Trim() ?? string.Empty;
                if (reader["DateOfTrans"] != DBNull.Value)
                {
                    dateOfTrans = Convert.ToDateTime(reader["DateOfTrans"], CultureInfo.InvariantCulture);
                }
            }

            if (!string.IsNullOrWhiteSpace(borrowerStuId))
            {
                var stuCandidates = BuildStudentIdCandidates(borrowerStuId);
                var stuFilter = string.Join(" OR ", stuCandidates.Select(_ => "[stuID] = ?"));
                using var borrowerCmd = new OleDbCommand($"SELECT TOP 1 * FROM [Student] WHERE {stuFilter}", connection, tx);
                foreach (var c in stuCandidates) borrowerCmd.Parameters.AddWithValue("@p", c);

                using var reader = borrowerCmd.ExecuteReader();
                if (reader != null && reader.Read())
                {
                    borrowerInfo = MapStudent(reader);
                }
            }

            using (var updateLoan = new OleDbCommand("UPDATE [Transaction] SET History = '1' WHERE Barcode = ?", connection, tx))
            {
                updateLoan.Parameters.AddWithValue("@p1", normalizedBarcode);
                updateLoan.ExecuteNonQuery();
            }

            var currentHistoryIndex = 0;
            InsertHistoryRow(connection, tx, normalizedBarcode, borrowerStuId, bkId, dateOfTrans, true, string.Empty, ref currentHistoryIndex);

            using (var updateCopy = new OleDbCommand("UPDATE [BookCopy] SET CopyReturnStatus = '0' WHERE Barcode = ?", connection, tx))
            {
                updateCopy.Parameters.AddWithValue("@p1", normalizedBarcode);
                updateCopy.ExecuteNonQuery();
            }

            tx.Commit();
            message = "ההחזרה בוצעה בהצלחה!";
            return ReturnOutcome.Ok;
        }
        catch (Exception ex)
        {
            try { tx.Rollback(); } catch { }
            ErrorLogger.Log(ex, $"Return failed: {copyBarcode}");
            message = $"שגיאת החזרה, ההשאלה נשארה פתוחה: {ex.Message}";
            return ReturnOutcome.Error;
        }
    }

    private static void InsertHistoryRow(OleDbConnection connection, OleDbTransaction tx,
        string barcode, string stuId, string bkId, DateTime? dateOfTrans, bool normalReturn, string remarks,
        ref int historyIndexTracker)
    {
        if (historyIndexTracker <= 0)
        {
            historyIndexTracker = GetMaxHistoryIndex(connection, tx);
        }

        historyIndexTracker++;

        var bkName = string.Empty;
        var endCondition = string.Empty;
        var startCondition = string.Empty;

        using (var infoCmd = new OleDbCommand(
            @"SELECT TOP 1 b.bkName, bc.CopyCondition
              FROM [BookCopy] bc LEFT JOIN [Book] b ON bc.bkID = b.bkID
              WHERE bc.Barcode = ?", connection, tx))
        {
            infoCmd.Parameters.AddWithValue("@p1", barcode);
            using var r = infoCmd.ExecuteReader();
            if (r is not null && r.Read())
            {
                bkName = r["bkName"]?.ToString()?.Trim() ?? string.Empty;
                endCondition = r["CopyCondition"]?.ToString()?.Trim() ?? string.Empty;
            }
        }

        using (var startCmd = new OleDbCommand("SELECT TOP 1 [Condition] FROM [Transaction] WHERE Barcode = ?", connection, tx))
        {
            startCmd.Parameters.AddWithValue("@p1", barcode);
            startCondition = startCmd.ExecuteScalar()?.ToString()?.Trim() ?? string.Empty;
        }

        using var cmd = new OleDbCommand(
            @"INSERT INTO [History] ([Index], Barcode, DateOfReturn, stuID, bkID, DateOfTrans, NormalReturn, bkName, Remarks, StartCondition, EndCondition)
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", connection, tx);
        cmd.Parameters.AddWithValue("@p0", historyIndexTracker.ToString(CultureInfo.InvariantCulture));
        cmd.Parameters.AddWithValue("@p1", barcode);
        cmd.Parameters.Add("@p2", OleDbType.Date).Value = DateTime.Now;
        cmd.Parameters.AddWithValue("@p3", stuId);
        cmd.Parameters.AddWithValue("@p4", bkId);
        if (dateOfTrans.HasValue)
            cmd.Parameters.Add("@p5", OleDbType.Date).Value = dateOfTrans.Value;
        else
            cmd.Parameters.AddWithValue("@p5", DBNull.Value);
        cmd.Parameters.AddWithValue("@p6", normalReturn ? "1" : "0");
        cmd.Parameters.AddWithValue("@p7", bkName);
        cmd.Parameters.AddWithValue("@p8", remarks);
        cmd.Parameters.AddWithValue("@p9", startCondition);
        cmd.Parameters.AddWithValue("@p10", endCondition);
        cmd.ExecuteNonQuery();
    }

    private static void CloseActiveLoanWithHistory(OleDbConnection connection, OleDbTransaction tx,
        string barcode, string remarks, ref int closed, ref int historyIndexTracker)
    {
        var loans = new List<(string StuId, string BkId, DateTime? Date)>();
        using (var sel = new OleDbCommand(
            "SELECT stuID, bkID, DateOfTrans FROM [Transaction] WHERE Barcode = ? AND History = '0'", connection, tx))
        {
            sel.Parameters.AddWithValue("@p1", barcode);
            using var r = sel.ExecuteReader();
            while (r is not null && r.Read())
            {
                DateTime? d = r["DateOfTrans"] == DBNull.Value ? null : Convert.ToDateTime(r["DateOfTrans"], CultureInfo.InvariantCulture);
                loans.Add((r["stuID"]?.ToString()?.Trim() ?? string.Empty, r["bkID"]?.ToString()?.Trim() ?? string.Empty, d));
            }
        }

        using (var upd = new OleDbCommand("UPDATE [Transaction] SET History = '1' WHERE Barcode = ? AND History = '0'", connection, tx))
        {
            upd.Parameters.AddWithValue("@p1", barcode);
            closed += upd.ExecuteNonQuery();
        }

        foreach (var l in loans)
        {
            InsertHistoryRow(connection, tx, barcode, l.StuId, l.BkId, l.Date, false, remarks, ref historyIndexTracker);
        }
    }

    private static int GetMaxHistoryIndex(OleDbConnection connection, OleDbTransaction tx)
    {
        using var cmd = new OleDbCommand("SELECT MAX(Val([Index])) FROM [History]", connection, tx);
        var result = cmd.ExecuteScalar();
        return result == null || result == DBNull.Value
            ? 0
            : (int)Convert.ToDouble(result, CultureInfo.InvariantCulture);
    }

    public bool TryGetBorrowCartItem(string copyBarcode, out string title, out string priceDisplay, out string message, string? currentStuId = null)
    {
        title = string.Empty;
        priceDisplay = "0";
        var normalizedBarcode = (copyBarcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedBarcode))
        {
            message = "יש להזין ברקוד עותק.";
            return false;
        }

        using var connection = CreateOpenConnection();

        using (var copyCmd = new OleDbCommand(@"SELECT TOP 1 bc.bkID,
                                                   IIF(b.bkName IS NULL, '', b.bkName) AS Title,
                                                   IIF(b.bkPrice IS NULL, '0', CStr(b.bkPrice)) AS PriceDisplay,
                                                   IIF(bc.CopyCondition IS NULL, '', CStr(bc.CopyCondition)) AS CopyCondition
                                                FROM [BookCopy] bc
                                                LEFT JOIN [Book] b ON bc.bkID = b.bkID
                                                WHERE bc.Barcode = ?", connection))
        {
            copyCmd.Parameters.AddWithValue("@p1", normalizedBarcode);
            using var reader = copyCmd.ExecuteReader();
            if (reader is null || !reader.Read())
            {
                message = $"עותק עם ברקוד {normalizedBarcode} אינו קיים במאגר.";
                return false;
            }

            var copyCondition = (reader["CopyCondition"]?.ToString() ?? string.Empty).Trim();
            if (copyCondition == "98" || copyCondition == "99")
            {
                message = "העותק מסומן כגרוט ואינו זמין להשאלה.";
                return false;
            }

            title = (reader["Title"]?.ToString() ?? string.Empty).Trim();
            priceDisplay = (reader["PriceDisplay"]?.ToString() ?? "0").Trim();
        }

        using (var activeLoanCmd = new OleDbCommand("SELECT COUNT(*) FROM [Transaction] WHERE Barcode = ? AND History = '0'", connection))
        {
            activeLoanCmd.Parameters.AddWithValue("@p1", normalizedBarcode);
            if (Convert.ToInt32(activeLoanCmd.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture) > 0)
            {
                message = BuildAlreadyBorrowedMessage(GetActiveLoanBorrowerId(connection, null, normalizedBarcode), currentStuId);
                return false;
            }
        }

        message = "העתק נוסף לרשימת ההשאלה.";
        return true;
    }

    public bool TryGetReturnCartItem(string copyBarcode, out string title, out string borrowerName, out string message)
    {
        title = string.Empty;
        borrowerName = string.Empty;
        var normalizedBarcode = (copyBarcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedBarcode))
        {
            message = "יש להזין ברקוד עותק.";
            return false;
        }

        if (IsCopyMarkedScrappedOrLost(normalizedBarcode))
        {
            message = "לא ניתן להחזיר עותק גרוט או אבוד. יש לשחזר אותו תחילה במסך ניהול כותרים ועותקים.";
            return false;
        }

        using var connection = CreateOpenConnection();
        using var cmd = new OleDbCommand(@"SELECT TOP 1
                                               IIF(b.bkName IS NULL, '', b.bkName) AS BookTitle,
                                               Trim(s.stuFname & ' ' & s.stuLname) AS StudentName,
                                               t.stuID
                                         FROM (([Transaction] t
                                         LEFT JOIN [BookCopy] bc ON t.Barcode = bc.Barcode)
                                         LEFT JOIN [Book] b ON bc.bkID = b.bkID)
                                         LEFT JOIN [Student] s ON t.stuID = s.stuID
                                         WHERE t.Barcode = ? AND t.History = '0'", connection);
        cmd.Parameters.AddWithValue("@p1", normalizedBarcode);
        using var reader = cmd.ExecuteReader();
        if (reader is null || !reader.Read())
        {
            message = "לא נמצאה השאלה פעילה עבור עותק זה.";
            return false;
        }

        title = (reader["BookTitle"]?.ToString() ?? string.Empty).Trim();
        borrowerName = (reader["StudentName"]?.ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(borrowerName))
        {
            var stuId = (reader["stuID"]?.ToString() ?? string.Empty).Trim();
            borrowerName = string.IsNullOrWhiteSpace(stuId) ? "-" : $"חניך לא במצבת ({stuId})";
        }

        message = "העותק נוסף לרשימת ההחזרה.";
        return true;
    }

    public bool TryBorrowCopiesBatch(string stuId, IEnumerable<string> barcodes, out string message)
    {
        var normalizedStuId = (stuId ?? string.Empty).Trim();
        var list = (barcodes ?? Array.Empty<string>())
            .Select(b => (b ?? string.Empty).Trim())
            .Where(b => !string.IsNullOrWhiteSpace(b))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (string.IsNullOrWhiteSpace(normalizedStuId))
        {
            message = "חניך לא תקין.";
            return false;
        }

        if (list.Count == 0)
        {
            message = "לא נבחרו ספרים להשאלה.";
            return false;
        }

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();

        try
        {
            var stuCandidates = BuildStudentIdCandidates(normalizedStuId);
            var stuFilter = string.Join(" OR ", stuCandidates.Select(_ => "[stuID] = ?"));
            string? resolvedStuId;
            using (var studentCommand = new OleDbCommand($"SELECT TOP 1 [stuID] FROM [Student] WHERE {stuFilter}", connection, tx))
            {
                foreach (var candidate in stuCandidates)
                {
                    studentCommand.Parameters.AddWithValue("@p", candidate);
                }

                resolvedStuId = studentCommand.ExecuteScalar()?.ToString()?.Trim();
            }

            if (string.IsNullOrWhiteSpace(resolvedStuId))
            {
                tx.Rollback();
                message = "חניך לא נמצא במערכת.";
                return false;
            }

            foreach (var normalizedBarcode in list)
            {
                string? bkId;
                using (var copyCmd = new OleDbCommand("SELECT TOP 1 bkID FROM [BookCopy] WHERE Barcode = ? AND (CopyCondition IS NULL OR Trim(CStr(CopyCondition)) NOT IN ('98','99'))", connection, tx))
                {
                    copyCmd.Parameters.AddWithValue("@p1", normalizedBarcode);
                    var bkObj = copyCmd.ExecuteScalar();
                    if (bkObj is null)
                    {
                        tx.Rollback();
                        message = $"עותק עם ברקוד {normalizedBarcode} אינו זמין להשאלה.";
                        return false;
                    }

                    bkId = bkObj.ToString();
                }

                using (var activeCheck = new OleDbCommand("SELECT COUNT(*) FROM [Transaction] WHERE Barcode = ? AND History = '0'", connection, tx))
                {
                    activeCheck.Parameters.AddWithValue("@p1", normalizedBarcode);
                    if (Convert.ToInt32(activeCheck.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture) > 0)
                    {
                        tx.Rollback();
                        message = $"העותק {normalizedBarcode} כבר מושאל.";
                        return false;
                    }
                }

                bool recordExists;
                using (var existsCmd = new OleDbCommand("SELECT COUNT(*) FROM [Transaction] WHERE Barcode = ?", connection, tx))
                {
                    existsCmd.Parameters.AddWithValue("@p1", normalizedBarcode);
                    recordExists = Convert.ToInt32(existsCmd.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture) > 0;
                }
                var defaultReturnDate = GetEndOfSchoolYearReturnDate();

                if (recordExists)
                {
                    using var updateLoan = new OleDbCommand(
                        "UPDATE [Transaction] SET stuID = ?, bkID = ?, DateOfTrans = ?, DateToReturn = ?, History = '0' WHERE Barcode = ?", connection, tx);
                    updateLoan.Parameters.AddWithValue("@p1", resolvedStuId);
                    updateLoan.Parameters.AddWithValue("@p2", bkId ?? (object)DBNull.Value);
                    updateLoan.Parameters.Add("@p3", OleDbType.Date).Value = DateTime.Now;
                    updateLoan.Parameters.Add("@p4", OleDbType.Date).Value = defaultReturnDate;
                    updateLoan.Parameters.AddWithValue("@p5", normalizedBarcode);
                    updateLoan.ExecuteNonQuery();
                }
                else
                {
                    using var insertLoan = new OleDbCommand(
                        "INSERT INTO [Transaction] (Barcode, stuID, bkID, DateOfTrans, DateToReturn, History) VALUES (?, ?, ?, ?, ?, '0')", connection, tx);
                    insertLoan.Parameters.AddWithValue("@p1", normalizedBarcode);
                    insertLoan.Parameters.AddWithValue("@p2", resolvedStuId);
                    insertLoan.Parameters.AddWithValue("@p3", bkId ?? (object)DBNull.Value);
                    insertLoan.Parameters.Add("@p4", OleDbType.Date).Value = DateTime.Now;
                    insertLoan.Parameters.Add("@p5", OleDbType.Date).Value = defaultReturnDate;
                    insertLoan.ExecuteNonQuery();
                }


                using (var updateCopy = new OleDbCommand("UPDATE [BookCopy] SET CopyReturnStatus = '1' WHERE Barcode = ?", connection, tx))
                {
                    updateCopy.Parameters.AddWithValue("@p1", normalizedBarcode);
                    updateCopy.ExecuteNonQuery();
                }
            }

            tx.Commit();
            message = $"בוצעה השאלה ל-{list.Count} עותקים.";
            return true;
        }
        catch (Exception ex)
        {
            try { tx.Rollback(); } catch { }
            ErrorLogger.Log(ex, "Borrow batch failed");
            message = $"שגיאת ביצוע: {ex.Message}";
            return false;
        }
    }

    public bool TryReturnCopiesBatch(IEnumerable<string> barcodes, out string message, ISet<string>? restoreScrappedBarcodes = null)
    {
        var list = (barcodes ?? Array.Empty<string>())
            .Select(b => (b ?? string.Empty).Trim())
            .Where(b => !string.IsNullOrWhiteSpace(b))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (list.Count == 0)
        {
            message = "לא נבחרו ספרים להחזרה.";
            return false;
        }

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();

        try
        {
            var currentHistoryIndex = 0;
            foreach (var normalizedBarcode in list)
            {
                string borrowerStuId = string.Empty;
                string bkId = string.Empty;
                DateTime? dateOfTrans = null;

                using (var findCmd = new OleDbCommand("SELECT stuID, bkID, DateOfTrans FROM [Transaction] WHERE Barcode = ? AND History = '0'", connection, tx))
                {
                    findCmd.Parameters.AddWithValue("@p1", normalizedBarcode);
                    using var reader = findCmd.ExecuteReader();
                    if (reader == null || !reader.Read())
                    {
                        reader?.Close();
                        tx.Rollback();
                        message = $"לא נמצאה השאלה פעילה עבור {normalizedBarcode}.";
                        return false;
                    }
                    borrowerStuId = reader["stuID"]?.ToString()?.Trim() ?? string.Empty;
                    bkId = reader["bkID"]?.ToString()?.Trim() ?? string.Empty;
                    if (reader["DateOfTrans"] != DBNull.Value)
                    {
                        dateOfTrans = Convert.ToDateTime(reader["DateOfTrans"], CultureInfo.InvariantCulture);
                    }
                }

                using (var updateLoan = new OleDbCommand("UPDATE [Transaction] SET History = '1' WHERE Barcode = ?", connection, tx))
                {
                    updateLoan.Parameters.AddWithValue("@p1", normalizedBarcode);
                    updateLoan.ExecuteNonQuery();
                }

                InsertHistoryRow(connection, tx, normalizedBarcode, borrowerStuId, bkId, dateOfTrans, true, string.Empty, ref currentHistoryIndex);

                using (var updateCopy = new OleDbCommand("UPDATE [BookCopy] SET CopyReturnStatus = '0' WHERE Barcode = ?", connection, tx))
                {
                    updateCopy.Parameters.AddWithValue("@p1", normalizedBarcode);
                    updateCopy.ExecuteNonQuery();
                }

                if (restoreScrappedBarcodes is not null && restoreScrappedBarcodes.Contains(normalizedBarcode))
                {
                    RestoreCopyAsUsed(connection, tx, normalizedBarcode);
                }
            }

            tx.Commit();
            message = $"בוצעה החזרה של {list.Count} עותקים.";
            return true;
        }
        catch (Exception ex)
        {
            try { tx.Rollback(); } catch { }
            ErrorLogger.Log(ex, "Return batch failed");
            message = $"שגיאת החזרה: {ex.Message}";
            return false;
        }
    }

    private void EnsureSystemSettings()
    {
        try
        {
            using var connection = CreateOpenConnection();
            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            bool exists = false;
            if (schema != null)
            {
                foreach (DataRow row in schema.Rows)
                {
                    if (string.Equals(row["TABLE_NAME"]?.ToString(), "SystemSettings", StringComparison.OrdinalIgnoreCase))
                    {
                        exists = true;
                        break;
                    }
                }
            }

            if (!exists)
            {
                using var cmd = new OleDbCommand("CREATE TABLE SystemSettings (SettingKey TEXT(50) PRIMARY KEY, SettingValue TEXT(255))", connection);
                cmd.ExecuteNonQuery();
                using var ins = new OleDbCommand("INSERT INTO SystemSettings (SettingKey, SettingValue) VALUES ('CurrentSchoolYear', 'תשפז')", connection);
                ins.ExecuteNonQuery();
            }
        }
        catch { }
    }

    private static string[] BuildStudentIdCandidates(string input)
    {
        var normalized = (input ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized)) return Array.Empty<string>();

        var digits = new string(normalized.Where(char.IsDigit).ToArray());
        var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { normalized };

        if (!string.IsNullOrWhiteSpace(digits))
        {
            candidates.Add(digits);
            if (digits.Length < 9) candidates.Add(digits.PadLeft(9, '0'));
            else if (digits.Length > 9) candidates.Add(digits[^9..]);

            var core = digits.TrimStart('0');
            if (core.Length > 0 && core.Length <= 9)
            {
                for (var length = core.Length; length <= 12; length++)
                {
                    candidates.Add(core.PadLeft(length, '0'));
                }
            }
        }

        return candidates.ToArray();
    }

    private static bool IsCurrentSchoolYearValue(string? paymentYear, string currentYear)
    {
        var normalizedPaymentYear = NormalizeSchoolYear(paymentYear);
        var normalizedCurrentYear = NormalizeSchoolYear(currentYear);
        if (normalizedPaymentYear.Length == 0 || normalizedCurrentYear.Length == 0)
        {
            return false;
        }

        return string.Equals(normalizedPaymentYear, normalizedCurrentYear, StringComparison.OrdinalIgnoreCase);
    }

    private static decimal ReadAmountValue(object? amountValue)
    {
        if (amountValue is null || amountValue == DBNull.Value)
        {
            return 0;
        }

        if (amountValue is decimal decValue)
        {
            return decValue;
        }

        if (amountValue is double doubleValue)
        {
            return Convert.ToDecimal(doubleValue, CultureInfo.InvariantCulture);
        }

        if (amountValue is float floatValue)
        {
            return Convert.ToDecimal(floatValue, CultureInfo.InvariantCulture);
        }

        var raw = amountValue.ToString()?.Trim() ?? string.Empty;
        if (decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedInvariant)
            || decimal.TryParse(raw, NumberStyles.Any, CultureInfo.CurrentCulture, out parsedInvariant))
        {
            return parsedInvariant;
        }

        return 0;
    }

    private static bool IsExemptionPaymentRow(decimal amount, string? reciptNum)
    {
        var normalizedReceipt = (reciptNum ?? string.Empty).Trim();
        return amount == 0 && string.Equals(normalizedReceipt, "פטור מאושר", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeStuId(string rawId)
    {
        var normalized = (rawId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return string.Empty;
        }

        if (normalized.All(char.IsDigit) && normalized.Length < 9)
        {
            normalized = normalized.PadLeft(9, '0');
        }

        return normalized;
    }

    private static Student WithClassName(OleDbConnection connection, Student student)
    {
        try
        {
            using var nameCmd = new OleDbCommand("SELECT ClassName FROM [Classes] WHERE ClassID = ?", connection);
            nameCmd.Parameters.AddWithValue("@p1", student.ClassCode);
            student.ClassName = Convert.ToString(nameCmd.ExecuteScalar())?.Trim() ?? string.Empty;
        }
        catch
        {
            student.ClassName = string.Empty;
        }

        return student;
    }

    private static Student MapStudent(OleDbDataReader reader)
    {
        return new Student
        {
            StuId = Convert.ToString(reader["stuID"])?.Trim() ?? string.Empty,
            FirstName = Convert.ToString(reader["stuFname"])?.Trim() ?? string.Empty,
            LastName = Convert.ToString(reader["stuLname"])?.Trim() ?? string.Empty,
            ClassCode = Convert.ToString(reader["stuClassID"])?.Trim() ?? "0000"
        };
    }

    public DataTable GetSubjects()
    {
        var dt = new DataTable();
        using var connection = CreateOpenConnection();
        using var cmd = new OleDbCommand("SELECT SubjectID, SubjectName FROM [Subject] ORDER BY SubjectName", connection);
        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    private int ResolveSubjectIdByName(string? subjectName)
    {
        var cleanSubject = (subjectName ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanSubject))
        {
            return 0;
        }

        using var connection = CreateOpenConnection();
        using var cmd = new OleDbCommand("SELECT TOP 1 SubjectID FROM [Subject] WHERE SubjectName = ?", connection);
        cmd.Parameters.AddWithValue("@p1", cleanSubject);
        var value = cmd.ExecuteScalar();
        if (value is null || value == DBNull.Value)
        {
            return 0;
        }

        return int.TryParse(value.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var id)
            ? id
            : 0;
    }

    public string NormalizeStudentIdForDisplay(string rawId)
    {
        return NormalizeStudentIdForSync(rawId);
    }

    public bool TryAddManualStudent(string rawStuId, string firstName, string lastName, string classCode, out string message)
    {
        var normalizedStuId = NormalizeStudentIdForSync(rawStuId);
        var normalizedFirstName = (firstName ?? string.Empty).Trim();
        var normalizedLastName = (lastName ?? string.Empty).Trim();
        var normalizedClassCode = NormalizeClassCode(classCode);

        if (string.IsNullOrWhiteSpace(normalizedStuId) || normalizedStuId.Length != 9 || !normalizedStuId.All(char.IsDigit))
        {
            message = "תעודת הזהות חייבת להכיל בדיוק 9 ספרות.";
            return false;
        }

        if (!IsValidIsraeliId(normalizedStuId))
        {
            message = "תעודת הזהות אינה תקינה (ספרת ביקורת שגויה).";
            return false;
        }

        if (string.IsNullOrWhiteSpace(normalizedFirstName) || string.IsNullOrWhiteSpace(normalizedLastName))
        {
            message = "שם פרטי ושם משפחה הם שדות חובה.";
            return false;
        }

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();

        try
        {
            var existingRemarks = string.Empty;
            var studentExists = false;
            var existingStuId = string.Empty;
            var existingClassId = string.Empty;
            var idCandidates = BuildStudentIdCandidates(normalizedStuId);
            if (idCandidates.Length == 0)
            {
                tx.Rollback();
                message = "חניך לא תקין.";
                return false;
            }

            var existingFilter = string.Join(" OR ", idCandidates.Select(_ => "stuID = ?"));
            using (var existsCmd = new OleDbCommand($"SELECT TOP 1 stuID, stuRemarks, stuClassID FROM [Student] WHERE {existingFilter}", connection, tx))
            {
                foreach (var candidate in idCandidates)
                {
                    existsCmd.Parameters.AddWithValue("@p", candidate);
                }

                using var reader = existsCmd.ExecuteReader();
                if (reader is not null && reader.Read())
                {
                    studentExists = true;
                    existingStuId = Convert.ToString(reader["stuID"])?.Trim() ?? normalizedStuId;
                    existingRemarks = Convert.ToString(reader["stuRemarks"])?.Trim() ?? string.Empty;
                    existingClassId = Convert.ToString(reader["stuClassID"])?.Trim() ?? string.Empty;
                }
            }

            if (studentExists && string.Equals(existingClassId, "9999", StringComparison.Ordinal))
            {
                tx.Rollback();
                message = "תעודת זהות זו כבר קיימת במערכת עבור חניך לא פעיל. אנא הצג לא פעילים ושחזר אותו מהארכיון.";
                return false;
            }

            if (!studentExists)
            {
                using var insertCmd = new OleDbCommand("INSERT INTO [Student] (stuID, stuFname, stuLname, stuClassID) VALUES (?, ?, ?, ?)", connection, tx);
                insertCmd.Parameters.AddWithValue("@p1", normalizedStuId);
                insertCmd.Parameters.AddWithValue("@p2", normalizedFirstName);
                insertCmd.Parameters.AddWithValue("@p3", normalizedLastName);
                insertCmd.Parameters.AddWithValue("@p4", normalizedClassCode);
                insertCmd.ExecuteNonQuery();
            }
            else if (HasInactiveTag(existingRemarks))
            {
                using var reactivateCmd = new OleDbCommand(
                    "UPDATE [Student] SET stuFname = ?, stuLname = ?, stuClassID = ?, stuRemarks = ? WHERE stuID = ?",
                    connection,
                    tx);
                reactivateCmd.Parameters.AddWithValue("@p1", normalizedFirstName);
                reactivateCmd.Parameters.AddWithValue("@p2", normalizedLastName);
                reactivateCmd.Parameters.AddWithValue("@p3", normalizedClassCode);
                var activeRemarks = RemoveInactiveTag(existingRemarks);
                reactivateCmd.Parameters.AddWithValue("@p4", string.IsNullOrWhiteSpace(activeRemarks) ? (object)DBNull.Value : activeRemarks);
                reactivateCmd.Parameters.AddWithValue("@p5", existingStuId);
                reactivateCmd.ExecuteNonQuery();
            }
            else
            {
                tx.Rollback();
                message = "חניך עם תעודת זהות זו כבר קיים במערכת.";
                return false;
            }

            tx.Commit();
            message = studentExists
                ? "החניך הופעל מחדש בהצלחה."
                : "החניך נוסף בהצלחה.";
            return true;
        }
        catch (Exception ex)
        {
            try { tx.Rollback(); } catch { }
            ErrorLogger.Log(ex, "TryAddManualStudent failed");
            message = $"הוספת חניך ידנית נכשלה: {ex.Message}";
            return false;
        }
    }

    public DataTable GetBookCopiesByBookId(string bkId, bool includeScrapped = false)
    {
        var dt = new DataTable();
        var cleanId = (bkId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanId))
        {
            dt.Columns.Add("Barcode", typeof(string));
            dt.Columns.Add("CopyReturnStatus", typeof(string));
            dt.Columns.Add("CopyCondition", typeof(string));
            dt.Columns.Add("BorrowerName", typeof(string));
            dt.Columns.Add("IsBorrowerIssue", typeof(bool));
            dt.Columns.Add("StatusDisplay", typeof(string));
            return dt;
        }

        try
        {
            using var connection = CreateOpenConnection();
            var copyLinkColumn = ResolveBookCopyLinkColumn(connection);
            var escapedCopyLinkColumn = EscapeIdentifier(copyLinkColumn);

            var conditionFilter = includeScrapped
                ? string.Empty
                : "AND (bc.CopyCondition IS NULL OR Trim(CStr(bc.CopyCondition)) NOT IN ('98','99'))";

            var sql = $@"SELECT bc.Barcode,
                          bc.CopyReturnStatus,
                          bc.CopyCondition
                   FROM [BookCopy] bc
                   WHERE bc.{escapedCopyLinkColumn} = ?
                     {conditionFilter}
                   ORDER BY bc.Barcode";

            using (var copiesCmd = new OleDbCommand(sql, connection))
            {
                copiesCmd.Parameters.AddWithValue("@p1", cleanId);
                using var adapter = new OleDbDataAdapter(copiesCmd);
                adapter.Fill(dt);
            }

            var studentNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var studentsCmd = new OleDbCommand("SELECT stuID, stuFname, stuLname FROM [Student]", connection))
            using (var reader = studentsCmd.ExecuteReader())
            {
                while (reader is not null && reader.Read())
                {
                    var rawStuId = reader["stuID"]?.ToString() ?? string.Empty;
                    var stuId = rawStuId.Trim();
                    var normalizedStuId = NormalizeStuId(rawStuId);
                    if (string.IsNullOrWhiteSpace(stuId))
                    {
                        continue;
                    }

                    var first = reader["stuFname"]?.ToString()?.Trim() ?? string.Empty;
                    var last = reader["stuLname"]?.ToString()?.Trim() ?? string.Empty;
                    var fullName = $"{first} {last}".Trim();
                    var displayName = string.IsNullOrWhiteSpace(fullName) ? stuId : fullName;
                    studentNames[stuId] = displayName;
                    if (!string.IsNullOrWhiteSpace(normalizedStuId))
                    {
                        studentNames[normalizedStuId] = displayName;
                    }
                }
            }

            var copyBarcodes = new HashSet<string>(
                dt.Rows.Cast<DataRow>()
                    .Select(r => r["Barcode"]?.ToString()?.Trim())
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .Select(v => v!),
                StringComparer.OrdinalIgnoreCase);

            var activeBorrowersByBarcode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var invalidBorrowerStuIdByBarcode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var transCmd = new OleDbCommand("SELECT Barcode, stuID FROM [Transaction] WHERE History = '0'", connection))
            using (var reader = transCmd.ExecuteReader())
            {
                while (reader is not null && reader.Read())
                {
                    var barcode = reader["Barcode"]?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(barcode) || !copyBarcodes.Contains(barcode))
                    {
                        continue;
                    }

                    var stuId = reader["stuID"]?.ToString()?.Trim() ?? string.Empty;
                    var normalizedStuId = NormalizeStuId(stuId);
                    if (!string.IsNullOrWhiteSpace(stuId)
                        && (studentNames.TryGetValue(stuId, out var name)
                            || (!string.IsNullOrWhiteSpace(normalizedStuId)
                                && studentNames.TryGetValue(normalizedStuId, out name))))
                    {
                        if (!activeBorrowersByBarcode.ContainsKey(barcode))
                        {
                            activeBorrowersByBarcode[barcode] = name;
                        }
                    }
                    else if (!string.IsNullOrWhiteSpace(stuId))
                    {
                        invalidBorrowerStuIdByBarcode[barcode] = stuId;
                    }
                }
            }

            if (!dt.Columns.Contains("StatusDisplay"))
            {
                dt.Columns.Add("StatusDisplay", typeof(string));
            }

            if (!dt.Columns.Contains("IsBorrowerIssue"))
            {
                dt.Columns.Add("IsBorrowerIssue", typeof(bool));
            }

            if (!dt.Columns.Contains("BorrowerName"))
            {
                dt.Columns.Add("BorrowerName", typeof(string));
            }

            foreach (DataRow row in dt.Rows)
            {
                var copyCondition = (row["CopyCondition"]?.ToString() ?? string.Empty).Trim();
                var barcode = row["Barcode"]?.ToString()?.Trim() ?? string.Empty;

                string resolvedBorrowerName = string.Empty;
                var hasValidBorrower = false;
                if (!string.IsNullOrWhiteSpace(barcode)
                    && activeBorrowersByBarcode.TryGetValue(barcode, out var borrowerNameFromMap))
                {
                    resolvedBorrowerName = borrowerNameFromMap;
                    hasValidBorrower = true;
                }

                string invalidStuId = string.Empty;
                var hasInvalidBorrower = false;
                if (!hasValidBorrower && !string.IsNullOrWhiteSpace(barcode)
                                   && invalidBorrowerStuIdByBarcode.TryGetValue(barcode, out var invalidStuIdFromMap))
                {
                    invalidStuId = invalidStuIdFromMap;
                    hasInvalidBorrower = true;
                }
                var isBorrowerIssue = hasInvalidBorrower;
                row["CopyReturnStatus"] = (hasValidBorrower || hasInvalidBorrower) ? "1" : "0";

                row["IsBorrowerIssue"] = isBorrowerIssue;
                row["BorrowerName"] = hasValidBorrower
                    ? resolvedBorrowerName
                    : (hasInvalidBorrower ? $"חניך לא במצבת ({invalidStuId})" : "-");
                row["StatusDisplay"] = (copyCondition == "98" || copyCondition == "99")
                    ? "גרוט"
                    : ((hasValidBorrower || hasInvalidBorrower) ? "מושאל" : "זמין");
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"GetBookCopiesByBookId failed for bkID '{cleanId}'");
            throw;
        }

        return dt;
    }

    public DataTable GetBookCopies(string bkId)
    {
        return GetBookCopiesByBookId(bkId, includeScrapped: true);
    }

    public DataTable GetActiveBorrowersForBook(string bkId)
    {
        var dt = new DataTable();
        dt.Columns.Add("StudentId", typeof(string));
        dt.Columns.Add("StudentName", typeof(string));
        dt.Columns.Add("ClassName", typeof(string));
        dt.Columns.Add("Barcode", typeof(string));
        dt.Columns.Add("LoanDate", typeof(DateTime));

        var cleanId = (bkId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanId))
        {
            return dt;
        }

        try
        {
            using var connection = CreateOpenConnection();
            var copyLinkColumn = ResolveBookCopyLinkColumn(connection);
            var escapedCopyLinkColumn = EscapeIdentifier(copyLinkColumn);

            var sql = $@"SELECT t.stuID AS StudentId,
                                Trim(s.stuFname & ' ' & s.stuLname) AS StudentName,
                                IIF(IsNull(c.ClassName), Trim(CStr(IIF(s.stuClassID IS NULL, '', s.stuClassID))), c.ClassName) AS ClassName,
                                t.Barcode,
                                t.DateOfTrans AS LoanDate
                         FROM (([Transaction] AS t
                         LEFT JOIN [Student] AS s ON t.stuID = s.stuID)
                         LEFT JOIN [Classes] AS c ON s.stuClassID = c.ClassID)
                         INNER JOIN [BookCopy] AS bc ON t.Barcode = bc.Barcode
                         WHERE t.History = '0' AND Trim(CStr(bc.{escapedCopyLinkColumn})) = ?
                         ORDER BY IIF(IsNull(c.ClassName), Trim(CStr(IIF(s.stuClassID IS NULL, '', s.stuClassID))), c.ClassName), s.stuLname, s.stuFname, t.DateOfTrans";

            using var cmd = new OleDbCommand(sql, connection);
            cmd.Parameters.AddWithValue("@pBookId", cleanId);
            using var adapter = new OleDbDataAdapter(cmd);
            adapter.Fill(dt);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"GetActiveBorrowersForBook failed for bkID '{cleanId}'");
            throw;
        }

        return dt;
    }

    public bool TryMarkBookCopyAsScrapped(string barcode, out string message)
    {
        try
        {
            return TryMarkBookCopyAsScrappedCore(barcode, out message);
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"Scrap copy failed for barcode '{barcode}'");
            message = "גריטת העותק נכשלה.";
            return false;
        }
    }

    private static bool HasActiveLoan(OleDbConnection connection, OleDbTransaction tx, string barcode)
    {
        using var cmd = new OleDbCommand("SELECT COUNT(*) FROM [Transaction] WHERE Barcode = ? AND History = '0'", connection, tx);
        cmd.Parameters.AddWithValue("@p1", barcode);
        return Convert.ToInt32(cmd.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture) > 0;
    }

    private string? GetActiveLoanBorrowerId(OleDbConnection connection, OleDbTransaction? tx, string barcode)
    {
        using var cmd = new OleDbCommand("SELECT TOP 1 stuID FROM [Transaction] WHERE Barcode = ? AND History = '0'", connection, tx);
        cmd.Parameters.AddWithValue("@p1", barcode);
        return cmd.ExecuteScalar()?.ToString()?.Trim();
    }

    private string BuildAlreadyBorrowedMessage(string? borrowerStuId, string? currentStuId)
    {
        if (!string.IsNullOrWhiteSpace(borrowerStuId) && !string.IsNullOrWhiteSpace(currentStuId)
            && BuildStudentIdCandidates(currentStuId).Contains(borrowerStuId, StringComparer.OrdinalIgnoreCase))
        {
            return "עותק זה כבר מושאל לחניך זה.";
        }

        return "העותק כבר מושאל לחניך אחר.";
    }

    private bool TryMarkBookCopyAsScrappedCore(string barcode, out string message)
    {
        var cleanBarcode = (barcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanBarcode))
        {
            message = "ברקוד עותק חסר.";
            return false;
        }

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();

        using var copyLookupCmd = new OleDbCommand(@"SELECT TOP 1 bc.bkID, bc.CopyCondition
                                                     FROM [BookCopy] bc
                                                     WHERE Trim(CStr(bc.Barcode)) = ?", connection, tx);
        copyLookupCmd.Parameters.Add("@p1", OleDbType.VarWChar).Value = cleanBarcode;

        string? copyBookId = null;
        string currentCondition;
        using (var reader = copyLookupCmd.ExecuteReader())
        {
            if (reader is null || !reader.Read())
            {
                reader?.Close();
                tx.Rollback();
                message = "העותק לא נמצא במלאי.";
                return false;
            }

            copyBookId = reader["bkID"]?.ToString()?.Trim();
            currentCondition = reader["CopyCondition"]?.ToString()?.Trim() ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(copyBookId))
        {
            tx.Rollback();
            message = "לא ניתן לזהות את הכותר עבור העותק.";
            return false;
        }

        if (currentCondition == "98" || currentCondition == "99")
        {
            tx.Rollback();
            message = "העותק כבר מסומן כבלאי.";
            return false;
        }

        if (HasActiveLoan(connection, tx, cleanBarcode))
        {
            tx.Rollback();
            message = "לא ניתן לגרוט או להכריז כאבוד עותק שנמצא כרגע בהשאלה פעילה. יש לבצע לו החזרה תחילה.";
            return false;
        }

        var closedLoans = 0;
        var historyIndexTracker = 0;
        CloseActiveLoanWithHistory(connection, tx, cleanBarcode, "ההשאלה נסגרה עקב גריטת עותק", ref closedLoans, ref historyIndexTracker);

        using var updateCmd = new OleDbCommand(@"UPDATE [BookCopy]
                                                SET CopyCondition = '98', CopyReturnStatus = '0'
                                                WHERE bkID = ? AND Trim(CStr(Barcode)) = ?", connection, tx);
        updateCmd.Parameters.Add("@p1", OleDbType.VarWChar).Value = copyBookId;
        updateCmd.Parameters.Add("@p2", OleDbType.VarWChar).Value = cleanBarcode;

        var affected = updateCmd.ExecuteNonQuery();
        if (affected <= 0)
        {
            tx.Rollback();
            message = "לא ניתן היה לעדכן את סטטוס העותק.";
            return false;
        }

        tx.Commit();
        message = "העותק סומן כבלאי והוסר מהמלאי הפעיל.";
        return true;
    }

    public bool TryMarkBookTitleAsScrapped(string bkId, out string message)
    {
        var cleanBookId = (bkId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanBookId))
        {
            message = "יש לבחור כותר לגריטה.";
            return false;
        }

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();

        try
        {
            var titleBarcodes = new List<string>();
            using (var bcCmd = new OleDbCommand("SELECT Barcode FROM [BookCopy] WHERE bkID = ?", connection, tx))
            {
                bcCmd.Parameters.AddWithValue("@p1", cleanBookId);
                using var bcReader = bcCmd.ExecuteReader();
                while (bcReader is not null && bcReader.Read())
                {
                    var bc = bcReader["Barcode"]?.ToString()?.Trim() ?? string.Empty;
                    if (bc.Length > 0) titleBarcodes.Add(bc);
                }
            }

            var closedLoans = 0;
            var historyIndexTracker = 0;
            foreach (var bc in titleBarcodes)
            {
                CloseActiveLoanWithHistory(connection, tx, bc, "ההשאלה נסגרה עקב גריטת כותר", ref closedLoans, ref historyIndexTracker);
            }

            using var scrapTitleCmd = new OleDbCommand("UPDATE [BookCopy] SET CopyCondition = '98', CopyReturnStatus = '0' WHERE bkID = ?", connection, tx);
            scrapTitleCmd.Parameters.AddWithValue("@p1", cleanBookId);
            var affected = scrapTitleCmd.ExecuteNonQuery();
            if (affected <= 0)
            {
                tx.Rollback();
                message = "לא נמצאו עותקים לגריטה עבור הכותר שנבחר.";
                return false;
            }

            tx.Commit();
            message = $"הכותר {cleanBookId} וכל עותקיו סומנו כבלאי.";
            return true;
        }
        catch (Exception ex)
        {
            tx.Rollback();
            ErrorLogger.Log(ex, $"Scrap title failed for bkID '{cleanBookId}'");
            message = "גריטת הכותר נכשלה.";
            return false;
        }
    }

    public bool TryRestoreScrappedCopy(string barcode, out string message)
    {
        var cleanBarcode = (barcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanBarcode))
        {
            message = "ברקוד עותק חסר.";
            return false;
        }

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();

        try
        {
            using (var checkCmd = new OleDbCommand("SELECT TOP 1 CopyCondition, CopyReturnStatus FROM [BookCopy] WHERE Barcode = ?", connection, tx))
            {
                checkCmd.Parameters.AddWithValue("@p1", cleanBarcode);
                using var checkReader = checkCmd.ExecuteReader();
                if (checkReader is null || !checkReader.Read())
                {
                    checkReader?.Close();
                    tx.Rollback();
                    message = "העותק לא נמצא במלאי.";
                    return false;
                }

                var condition = checkReader["CopyCondition"]?.ToString()?.Trim() ?? string.Empty;
                if (condition != "98" && condition != "99")
                {
                    checkReader.Close();
                    tx.Rollback();
                    message = "לא ניתן לשחזר עותק שאינו מסומן כגרוט";
                    return false;
                }
            }

            var closedLoans = 0;
            var historyIndexTracker = 0;
            CloseActiveLoanWithHistory(connection, tx, cleanBarcode, "ההשאלה נסגרה עקב שחזור עותק מגריטה", ref closedLoans, ref historyIndexTracker);

            using var updateCmd = new OleDbCommand("UPDATE [BookCopy] SET CopyCondition = '2', CopyReturnStatus = '0' WHERE Barcode = ?", connection, tx);
            updateCmd.Parameters.AddWithValue("@p1", cleanBarcode);
            var affected = updateCmd.ExecuteNonQuery();
            if (affected <= 0)
            {
                tx.Rollback();
                message = "העותק לא נמצא במלאי.";
                return false;
            }

            tx.Commit();
            message = "העותק הוחזר לבעלות הספרייה כמלאי זמין.";
            return true;
        }
        catch (Exception ex)
        {
            tx.Rollback();
            ErrorLogger.Log(ex, $"Restore scrapped copy failed for barcode '{cleanBarcode}'");
            message = "שחזור העותק מגריטה נכשל.";
            return false;
        }
    }

    public DataTable GetPublishers()
    {
        var dt = new DataTable();
        using var connection = CreateOpenConnection();
        using var cmd = new OleDbCommand("SELECT PublishID, PublishName FROM [Publisher] ORDER BY PublishName", connection);
        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    public DataTable GetBookTypes()
    {
        var dt = new DataTable();
        using var connection = CreateOpenConnection();
        using var cmd = new OleDbCommand("SELECT TypeID, TypeName FROM [Type] ORDER BY TypeName", connection);
        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    private sealed record AuthorLookup(string TableName, string IdColumnName, string NameColumnName);

    private static string EscapeIdentifier(string identifier)
    {
        return $"[{identifier.Replace("]", "]]", StringComparison.Ordinal)}]";
    }

    private static string ResolveBookCopyLinkColumn(OleDbConnection connection)
    {
        var columnCandidates = new[] { "bkID", "BookID", "BookCode", "CatalogNumber" };
        var columnSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object?[] { null, null, "BookCopy", null });
        if (columnSchema is null)
        {
            return "bkID";
        }

        var availableColumns = new HashSet<string>(
            columnSchema.Rows
                .Cast<DataRow>()
                .Select(row => row["COLUMN_NAME"]?.ToString())
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name!),
            StringComparer.OrdinalIgnoreCase);

        return columnCandidates.FirstOrDefault(availableColumns.Contains) ?? "bkID";
    }

    private static AuthorLookup? TryResolveAuthorLookup(OleDbConnection connection)
    {
        var tableCandidates = new[] { "Author", "Authors", "tblAuthor", "tblAuthors" };
        var idColumnCandidates = new[] { "AuthorID", "AuthorId", "AuthorCode", "Code", "ID" };
        var nameColumnCandidates = new[] { "AuthorName", "Name", "AuthorDesc", "Description", "Author" };

        var tableSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object?[] { null, null, null, "TABLE" });
        if (tableSchema is null)
        {
            return null;
        }

        var tableNames = tableSchema.Rows
            .Cast<DataRow>()
            .Select(row => row["TABLE_NAME"]?.ToString())
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Select(name => name!)
            .ToArray();

        var resolvedTableName = tableCandidates
            .Select(candidate => tableNames.FirstOrDefault(name => string.Equals(name, candidate, StringComparison.OrdinalIgnoreCase)))
            .FirstOrDefault(name => !string.IsNullOrWhiteSpace(name));

        if (string.IsNullOrWhiteSpace(resolvedTableName))
        {
            return null;
        }

        var columnSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object?[] { null, null, resolvedTableName, null });
        if (columnSchema is null)
        {
            return null;
        }

        var columns = new HashSet<string>(
            columnSchema.Rows
                .Cast<DataRow>()
                .Select(row => row["COLUMN_NAME"]?.ToString())
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name!),
            StringComparer.OrdinalIgnoreCase);

        var idColumn = idColumnCandidates.FirstOrDefault(columns.Contains);
        var nameColumn = nameColumnCandidates.FirstOrDefault(columns.Contains);
        if (string.IsNullOrWhiteSpace(idColumn) || string.IsNullOrWhiteSpace(nameColumn))
        {
            return null;
        }

        return new AuthorLookup(resolvedTableName, idColumn, nameColumn);
    }

    private static string NormalizeAuthorName(string? authorName)
    {
        return string.Join(" ", (authorName ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    public int GetOrCreateAuthorId(string authorName)
    {
        using var connection = CreateOpenConnection();
        return GetOrCreateAuthorId(authorName, connection);
    }

    private static int GetOrCreateAuthorId(string authorName, OleDbConnection connection)
    {
        var normalizedAuthorName = NormalizeAuthorName(authorName);
        if (string.IsNullOrWhiteSpace(normalizedAuthorName) || normalizedAuthorName == "-")
        {
            return 0;
        }

        var authorLookup = TryResolveAuthorLookup(connection);
        if (authorLookup is null)
        {
            return int.TryParse(normalizedAuthorName, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedId)
                ? parsedId
                : 0;
        }

        var authorTable = EscapeIdentifier(authorLookup.TableName);
        var authorIdColumn = EscapeIdentifier(authorLookup.IdColumnName);
        var authorNameColumn = EscapeIdentifier(authorLookup.NameColumnName);

        using (var findCmd = new OleDbCommand(
                    $"SELECT {authorIdColumn}, {authorNameColumn} FROM {authorTable}",
                    connection))
        using (var reader = findCmd.ExecuteReader())
        {
            while (reader is not null && reader.Read())
            {
                var existingName = NormalizeAuthorName(reader[authorLookup.NameColumnName]?.ToString());
                if (!string.Equals(existingName, normalizedAuthorName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return Convert.ToInt32(reader[authorLookup.IdColumnName], CultureInfo.InvariantCulture);
            }
        }

        using (var insertCmd = new OleDbCommand(
                    $"INSERT INTO {authorTable} ({authorNameColumn}) VALUES (?)",
                    connection))
        {
            insertCmd.Parameters.AddWithValue("@p1", normalizedAuthorName);
            insertCmd.ExecuteNonQuery();
        }

        using var identityCmd = new OleDbCommand("SELECT @@IDENTITY", connection);
        var identityValue = identityCmd.ExecuteScalar();
        return identityValue is null || identityValue == DBNull.Value
            ? 0
            : Convert.ToInt32(identityValue, CultureInfo.InvariantCulture);
    }

    public DataTable SearchBooks(string query)
    {
        var dt = new DataTable();
        var q = (query ?? string.Empty).Trim();
        try
        {
            using var connection = CreateOpenConnection();

            var authorLookup = TryResolveAuthorLookup(connection);
            var authorDisplayExpression = "IIF(IsNull(b.bkAuthor) OR Trim(CStr(b.bkAuthor)) = '' OR Trim(CStr(b.bkAuthor)) = '0', '-', CStr(b.bkAuthor))";
            var booksWithDetailsFromClause = @"(([Book] b
                        LEFT JOIN [Subject] s ON b.bkSubjectID = s.SubjectID)
                        LEFT JOIN [Type] t ON b.bkTypeID = t.TypeID)";

            if (authorLookup is not null)
            {
                var authorTable = EscapeIdentifier(authorLookup.TableName);
                var authorIdColumn = EscapeIdentifier(authorLookup.IdColumnName);
                var authorNameColumn = EscapeIdentifier(authorLookup.NameColumnName);

                booksWithDetailsFromClause = $@"((([Book] b
                        LEFT JOIN {authorTable} a ON CStr(b.bkAuthor) = CStr(a.{authorIdColumn}))
                        LEFT JOIN [Subject] s ON b.bkSubjectID = s.SubjectID)
                        LEFT JOIN [Type] t ON b.bkTypeID = t.TypeID)";

                authorDisplayExpression = $@"IIF(IsNull(a.{authorNameColumn}) OR Trim(CStr(a.{authorNameColumn})) = '',
                                IIF(IsNull(b.bkAuthor) OR Trim(CStr(b.bkAuthor)) = '' OR Trim(CStr(b.bkAuthor)) = '0', '-', CStr(b.bkAuthor)),
                                CStr(a.{authorNameColumn}))";
            }

            var booksSql = $@"SELECT b.bkID,
                                   b.bkName,
                                   {authorDisplayExpression} AS bkAuthor,
                                   b.bkPrice,
                                   b.bkPublish,
                                   s.SubjectName,
                                   t.TypeName
                            FROM {booksWithDetailsFromClause}
                            ORDER BY b.bkName";

            using (var booksCmd = new OleDbCommand(booksSql, connection))
            using (var adapter = new OleDbDataAdapter(booksCmd))
            {
                adapter.Fill(dt);
            }

            if (!dt.Columns.Contains("TotalCopies"))
            {
                dt.Columns.Add("TotalCopies", typeof(int));
            }

            if (!dt.Columns.Contains("BorrowedCopies"))
            {
                dt.Columns.Add("BorrowedCopies", typeof(int));
            }

            if (!dt.Columns.Contains("AvailableCopies"))
            {
                dt.Columns.Add("AvailableCopies", typeof(int));
            }

            var totalCopiesByBookId = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var activeBarcodeToBookId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var copiesCmd = new OleDbCommand("SELECT bkID, Barcode, CopyCondition FROM [BookCopy]", connection))
            using (var reader = copiesCmd.ExecuteReader())
            {
                while (reader is not null && reader.Read())
                {
                    var bookId = reader["bkID"]?.ToString()?.Trim() ?? string.Empty;
                    var barcode = reader["Barcode"]?.ToString()?.Trim() ?? string.Empty;
                    var copyCondition = reader["CopyCondition"]?.ToString()?.Trim() ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(bookId) || string.IsNullOrWhiteSpace(barcode)
                        || copyCondition == "98" || copyCondition == "99")
                    {
                        continue;
                    }

                    totalCopiesByBookId.TryGetValue(bookId, out var currentTotal);
                    totalCopiesByBookId[bookId] = currentTotal + 1;
                    activeBarcodeToBookId[barcode] = bookId;
                }
            }

            var existingStudentIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var studentsCmd = new OleDbCommand("SELECT stuID FROM [Student]", connection))
            using (var reader = studentsCmd.ExecuteReader())
            {
                while (reader is not null && reader.Read())
                {
                    var rawStuId = reader["stuID"]?.ToString() ?? string.Empty;
                    var stuId = rawStuId.Trim();
                    var normalizedStuId = NormalizeStuId(rawStuId);
                    if (!string.IsNullOrWhiteSpace(stuId))
                    {
                        existingStudentIds.Add(stuId);
                        if (!string.IsNullOrWhiteSpace(normalizedStuId))
                        {
                            existingStudentIds.Add(normalizedStuId);
                        }
                    }
                }
            }

            var borrowedBarcodesByBookId = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            using (var transCmd = new OleDbCommand("SELECT bkID, Barcode FROM [Transaction] WHERE History = '0'", connection))
            using (var reader = transCmd.ExecuteReader())
            {
                while (reader is not null && reader.Read())
                {
                    var barcode = reader["Barcode"]?.ToString()?.Trim() ?? string.Empty;
                    var transBookId = reader["bkID"]?.ToString()?.Trim() ?? string.Empty;

                    var bookId = string.Empty;
                    if (!string.IsNullOrWhiteSpace(barcode) && activeBarcodeToBookId.TryGetValue(barcode, out var mappedBookId))
                    {
                        bookId = mappedBookId;
                    }
                    else
                    {
                        bookId = transBookId;
                    }

                    if (string.IsNullOrWhiteSpace(bookId)
                        || string.IsNullOrWhiteSpace(barcode)
                        || !totalCopiesByBookId.ContainsKey(bookId))
                    {
                        continue;
                    }

                    if (!borrowedBarcodesByBookId.TryGetValue(bookId, out var barcodeSet))
                    {
                        barcodeSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        borrowedBarcodesByBookId[bookId] = barcodeSet;
                    }

                    barcodeSet.Add(barcode);
                }
            }

            HashSet<string> barcodeMatchedBookIds;
            if (!string.IsNullOrWhiteSpace(q))
            {
                barcodeMatchedBookIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using var barcodeSearchCmd = new OleDbCommand("SELECT DISTINCT bkID FROM [BookCopy] WHERE Barcode = ?", connection);
                barcodeSearchCmd.Parameters.AddWithValue("@p1", q);
                using var barcodeReader = barcodeSearchCmd.ExecuteReader();
                while (barcodeReader is not null && barcodeReader.Read())
                {
                    var bookId = barcodeReader["bkID"]?.ToString()?.Trim();
                    if (!string.IsNullOrWhiteSpace(bookId))
                    {
                        barcodeMatchedBookIds.Add(bookId);
                    }
                }
            }
            else
            {
                barcodeMatchedBookIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            var rowsToRemove = new List<DataRow>();
            foreach (DataRow row in dt.Rows)
            {
                var bookId = row["bkID"]?.ToString()?.Trim() ?? string.Empty;
                totalCopiesByBookId.TryGetValue(bookId, out var totalCopies);
                borrowedBarcodesByBookId.TryGetValue(bookId, out var borrowedSet);
                var borrowedCopies = borrowedSet?.Count ?? 0;
                var availableCopies = Math.Max(0, totalCopies - borrowedCopies);

                row["TotalCopies"] = totalCopies;
                row["BorrowedCopies"] = borrowedCopies;
                row["AvailableCopies"] = availableCopies;

                if (string.IsNullOrWhiteSpace(q))
                {
                    continue;
                }

                var bookName = row["bkName"]?.ToString() ?? string.Empty;
                var bookAuthor = row["bkAuthor"]?.ToString() ?? string.Empty;
                var isMatch = bookName.Contains(q, StringComparison.OrdinalIgnoreCase)
                            || bookAuthor.Contains(q, StringComparison.OrdinalIgnoreCase)
                            || string.Equals(bookId, q, StringComparison.OrdinalIgnoreCase)
                            || barcodeMatchedBookIds.Contains(bookId);

                if (!isMatch)
                {
                    rowsToRemove.Add(row);
                }
            }

            foreach (var row in rowsToRemove)
            {
                dt.Rows.Remove(row);
            }

            dt.AcceptChanges();
            return dt;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "SearchBooks failed");
            throw;
        }
    }

    private static List<string> NormalizeClassList(List<string>? classes)
    {
        return (classes ?? new List<string>())
            .Select(c => (c ?? string.Empty).Trim())
            .Where(c => c.Length > 0 && c != "הכל")
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }

    public List<string> GetPaymentSchoolYears()
    {
        var years = new HashSet<string>(StringComparer.Ordinal);
        using var connection = CreateFreshPaymentReadConnection();
        using var cmd = new OleDbCommand("SELECT DISTINCT [Year] FROM [Payment] WHERE [Year] IS NOT NULL", connection);
        using var reader = cmd.ExecuteReader();
        while (reader != null && reader.Read())
        {
            var normalized = NormalizeSchoolYear(reader[0]?.ToString());
            if (normalized.Length > 0 && TryParseHebrewYear(normalized, out _))
            {
                years.Add(normalized);
            }
        }

        return years
            .OrderByDescending(y => TryParseHebrewYear(y, out var hy) ? hy : 0)
            .ToList();
    }

    public DataTable GetStudentsWithNoLoansThisYear(List<string> classes, string currentSchoolYear)
    {
        var dt = new DataTable();
        using var connection = CreateOpenConnection();

        var (from, _) = GetSchoolYearDateRange(currentSchoolYear);
        var normalizedClasses = NormalizeClassList(classes);

        var conditions = new List<string>
        {
            "(s.stuClassID IS NULL OR CStr(s.stuClassID) <> '9999')",
            "NOT EXISTS (SELECT 1 FROM [Transaction] t WHERE t.stuID = s.stuID AND t.DateOfTrans >= ?)",
            "NOT EXISTS (SELECT 1 FROM [History] h WHERE h.stuID = s.stuID AND h.DateOfTrans >= ?)"
        };

        if (normalizedClasses.Count > 0)
        {
            conditions.Add("Trim(CStr(IIF(s.stuClassID IS NULL, '', s.stuClassID))) IN (" + string.Join(", ", normalizedClasses.Select(_ => "?")) + ")");
        }

        var sql = $@"SELECT s.stuID,
                            Trim(s.stuFname & ' ' & s.stuLname) AS StudentName,
                            IIF(IsNull(c.ClassName), Trim(CStr(IIF(s.stuClassID IS NULL, '', s.stuClassID))), c.ClassName) AS ClassName,
                            s.stuPhone
                     FROM [Student] s
                     LEFT JOIN [Classes] c ON s.stuClassID = c.ClassID
                     WHERE {string.Join(" AND ", conditions)}
                     ORDER BY c.ClassName, s.stuLname, s.stuFname";

        using var cmd = new OleDbCommand(sql, connection);
        for (var i = 0; i < 2; i++)
        {
            cmd.Parameters.Add("@pStart" + i, OleDbType.Date).Value = from;
        }

        for (var i = 0; i < normalizedClasses.Count; i++)
        {
            cmd.Parameters.AddWithValue("@pClass" + i, normalizedClasses[i]);
        }

        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    public DataTable GetOverdueLoans(List<string>? classes = null)
    {
        var dt = new DataTable();
        using var connection = CreateOpenConnection();

        var conditions = new List<string>
        {
            "t.History = '0'",
            "t.DateToReturn < ?",
            "(s.stuClassID IS NULL OR CStr(s.stuClassID) <> '9999')"
        };

        var normalizedClasses = NormalizeClassList(classes);
        var hasClassFilter = normalizedClasses.Count > 0;

        if (hasClassFilter)
        {
            conditions.Add("Trim(CStr(IIF(s.stuClassID IS NULL, '', s.stuClassID))) IN (" + string.Join(", ", normalizedClasses.Select(_ => "?")) + ")");
        }

        var whereClause = "WHERE " + string.Join(" AND ", conditions);

        var sql = $@"SELECT t.Barcode, 
                            IIF(b.bkName IS NULL, 'ללא כותרת', b.bkName) AS BookTitle,
                            t.stuID, 
                            Trim(s.stuFname & ' ' & s.stuLname) AS StudentName,
                            c.ClassName,
                            s.stuPhone,
                            s.stuCellular,
                            t.DateOfTrans AS LoanDate,
                            t.DateToReturn AS DueDate,
                            DATEDIFF('d', t.DateToReturn, ?) AS DaysOverdue
                     FROM ((([Transaction] t
                     LEFT JOIN [Student] s ON t.stuID = s.stuID)
                     LEFT JOIN [Classes] c ON s.stuClassID = c.ClassID)
                     LEFT JOIN [BookCopy] bc ON t.Barcode = bc.Barcode)
                     LEFT JOIN [Book] b ON bc.bkID = b.bkID
                     {whereClause}
                     ORDER BY t.DateToReturn ASC, c.ClassName, s.stuLname, s.stuFname";

        using var cmd = new OleDbCommand(sql, connection);
        var today = DateTime.Today;
        cmd.Parameters.Add("@pToday1", OleDbType.Date).Value = today;
        cmd.Parameters.Add("@pDueDate", OleDbType.Date).Value = today;

        for (var i = 0; i < normalizedClasses.Count; i++)
        {
            cmd.Parameters.AddWithValue("@pClass" + i, normalizedClasses[i]);
        }

        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    public DataTable GetActiveLoansReport(string? classFilter = null)
    {
        var dt = new DataTable();
        using var connection = CreateOpenConnection();

        var conditions = new List<string>
        {
            "t.History = '0'"
        };

        var normalizedClassFilter = (classFilter ?? string.Empty).Trim();
        var hasClassFilter = !string.IsNullOrWhiteSpace(normalizedClassFilter) && normalizedClassFilter != "הכל";

        if (hasClassFilter)
        {
            conditions.Add("Trim(CStr(IIF(s.stuClassID IS NULL, '', s.stuClassID))) = ?");
        }

        var whereClause = "WHERE " + string.Join(" AND ", conditions);

        var sql = $@"SELECT t.Barcode, 
                            IIF(b.bkName IS NULL, 'ללא כותרת', b.bkName) AS BookTitle,
                            t.stuID, 
                            Trim(s.stuFname & ' ' & s.stuLname) AS StudentName,
                            IIF(IsNull(c.ClassName), Trim(CStr(IIF(s.stuClassID IS NULL, '', s.stuClassID))), c.ClassName) AS ClassName,
                            s.stuPhone,
                            t.DateOfTrans AS LoanDate,
                            t.DateToReturn AS DueDate
                     FROM ((([Transaction] t
                     LEFT JOIN [Student] s ON t.stuID = s.stuID)
                     LEFT JOIN [Classes] c ON s.stuClassID = c.ClassID)
                     LEFT JOIN [BookCopy] bc ON t.Barcode = bc.Barcode)
                     LEFT JOIN [Book] b ON bc.bkID = b.bkID
                     {whereClause}
                     ORDER BY c.ClassName, s.stuLname, s.stuFname, t.DateOfTrans";

        using var cmd = new OleDbCommand(sql, connection);

        if (hasClassFilter)
        {
            cmd.Parameters.AddWithValue("@pClass", normalizedClassFilter);
        }

        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    public DbHealthReport RunDatabaseHealthCheck()
    {
        var warnings = new List<string>();
        var sampleBarcodes = new List<string>();

        using var connection = CreateOpenConnection();

        int totalBooks = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [Book]", connection).ExecuteScalar() ?? 0);
        int totalCopies = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [BookCopy]", connection).ExecuteScalar() ?? 0);
        int totalStudents = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [Student]", connection).ExecuteScalar() ?? 0);
        int activeTrans = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [Transaction] WHERE History = '0'", connection).ExecuteScalar() ?? 0);
        int historyTrans = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [History]", connection).ExecuteScalar() ?? 0);
        int totalPayments = Convert.ToInt32(new OleDbCommand("SELECT COUNT(*) FROM [Payment]", connection).ExecuteScalar() ?? 0);

        var orphanSql = "SELECT COUNT(*) FROM [BookCopy] bc WHERE NOT EXISTS (SELECT 1 FROM [Book] b WHERE b.bkID = bc.bkID)";
        int orphanCopies = Convert.ToInt32(new OleDbCommand(orphanSql, connection).ExecuteScalar() ?? 0);
        if (orphanCopies > 0)
        {
            warnings.Add($"נמצאו {orphanCopies} עותקים ב-BookCopy שמזהה הספר שלהם (bkID) אינו קיים בטבלת Book.");
        }

        var dupSql = "SELECT COUNT(*) FROM (SELECT Barcode FROM [BookCopy] GROUP BY Barcode HAVING COUNT(*) > 1)";
        int duplicateBarcodes = Convert.ToInt32(new OleDbCommand(dupSql, connection).ExecuteScalar() ?? 0);
        if (duplicateBarcodes > 0)
        {
            warnings.Add($"נמצאו {duplicateBarcodes} ברקודים כפולים בטבלת BookCopy!");
        }

        var brokenTransBarcodeSql = @"SELECT COUNT(*) FROM [Transaction] t 
                                     WHERE t.History = '0' 
                                       AND NOT EXISTS (SELECT 1 FROM [BookCopy] bc WHERE bc.Barcode = t.Barcode)";
        int brokenTransBarcodes = Convert.ToInt32(new OleDbCommand(brokenTransBarcodeSql, connection).ExecuteScalar() ?? 0);
        if (brokenTransBarcodes > 0)
        {
            warnings.Add($"נמצאו {brokenTransBarcodes} השאלות פעילות שהברקוד שלהן אינו קיים ב-BookCopy.");
        }

        var brokenTransStudentSql = @"SELECT COUNT(*) FROM [Transaction] t 
                                     WHERE t.History = '0' 
                                       AND NOT EXISTS (SELECT 1 FROM [Student] s WHERE s.stuID = t.stuID)";
        int brokenTransStudents = Convert.ToInt32(new OleDbCommand(brokenTransStudentSql, connection).ExecuteScalar() ?? 0);
        if (brokenTransStudents > 0)
        {
            warnings.Add($"נמצאו {brokenTransStudents} השאלות פעילות שהחניך שלהן אינו קיים בטבלת Student.");
        }

        using (var sampleCmd = new OleDbCommand("SELECT TOP 10 Barcode FROM [BookCopy] ORDER BY EnterDate DESC, Barcode", connection))
        using (var reader = sampleCmd.ExecuteReader())
        {
            while (reader is not null && reader.Read())
            {
                var bc = reader["Barcode"]?.ToString();
                if (!string.IsNullOrWhiteSpace(bc)) sampleBarcodes.Add(bc);
            }
        }

        return new DbHealthReport(
            totalStudents,
            totalBooks,
            totalCopies,
            activeTrans,
            historyTrans,
            totalPayments,
            orphanCopies,
            duplicateBarcodes
        )
        {
            Warnings = warnings,
            SampleBarcodes = sampleBarcodes
        };
    }

    public DataTable GetBrokenActiveLoans()
    {
        var dt = new DataTable();
        using var connection = CreateOpenConnection();
        var sql = @"SELECT t.Barcode, t.stuID, t.bkID,
						   IIF(IsNull(s.stuID), 'חניך לא ידוע', (s.stuFname & '') & ' ' & (s.stuLname & '')) AS StudentName,
						   c.ClassName,
						   t.DateOfTrans AS LoanDate,
						   t.DateToReturn AS DueDate,
						   IIF(IsNull(bc.Barcode), 'ברקוד אינו קיים במלאי (BookCopy)',
							   IIF(IsNull(s.stuID), 'חניך אינו קיים',
								   'עותק גרוט רשום כמושאל (שאריות סלקוד)')) AS Reason,
						   IIF(IsNull(bc.Barcode), 1, 0) AS MissingBarcode,
						   IIF(IsNull(bc.Barcode), 1, IIF(Trim(bc.CopyCondition & '') IN ('98','99'), 1, 0)) AS AutoClean
					FROM ((([Transaction] t
					LEFT JOIN [BookCopy] bc ON t.Barcode = bc.Barcode)
					LEFT JOIN [Student] s ON t.stuID = s.stuID)
					LEFT JOIN [Classes] c ON s.stuClassID = c.ClassID)
					WHERE t.History = '0'
					  AND (IsNull(bc.Barcode) OR IsNull(s.stuID) OR Trim(bc.CopyCondition & '') IN ('98','99'))";

        using var cmd = new OleDbCommand(sql, connection);
        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    private static int CleanOrCloseOrphanedLoanCore(OleDbConnection connection, OleDbTransaction tx, string barcode, ref int historyIndexTracker)
    {
        var closed = 0;
        CloseActiveLoanWithHistory(connection, tx, barcode, "ניקוי השאלה יתומה אוטומטי", ref closed, ref historyIndexTracker);

        using (var resetCopyCmd = new OleDbCommand("UPDATE [BookCopy] SET CopyReturnStatus = '0' WHERE Barcode = ? AND (CopyCondition IS NULL OR Trim(CStr(CopyCondition)) NOT IN ('98','99'))", connection, tx))
        {
            resetCopyCmd.Parameters.AddWithValue("@p1", barcode);
            resetCopyCmd.ExecuteNonQuery();
        }

        return closed;
    }

    public bool CleanOrCloseOrphanedLoan(string barcode, out string message)
    {
        barcode = (barcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(barcode))
        {
            message = "לא ניתן לזהות ברקוד להשאלה החריגה.";
            return false;
        }

        try
        {
            using var connection = CreateOpenConnection();
            using var tx = connection.BeginTransaction();
            try
            {
                var historyIndexTracker = 0;
                var closed = CleanOrCloseOrphanedLoanCore(connection, tx, barcode, ref historyIndexTracker);
                tx.Commit();
                message = closed > 0 ? "ההשאלה החריגה נוקתה בהצלחה." : "ההשאלה לא נמצאה או שכבר נסגרה.";
                return closed > 0;
            }
            catch
            {
                try { tx.Rollback(); } catch { }
                throw;
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, $"CleanOrCloseOrphanedLoan failed for barcode {barcode}");
            message = $"שגיאה בניקוי ההשאלה: {ex.Message}";
            return false;
        }
    }

    public bool CleanAllAnomalousLoans(out int cleaned, out string message)
    {
        cleaned = 0;
        try
        {
            using var connection = CreateOpenConnection();
            using var tx = connection.BeginTransaction();
            try
            {
                var barcodes = new List<string>();
                using (var findCmd = new OleDbCommand(
                    "SELECT t.Barcode FROM [Transaction] t WHERE t.History = '0' AND (NOT EXISTS (SELECT 1 FROM [BookCopy] bc WHERE bc.Barcode = t.Barcode) OR EXISTS (SELECT 1 FROM [BookCopy] bc2 WHERE bc2.Barcode = t.Barcode AND Trim(bc2.CopyCondition & '') IN ('98','99')))", connection, tx))
                using (var reader = findCmd.ExecuteReader())
                {
                    while (reader is not null && reader.Read())
                    {
                        var barcode = reader["Barcode"]?.ToString()?.Trim() ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(barcode))
                        {
                            barcodes.Add(barcode);
                        }
                    }
                }

                var historyIndexTracker = 0;
                foreach (var barcode in barcodes.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    cleaned += CleanOrCloseOrphanedLoanCore(connection, tx, barcode, ref historyIndexTracker);
                }

                tx.Commit();
                message = $"נוקו {cleaned} השאלות חריגות (ברקוד חסר או עותק גרוט).";
                return true;
            }
            catch
            {
                try { tx.Rollback(); } catch { }
                throw;
            }
        }
        catch (Exception ex)
        {
            cleaned = 0;
            ErrorLogger.Log(ex, "CleanAllAnomalousLoans failed");
            message = $"שגיאה בניקוי ההשאלות: {ex.Message}";
            return false;
        }
    }

    private static bool HasInactiveTag(string? remarks)
    {
        if (string.IsNullOrWhiteSpace(remarks))
        {
            return false;
        }

        return remarks.IndexOf("[INACTIVE]", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string RemoveInactiveTag(string? remarks)
    {
        var value = remarks ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var cleaned = value.Replace("[INACTIVE]", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();

        while (cleaned.Contains("  ", StringComparison.Ordinal))
        {
            cleaned = cleaned.Replace("  ", " ", StringComparison.Ordinal);
        }

        return cleaned;
    }

    private static string AddInactiveTag(string? remarks)
    {
        var cleaned = RemoveInactiveTag(remarks);
        return string.IsNullOrWhiteSpace(cleaned)
            ? "[INACTIVE]"
            : $"[INACTIVE] {cleaned}";
    }

    private static string NormalizeStudentIdForSync(string rawId)
    {
        var value = (rawId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var digits = new string(value.Where(char.IsDigit).ToArray());
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

    private static bool IsValidIsraeliId(string id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return false;
        }

        var normalized = new string(id.Trim().Where(char.IsDigit).ToArray());
        if (normalized.Length > 9)
        {
            normalized = normalized[^9..];
        }

        normalized = normalized.PadLeft(9, '0');
        if (normalized.Length != 9)
        {
            return false;
        }

        var sum = 0;
        for (var i = 0; i < 9; i++)
        {
            var digit = normalized[i] - '0';
            var step = digit * ((i % 2) + 1);
            if (step > 9)
            {
                step -= 9;
            }

            sum += step;
        }

        return sum % 10 == 0;
    }

    public record PaymentImportPreview(
        int TotalInFile,
        int ReadyToMarkPaid,
        int AlreadyPaidCount,
        int ExemptCount,
        int UnknownStudentsCount,
        List<string> UnknownIds
    );

    public PaymentImportPreview PreviewPaymentImport(IEnumerable<string> rawIds)
    {
        var cleanIds = (rawIds ?? Enumerable.Empty<string>())
            .Select(NormalizeStudentIdForSync)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        int alreadyPaid = 0;
        int exempt = 0;
        int readyToMark = 0;
        var unknownIds = new List<string>();
        var currentYear = GetCurrentSchoolYear();

        using var connection = CreateOpenConnection();

        var lookup = LoadPaymentLookup(connection, null, currentYear);

        foreach (var id in cleanIds)
        {
            var shortId = id.TrimStart('0');
            if (string.IsNullOrWhiteSpace(shortId))
            {
                shortId = id;
            }

            if (lookup.TryResolve(id, shortId, out _, out var criteria, out var hasPayment, out var hasExemptionPayment))
            {
                var isExempt = hasExemptionPayment || criteria == "פטור" || criteria == "1";

                if (isExempt)
                {
                    exempt++;
                }
                else if (hasPayment)
                {
                    alreadyPaid++;
                }
                else
                {
                    readyToMark++;
                }
            }
            else
            {
                unknownIds.Add(id);
            }
        }

        return new PaymentImportPreview(cleanIds.Count, readyToMark, alreadyPaid, exempt, unknownIds.Count, unknownIds);
    }

    private sealed class PaymentLookup
    {
        public Dictionary<string, string?> Students { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string?> ClassIds { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> Phones { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> Cities { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> Cores { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> Paid { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> Exemptions { get; } = new(StringComparer.OrdinalIgnoreCase);

        public bool TryResolve(string id, string shortId, out string? actualStuId, out string? criteria, out bool hasPayment, out bool hasExemption)
        {
            actualStuId = null;
            criteria = null;
            hasPayment = false;
            hasExemption = false;

            string? key = Students.ContainsKey(id) ? id : Students.ContainsKey(shortId) ? shortId : null;
            if (key is null && Cores.TryGetValue(id.TrimStart('0'), out var coreKey))
            {
                key = coreKey;
            }
            if (key is null)
            {
                return false;
            }

            actualStuId = key;
            criteria = Students[key];
            hasPayment = Paid.Contains(key) || Paid.Contains(id) || Paid.Contains(shortId);
            hasExemption = Exemptions.Contains(key) || Exemptions.Contains(id) || Exemptions.Contains(shortId);
            return true;
        }
    }

    private static PaymentLookup LoadPaymentLookup(OleDbConnection connection, OleDbTransaction? tx, object currentYear)
    {
        var lookup = new PaymentLookup();

        using (var studentCmd = new OleDbCommand(
            "SELECT stuID, stuCriteriaStatus, stuClassID, stuPhone, stuCity FROM [Student]", connection, tx))
        using (var reader = studentCmd.ExecuteReader())
        {
            while (reader.Read())
            {
                var stuId = reader["stuID"]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(stuId))
                {
                    lookup.Students.TryAdd(stuId, reader["stuCriteriaStatus"]?.ToString()?.Trim());
                    lookup.ClassIds.TryAdd(stuId, reader["stuClassID"]?.ToString()?.Trim());
                    lookup.Phones.TryAdd(stuId, reader["stuPhone"]?.ToString()?.Trim() ?? string.Empty);
                    lookup.Cities.TryAdd(stuId, reader["stuCity"]?.ToString()?.Trim() ?? string.Empty);
                    var coreId = stuId.TrimStart('0');
                    if (coreId.Length > 0)
                    {
                        lookup.Cores.TryAdd(coreId, stuId);
                    }
                }
            }
        }

        using (var paymentCmd = new OleDbCommand("SELECT stuID, Amount, ReciptNum, [Year] FROM [Payment]", connection, tx))
        {
            using var reader = paymentCmd.ExecuteReader();
            while (reader.Read())
            {
                var rowYear = reader["Year"] == DBNull.Value ? string.Empty : reader["Year"]?.ToString() ?? string.Empty;
                if (!IsCurrentSchoolYearValue(rowYear, currentYear?.ToString() ?? string.Empty))
                {
                    continue;
                }

                var stuId = reader["stuID"]?.ToString()?.Trim();
                if (string.IsNullOrEmpty(stuId))
                {
                    continue;
                }

                var amount = reader["Amount"] is DBNull ? 0m : Convert.ToDecimal(reader["Amount"], CultureInfo.InvariantCulture);
                if (amount > 0)
                {
                    lookup.Paid.Add(stuId);
                }
                else if (amount == 0 && string.Equals(reader["ReciptNum"]?.ToString(), "פטור מאושר", StringComparison.Ordinal))
                {
                    lookup.Exemptions.Add(stuId);
                }
            }
        }

        return lookup;
    }

    public (int NewlyMarkedPaid, int Skipped) ExecutePaymentImport(IEnumerable<string> rawIds, IReadOnlyDictionary<string, (string FirstName, string LastName, string City, string Phone)>? newStudentDetails = null)
    {
        var cleanIds = (rawIds ?? Enumerable.Empty<string>())
            .Select(NormalizeStudentIdForSync)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        int newlyMarked = 0;
        int skipped = 0;
        var currentYear = GetCurrentSchoolYear();

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();

        try
        {
            var lookup = LoadPaymentLookup(connection, tx, currentYear);

            using (var checkClassCmd = new OleDbCommand("SELECT COUNT(*) FROM [Classes] WHERE ClassID = '0000'", connection, tx))
            {
                if (Convert.ToInt32(checkClassCmd.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture) == 0)
                {
                    using var addClassCmd = new OleDbCommand("INSERT INTO [Classes] (ClassID, ClassName) VALUES ('0000', ?)", connection, tx);
                    addClassCmd.Parameters.AddWithValue("@p1", "אחר");
                    addClassCmd.ExecuteNonQuery();
                }
            }

            foreach (var id in cleanIds)
            {
                var shortId = id.TrimStart('0');
                if (string.IsNullOrWhiteSpace(shortId))
                {
                    shortId = id;
                }

                var found = lookup.TryResolve(id, shortId, out var actualStuId, out var criteria, out var hasPaid, out var hasExemptionPayment);
                var isExempt = found && (hasExemptionPayment || criteria == "פטור" || criteria == "1");

                if (found && actualStuId is not null && newStudentDetails is not null && newStudentDetails.TryGetValue(id, out var existingDetails))
                {
                    var newPhone = (existingDetails.Phone ?? string.Empty).Trim();
                    var newCity = (existingDetails.City ?? string.Empty).Trim();
                    var currentPhone = lookup.Phones.TryGetValue(actualStuId, out var cp) ? cp : string.Empty;
                    var currentCity = lookup.Cities.TryGetValue(actualStuId, out var cc) ? cc : string.Empty;
                    var phoneToSet = string.IsNullOrWhiteSpace(currentPhone) && newPhone.Length > 0 ? newPhone : currentPhone;
                    var cityToSet = string.IsNullOrWhiteSpace(currentCity) && newCity.Length > 0 ? newCity : currentCity;
                    if (!string.Equals(phoneToSet, currentPhone, StringComparison.Ordinal) || !string.Equals(cityToSet, currentCity, StringComparison.Ordinal))
                    {
                        using var completeCmd = new OleDbCommand("UPDATE [Student] SET stuPhone = ?, stuCity = ? WHERE stuID = ?", connection, tx);
                        completeCmd.Parameters.AddWithValue("@p1", string.IsNullOrWhiteSpace(phoneToSet) ? (object)DBNull.Value : phoneToSet);
                        completeCmd.Parameters.AddWithValue("@p2", string.IsNullOrWhiteSpace(cityToSet) ? (object)DBNull.Value : cityToSet);
                        completeCmd.Parameters.AddWithValue("@p3", actualStuId);
                        completeCmd.ExecuteNonQuery();
                        lookup.Phones[actualStuId] = phoneToSet ?? string.Empty;
                        lookup.Cities[actualStuId] = cityToSet ?? string.Empty;
                    }
                }

                if (!found)
                {
                    var hasDetails = newStudentDetails is not null && newStudentDetails.TryGetValue(id, out var details);
                    var detailValues = hasDetails ? newStudentDetails![id] : default;
                    var newFirstName = string.IsNullOrWhiteSpace(detailValues.FirstName) ? "חניך חדש" : detailValues.FirstName.Trim();
                    var newLastName = string.IsNullOrWhiteSpace(detailValues.LastName) ? "חניך חדש" : detailValues.LastName.Trim();
                    using var insertStudentCmd = new OleDbCommand(
                        "INSERT INTO [Student] (stuID, stuFname, stuLname, stuClassID, stuPhone, stuCity) VALUES (?, ?, ?, ?, ?, ?)", connection, tx);
                    insertStudentCmd.Parameters.AddWithValue("@p1", id);
                    insertStudentCmd.Parameters.AddWithValue("@p2", newFirstName);
                    insertStudentCmd.Parameters.AddWithValue("@p3", newLastName);
                    insertStudentCmd.Parameters.AddWithValue("@p4", "0000");
                    insertStudentCmd.Parameters.AddWithValue("@p5", (detailValues.Phone ?? string.Empty).Trim());
                    insertStudentCmd.Parameters.AddWithValue("@p6", (detailValues.City ?? string.Empty).Trim());
                    insertStudentCmd.ExecuteNonQuery();
                    actualStuId = id;
                    lookup.Students[id] = null;
                    hasPaid = false;
                    isExempt = false;
                }

                if (found
                    && actualStuId is not null
                    && lookup.ClassIds.TryGetValue(actualStuId, out var currentClassId)
                    && string.Equals(currentClassId, "9999", StringComparison.Ordinal))
                {
                    using var reviveCmd = new OleDbCommand(
                        "UPDATE [Student] SET stuClassID = '0000' WHERE stuID = ? AND stuClassID = '9999'", connection, tx);
                    reviveCmd.Parameters.AddWithValue("@p1", actualStuId);
                    reviveCmd.ExecuteNonQuery();
                    lookup.ClassIds[actualStuId] = "0000";
                }

                if (hasPaid || isExempt)
                {
                    skipped++;
                    continue;
                }

                var payKey = $"{actualStuId}_{currentYear}_{DateTime.Now:yyyyMMddHHmmssfff}";
                using (var insertCmd = new OleDbCommand(
                    "INSERT INTO [Payment] (PayKeyID, stuID, Amount, [Year], PayTypeID) VALUES (?, ?, 1, ?, '1')", connection, tx))
                {
                    insertCmd.Parameters.AddWithValue("@p1", payKey);
                    insertCmd.Parameters.AddWithValue("@p2", actualStuId);
                    insertCmd.Parameters.AddWithValue("@p3", currentYear);
                    insertCmd.ExecuteNonQuery();
                }

                lookup.Paid.Add(actualStuId!);
                newlyMarked++;
            }

            tx.Commit();
        }
        catch (Exception ex)
        {
            try { tx.Rollback(); } catch { }
            ErrorLogger.Log(ex, "ExecutePaymentImport batch failed");
            throw;
        }

        return (newlyMarked, skipped);
    }

    public DataTable SearchStudents(string? searchText, string? classFilter = null)
    {
        var result = new DataTable();
        using var connection = CreateOpenConnection();

        bool hasClassesTable = false;
        try
        {
            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (schema != null)
            {
                foreach (DataRow row in schema.Rows)
                {
                    if (string.Equals(row["TABLE_NAME"]?.ToString()?.Trim(), "Classes", StringComparison.OrdinalIgnoreCase))
                    {
                        hasClassesTable = true;
                        break;
                    }
                }
            }
        }
        catch { }

        var selectSql = hasClassesTable
            ? @"SELECT s.stuID, s.stuFname, s.stuLname, s.stuClassID, s.stuPhone, s.stuCity, s.stuRemarks, c.ClassName
                FROM [Student] AS s
                LEFT JOIN [Classes] AS c ON s.stuClassID = c.ClassID"
            : @"SELECT s.stuID, s.stuFname, s.stuLname, s.stuClassID, s.stuPhone, s.stuCity, s.stuRemarks, s.stuClassID AS ClassName
                FROM [Student] AS s";

        var whereParts = new List<string>();
        var normalizedSearch = (searchText ?? string.Empty).Trim();
        var normalizedClass = (classFilter ?? string.Empty).Trim();

        using var cmd = new OleDbCommand();
        cmd.Connection = connection;

        if (!string.IsNullOrWhiteSpace(normalizedSearch))
        {
            whereParts.Add("(s.stuID LIKE ? OR s.stuFname LIKE ? OR s.stuLname LIKE ?)");
            var like = $"%{normalizedSearch}%";
            cmd.Parameters.AddWithValue("@p1", like);
            cmd.Parameters.AddWithValue("@p2", like);
            cmd.Parameters.AddWithValue("@p3", like);
        }

        if (!string.IsNullOrWhiteSpace(normalizedClass) &&
            !string.Equals(normalizedClass, "הכל", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(normalizedClass, "כל הכיתות", StringComparison.OrdinalIgnoreCase))
        {
            whereParts.Add("Trim(CStr(IIF(s.stuClassID IS NULL, '', s.stuClassID))) = ?");
            cmd.Parameters.AddWithValue("@p4", normalizedClass);
        }

        var sql = selectSql;
        if (whereParts.Count > 0)
        {
            sql += " WHERE " + string.Join(" AND ", whereParts);
        }

        sql += " ORDER BY s.stuLname, s.stuFname, s.stuID";
        cmd.CommandText = sql;

        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(result);

        if (!result.Columns.Contains("PaymentStatus"))
        {
            result.Columns.Add("PaymentStatus", typeof(string));
        }

        if (!result.Columns.Contains("ActiveLoansCount"))
        {
            result.Columns.Add("ActiveLoansCount", typeof(int));
        }

        foreach (DataRow row in result.Rows)
        {
            row["PaymentStatus"] = "לא שולם";
            row["ActiveLoansCount"] = 0;
        }

        var studentIds = result.AsEnumerable()
            .Select(row => row["stuID"]?.ToString()?.Trim())
            .Where(stuId => !string.IsNullOrWhiteSpace(stuId))
            .Select(stuId => stuId!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (studentIds.Length == 0)
        {
            return result;
        }

        var currentYear = GetCurrentSchoolYear();
        var exemptStudentIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var paidStudentIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var paidCmd = new OleDbCommand("SELECT stuID, [Year], Amount, ReciptNum FROM [Payment]", connection);
            using var paidReader = paidCmd.ExecuteReader();
            while (paidReader != null && paidReader.Read())
            {
                var rawStuId = paidReader["stuID"]?.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(rawStuId))
                {
                    continue;
                }

                var yearValue = paidReader["Year"] == DBNull.Value ? string.Empty : paidReader["Year"]?.ToString()?.Trim() ?? string.Empty;
                if (!IsCurrentSchoolYearValue(yearValue, currentYear))
                {
                    continue;
                }

                var amount = ReadAmountValue(paidReader["Amount"]);

                var reciptNum = paidReader["ReciptNum"] == DBNull.Value
                    ? string.Empty
                    : paidReader["ReciptNum"]?.ToString()?.Trim() ?? string.Empty;

                if (IsExemptionPaymentRow(amount, reciptNum))
                {
                    foreach (var candidate in BuildStudentIdCandidates(rawStuId))
                    {
                        exemptStudentIds.Add(candidate);
                    }

                    continue;
                }

                if (amount > 0)
                {
                    foreach (var candidate in BuildStudentIdCandidates(rawStuId))
                    {
                        paidStudentIds.Add(candidate);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "SearchStudents: failed loading paid students map");
        }

        var activeLoansByStudentId = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var activeLoansCmd = new OleDbCommand("SELECT stuID, COUNT(*) AS LoansCount FROM [Transaction] WHERE History = '0' GROUP BY stuID", connection);
            using var activeLoansReader = activeLoansCmd.ExecuteReader();
            while (activeLoansReader != null && activeLoansReader.Read())
            {
                var rawStuId = activeLoansReader["stuID"]?.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(rawStuId))
                {
                    continue;
                }

                var loansCount = Convert.ToInt32(activeLoansReader["LoansCount"] ?? 0, CultureInfo.InvariantCulture);
                foreach (var candidate in BuildStudentIdCandidates(rawStuId))
                {
                    activeLoansByStudentId[candidate] = loansCount;
                }
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "SearchStudents: failed loading active loans map");
        }

        foreach (DataRow row in result.Rows)
        {
            var stuId = row["stuID"]?.ToString()?.Trim() ?? string.Empty;
            var idCandidates = BuildStudentIdCandidates(stuId);

            var isExempt = idCandidates.Any(exemptStudentIds.Contains);
            var hasPaid = idCandidates.Any(paidStudentIds.Contains);
            row["PaymentStatus"] = isExempt ? "פטור" : hasPaid ? "שולם" : "לא שולם";

            var activeLoans = 0;
            foreach (var candidate in idCandidates)
            {
                if (activeLoansByStudentId.TryGetValue(candidate, out activeLoans))
                {
                    break;
                }
            }

            row["ActiveLoansCount"] = activeLoans;
        }

        return result;
    }

    public DataTable GetRecentCopiesForPrinting(int limit = 50)
    {
        var dt = new DataTable();
        using var connection = CreateOpenConnection();
        var sql = $@"SELECT TOP {limit} bc.Barcode, b.bkID, IIF(b.bkName IS NULL, 'ללא כותרת', b.bkName) AS bkName, bc.CopyReturnStatus, bc.CopyCondition
                     FROM [BookCopy] bc
                     LEFT JOIN [Book] b ON bc.bkID = b.bkID
                     ORDER BY bc.EnterDate DESC";
        using var cmd = new OleDbCommand(sql, connection);
        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    public DataTable SearchCopiesForPrinting(string query)
    {
        var dt = new DataTable();
        var cleanQ = (query ?? string.Empty).Trim();
        using var connection = CreateOpenConnection();
        var sql = @"SELECT bc.Barcode, b.bkID, IIF(b.bkName IS NULL, 'ללא כותרת', b.bkName) AS bkName, bc.CopyReturnStatus, bc.CopyCondition
                     FROM [BookCopy] bc
                     LEFT JOIN [Book] b ON bc.bkID = b.bkID
                     WHERE bc.Barcode LIKE ? OR b.bkName LIKE ? OR b.bkID LIKE ?
                     ORDER BY bc.Barcode";
        using var cmd = new OleDbCommand(sql, connection);
        var like = $"%{cleanQ}%";
        cmd.Parameters.AddWithValue("@p1", like);
        cmd.Parameters.AddWithValue("@p2", like);
        cmd.Parameters.AddWithValue("@p3", like);
        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    public DataTable GetCopiesForBarcodePrinting(string? bkId = null)
    {
        var dt = new DataTable();
        using var connection = CreateOpenConnection();
        var sql = @"SELECT bc.Barcode, b.bkID, IIF(b.bkName IS NULL, 'ללא כותרת', b.bkName) AS bkName, bc.CopyReturnStatus, bc.CopyCondition
                     FROM [BookCopy] bc
                     LEFT JOIN [Book] b ON bc.bkID = b.bkID";

        if (!string.IsNullOrWhiteSpace(bkId))
        {
            sql += " WHERE b.bkID = ?";
        }
        sql += " ORDER BY bc.Barcode";

        using var cmd = new OleDbCommand(sql, connection);
        if (!string.IsNullOrWhiteSpace(bkId))
        {
            cmd.Parameters.AddWithValue("@p1", bkId.Trim());
        }
        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    public DataTable GetCopiesForBarcodePrinting(IEnumerable<string> barcodes)
    {
        var dt = new DataTable();
        var list = (barcodes ?? Array.Empty<string>())
            .Select(b => (b ?? string.Empty).Trim())
            .Where(b => !string.IsNullOrWhiteSpace(b))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        using var connection = CreateOpenConnection();
        if (list.Count == 0)
        {
            dt.Columns.Add("Barcode", typeof(string));
            dt.Columns.Add("bkID", typeof(string));
            dt.Columns.Add("bkName", typeof(string));
            dt.Columns.Add("CopyReturnStatus", typeof(string));
            dt.Columns.Add("CopyCondition", typeof(string));
            return dt;
        }

        var placeholders = string.Join(",", list.Select(_ => "?"));
        var sql = $@"SELECT bc.Barcode, b.bkID, IIF(b.bkName IS NULL, 'ללא כותרת', b.bkName) AS bkName, bc.CopyReturnStatus, bc.CopyCondition
                     FROM [BookCopy] bc
                     LEFT JOIN [Book] b ON bc.bkID = b.bkID
                     WHERE bc.Barcode IN ({placeholders})
                     ORDER BY bc.Barcode";

        using var cmd = new OleDbCommand(sql, connection);
        foreach (var bc in list)
        {
            cmd.Parameters.AddWithValue("@p", bc);
        }

        using var adapter = new OleDbDataAdapter(cmd);
        adapter.Fill(dt);
        return dt;
    }

    public DataTable GetCopiesForBarcodePrintingByBookId(string bkId)
    {
        return GetCopiesForBarcodePrinting(bkId);
    }

    public bool TryAddBook(string bkId, string bkName, string bkAuthor, string bkPrice, string bkPublish, string subjectIdStr, string typeIdStr, out string message)
    {
        int.TryParse(subjectIdStr, out var subjectId);
        int.TryParse(typeIdStr, out var typeId);
        return TryAddBook(bkId, bkName, bkAuthor, bkPrice, bkPublish, subjectId, typeId, out message);
    }

    public bool TryAddBook(string bkId, string bkName, string bkAuthor, string bkPrice, string bkPublish, int subjectId, int typeId, out string message)
    {
        var cleanId = (bkId ?? string.Empty).Trim();
        var cleanName = (bkName ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanId) || string.IsNullOrWhiteSpace(cleanName))
        {
            message = "מזהה ספר ושם ספר הם שדות חובה.";
            return false;
        }

        using var connection = CreateOpenConnection();
        try
        {
            using var cmd = new OleDbCommand("INSERT INTO [Book] (bkID, bkName, bkAuthor, bkPrice, bkPublish, bkSubjectID, bkTypeID, bkEnterDate) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", connection);
            cmd.Parameters.AddWithValue("@p1", cleanId);
            cmd.Parameters.AddWithValue("@p2", cleanName);
            cmd.Parameters.AddWithValue("@p3", bkAuthor ?? string.Empty);
            cmd.Parameters.AddWithValue("@p4", bkPrice ?? "0");
            cmd.Parameters.AddWithValue("@p5", bkPublish ?? string.Empty);
            cmd.Parameters.AddWithValue("@p6", subjectId);
            cmd.Parameters.AddWithValue("@p7", typeId);
            cmd.Parameters.Add("@p8", OleDbType.Date).Value = DateTime.Now;
            cmd.ExecuteNonQuery();

            message = "הספר נוסף בהצלחה.";
            return true;
        }
        catch (Exception ex)
        {
            message = $"שגיאה בהוספת הספר: {ex.Message}";
            return false;
        }
    }

    public bool UpdateBook(string bkId, string bkName, string bkAuthor, string bkPrice, string bkPublish, int subjectId, int typeId, out string message)
    {
        var cleanId = (bkId ?? string.Empty).Trim();
        var cleanName = (bkName ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanId))
        {
            message = "מזהה ספר חסר.";
            return false;
        }

        using var connection = CreateOpenConnection();
        try
        {
            using var cmd = new OleDbCommand("UPDATE [Book] SET bkName = ?, bkAuthor = ?, bkPrice = ?, bkPublish = ?, bkSubjectID = ?, bkTypeID = ? WHERE bkID = ?", connection);
            cmd.Parameters.AddWithValue("@p1", cleanName);
            cmd.Parameters.AddWithValue("@p2", bkAuthor ?? string.Empty);
            cmd.Parameters.AddWithValue("@p3", bkPrice ?? "0");
            cmd.Parameters.AddWithValue("@p4", bkPublish ?? string.Empty);
            cmd.Parameters.AddWithValue("@p5", subjectId);
            cmd.Parameters.AddWithValue("@p6", typeId);
            cmd.Parameters.AddWithValue("@p7", cleanId);
            cmd.ExecuteNonQuery();

            message = "הספר עודכן בהצלחה.";
            return true;
        }
        catch (Exception ex)
        {
            message = $"שגיאה בעדכון הספר: {ex.Message}";
            return false;
        }
    }

    public bool TryAddBookCopy(string bkId, string barcode, string condition, string returnStatus)
    {
        return TryAddBookCopy(bkId, barcode, out _);
    }

    public bool TryAddBookCopy(string bkId, string barcode, string location, out string message)
    {
        return TryAddBookCopy(bkId, barcode, out message);
    }

    public bool TryAddBookCopy(string bkId, string barcode, out string message)
    {
        var cleanBkId = (bkId ?? string.Empty).Trim();
        var cleanBarcode = (barcode ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanBkId) || string.IsNullOrWhiteSpace(cleanBarcode))
        {
            message = "מזהה ספר או ברקוד חסרים.";
            return false;
        }

        using var connection = CreateOpenConnection();
        try
        {
            using var cmd = new OleDbCommand("INSERT INTO [BookCopy] (bkID, Barcode, CopyReturnStatus, CopyCondition, EnterDate) VALUES (?, ?, '0', '1', ?)", connection);
            cmd.Parameters.AddWithValue("@p1", cleanBkId);
            cmd.Parameters.AddWithValue("@p2", cleanBarcode);
            cmd.Parameters.Add("@p3", OleDbType.Date).Value = DateTime.Now;
            cmd.ExecuteNonQuery();

            message = "העותק נוסף בהצלחה.";
            return true;
        }
        catch (Exception ex)
        {
            message = $"הוספת העותק נכשלה: {ex.Message}";
            return false;
        }
    }

    public bool TryCreateCopiesForBook(string bkId, int count, string initialCondition, out string message)
    {
        return TryCreateCopiesForBook(bkId, count, out _, out message);
    }

    public bool TryCreateCopiesForBook(string bkId, int count, out string message)
    {
        return TryCreateCopiesForBook(bkId, count, out _, out message);
    }

    public bool TryCreateCopiesForBook(string bkId, int count, out List<string> createdBarcodes, out string message)
    {
        createdBarcodes = [];
        message = string.Empty;
        var cleanBkId = (bkId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleanBkId) || count <= 0)
        {
            message = "נתונים שגויים ליצירת עותקים.";
            return false;
        }

        using var connection = CreateOpenConnection();
        using var tx = connection.BeginTransaction();
        try
        {
            var maxSuffix = 0;
            using (var maxCmd = new OleDbCommand("SELECT Barcode FROM [BookCopy] WHERE bkID = ?", connection, tx))
            {
                maxCmd.Parameters.AddWithValue("@p1", cleanBkId);
                using var reader = maxCmd.ExecuteReader();
                while (reader is not null && reader.Read())
                {
                    var existingBarcode = Convert.ToString(reader["Barcode"], CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(existingBarcode) || !existingBarcode.StartsWith(cleanBkId, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var suffix = existingBarcode.Substring(cleanBkId.Length);
                    if (suffix.Length == 0 || !int.TryParse(suffix, NumberStyles.None, CultureInfo.InvariantCulture, out var value))
                    {
                        continue;
                    }

                    if (value > maxSuffix)
                    {
                        maxSuffix = value;
                    }
                }
            }

            var created = new List<string>(count);
            var nextSuffix = maxSuffix;
            while (created.Count < count)
            {
                nextSuffix++;
                var barcode = $"{cleanBkId}{nextSuffix:D4}";
                if (BarcodeExists(connection, tx, barcode))
                {
                    continue;
                }

                using var cmd = new OleDbCommand("INSERT INTO [BookCopy] (bkID, Barcode, CopyReturnStatus, CopyCondition, EnterDate) VALUES (?, ?, '0', '1', ?)", connection, tx);
                cmd.Parameters.AddWithValue("@p1", cleanBkId);
                cmd.Parameters.AddWithValue("@p2", barcode);
                cmd.Parameters.Add("@p3", OleDbType.Date).Value = DateTime.Now;
                cmd.ExecuteNonQuery();
                created.Add(barcode);
            }

            tx.Commit();
            createdBarcodes = created;
            message = $"נוספו {count} עותקים בהצלחה.";
            return true;
        }
        catch (Exception ex)
        {
            tx.Rollback();
            message = $"יצירת העותקים נכשלה: {ex.Message}";
            return false;
        }
    }

    private static bool BarcodeExists(OleDbConnection connection, OleDbTransaction tx, string barcode)
    {
        using var cmd = new OleDbCommand("SELECT COUNT(*) FROM [BookCopy] WHERE Barcode = ?", connection, tx);
        cmd.Parameters.AddWithValue("@p1", barcode);
        return Convert.ToInt32(cmd.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture) > 0;
    }

    public StudentSyncResult SyncStudents(IEnumerable<StudentSyncRow> students, bool markMissingAsGraduates, bool dryRun = false)
    {
        var normalizedIncoming = (students ?? Enumerable.Empty<StudentSyncRow>())
            .Where(s => !string.IsNullOrWhiteSpace(s.StuId))
            .GroupBy(s => NormalizeStudentIdForSync(s.StuId), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.Last())
            .ToList();

        using var connection = CreateOpenConnection();
        OleDbTransaction? tx = null;
        if (!dryRun)
        {
            tx = connection.BeginTransaction();
        }

        try
        {
            var existingStudents = new Dictionary<string, (string StuId, string FirstName, string LastName, string ClassCode, string Phone, string City, string Remarks)>(StringComparer.OrdinalIgnoreCase);
            using (var readCmd = new OleDbCommand("SELECT stuID, stuFname, stuLname, stuClassID, stuPhone, stuCity, stuRemarks FROM [Student]", connection, tx))
            using (var reader = readCmd.ExecuteReader())
            {
                while (reader is not null && reader.Read())
                {
                    var stuId = Convert.ToString(reader["stuID"])?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(stuId))
                    {
                        continue;
                    }

                    var key = NormalizeStudentIdForSync(stuId);
                    existingStudents[key] = (
                        stuId,
                        Convert.ToString(reader["stuFname"])?.Trim() ?? string.Empty,
                        Convert.ToString(reader["stuLname"])?.Trim() ?? string.Empty,
                        NormalizeClassCode(Convert.ToString(reader["stuClassID"])?.Trim()),
                        Convert.ToString(reader["stuPhone"])?.Trim() ?? string.Empty,
                        Convert.ToString(reader["stuCity"])?.Trim() ?? string.Empty,
                        Convert.ToString(reader["stuRemarks"])?.Trim() ?? string.Empty);
                }
            }

            var newStudents = 0;
            var updatedStudents = 0;
            _ = markMissingAsGraduates;

            var classNameToCode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var maxClassCode = 0;
            var classIdToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var classCmd = new OleDbCommand("SELECT ClassID, ClassName FROM [Classes]", connection, tx))
            using (var classReader = classCmd.ExecuteReader())
            {
                while (classReader is not null && classReader.Read())
                {
                    var classId = Convert.ToString(classReader["ClassID"])?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(classId))
                    {
                        classIdToName[classId] = Convert.ToString(classReader["ClassName"])?.Trim() ?? string.Empty;
                    }

                    var className = (Convert.ToString(classReader["ClassName"])?.Trim() ?? string.Empty).Replace("'", "").Replace("\"", "").Replace("\u05F3", "").Replace("\u05F4", "");
                    if (!string.IsNullOrWhiteSpace(className) && !string.IsNullOrWhiteSpace(classId))
                    {
                        classNameToCode[className] = classId;
                    }

                    if (int.TryParse(classId, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numericId)
                        && numericId != 9999 && numericId > maxClassCode)
                    {
                        maxClassCode = numericId;
                    }
                }
            }

            string ResolveClassCode(string? rawClass)
            {
                var name = (rawClass ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(name) || IsOtherClassValue(name))
                {
                    return NormalizeClassCode(name);
                }

                if (classNameToCode.TryGetValue(name, out var code))
                {
                    return code;
                }

                maxClassCode++;
                code = maxClassCode.ToString(CultureInfo.InvariantCulture).PadLeft(4, '0');
                if (!dryRun)
                {
                    using var insertClassCmd = new OleDbCommand("INSERT INTO [Classes] (ClassID, ClassName) VALUES (?, ?)", connection, tx);
                    insertClassCmd.Parameters.AddWithValue("@p1", code);
                    insertClassCmd.Parameters.AddWithValue("@p2", name);
                    insertClassCmd.ExecuteNonQuery();
                }

                classNameToCode[name] = code;
                return code;
            }

            string EnsureClass(string classId, string className)
            {
                var id = (classId ?? string.Empty).Trim();
                var name = (className ?? string.Empty).Trim();
                if (!classIdToName.TryGetValue(id, out var currentName))
                {
                    if (!dryRun)
                    {
                        using var insertClassCmd = new OleDbCommand("INSERT INTO [Classes] (ClassID, ClassName) VALUES (?, ?)", connection, tx);
                        insertClassCmd.Parameters.AddWithValue("@p1", id);
                        insertClassCmd.Parameters.AddWithValue("@p2", name);
                        insertClassCmd.ExecuteNonQuery();
                    }

                    classIdToName[id] = name;
                }
                else if (!string.Equals(currentName, name, StringComparison.Ordinal))
                {
                    if (!dryRun)
                    {
                        using var updateClassCmd = new OleDbCommand("UPDATE [Classes] SET ClassName = ? WHERE ClassID = ?", connection, tx);
                        updateClassCmd.Parameters.AddWithValue("@p1", name);
                        updateClassCmd.Parameters.AddWithValue("@p2", id);
                        updateClassCmd.ExecuteNonQuery();
                    }

                    classIdToName[id] = name;
                }

                classNameToCode[name.Replace("'", string.Empty).Replace("\"", string.Empty)] = id;
                return id;
            }

            foreach (var (systemId, systemName) in new[] { ("0000", "אחר"), ("0001", "מורים"), ("9999", "לא פעיל") })
            {
                if (!classIdToName.ContainsKey(systemId))
                {
                    EnsureClass(systemId, systemName);
                }
            }

            foreach (var row in normalizedIncoming)
            {
                var incomingKey = NormalizeStudentIdForSync(row.StuId);
                var normalizedClass = string.IsNullOrWhiteSpace(row.ClassName)
                    ? ResolveClassCode(row.ClassCode)
                    : EnsureClass(NormalizeClassCode(row.ClassCode), row.ClassName);
                var normalizedPhone = (row.Phone ?? string.Empty).Trim();
                var normalizedCity = (row.City ?? string.Empty).Trim();
                var normalizedFirst = (row.FirstName ?? string.Empty).Trim();
                var normalizedLast = (row.LastName ?? string.Empty).Trim();

                if (!existingStudents.TryGetValue(incomingKey, out var existing))
                {
                    newStudents++;
                    if (!dryRun)
                    {
                        using var insertCmd = new OleDbCommand(
                            "INSERT INTO [Student] (stuID, stuFname, stuLname, stuClassID, stuPhone, stuCity) VALUES (?, ?, ?, ?, ?, ?)",
                            connection,
                            tx);
                        insertCmd.Parameters.AddWithValue("@p1", incomingKey);
                        insertCmd.Parameters.AddWithValue("@p2", string.IsNullOrWhiteSpace(normalizedFirst) ? (object)DBNull.Value : normalizedFirst);
                        insertCmd.Parameters.AddWithValue("@p3", string.IsNullOrWhiteSpace(normalizedLast) ? (object)DBNull.Value : normalizedLast);
                        insertCmd.Parameters.AddWithValue("@p4", normalizedClass);
                        insertCmd.Parameters.AddWithValue("@p5", string.IsNullOrWhiteSpace(normalizedPhone) ? (object)DBNull.Value : normalizedPhone);
                        insertCmd.Parameters.AddWithValue("@p6", string.IsNullOrWhiteSpace(normalizedCity) ? (object)DBNull.Value : normalizedCity);
                        insertCmd.ExecuteNonQuery();
                    }

                    continue;
                }

                updatedStudents++;
                if (!dryRun)
                {
                    using var updateCmd = new OleDbCommand(
                        "UPDATE [Student] SET stuClassID = ?, stuPhone = ?, stuCity = ?, stuRemarks = ? WHERE stuID = ?",
                        connection,
                        tx);
                    updateCmd.Parameters.AddWithValue("@p1", normalizedClass);
                    var phoneToSave = string.IsNullOrWhiteSpace(normalizedPhone) ? existing.Phone : normalizedPhone;
                    var cityToSave = string.IsNullOrWhiteSpace(normalizedCity) ? existing.City : normalizedCity;
                    updateCmd.Parameters.AddWithValue("@p2", string.IsNullOrWhiteSpace(phoneToSave) ? (object)DBNull.Value : phoneToSave);
                    updateCmd.Parameters.AddWithValue("@p3", string.IsNullOrWhiteSpace(cityToSave) ? (object)DBNull.Value : cityToSave);
                    updateCmd.Parameters.AddWithValue("@p4", string.IsNullOrWhiteSpace(existing.Remarks) ? (object)DBNull.Value : existing.Remarks);
                    updateCmd.Parameters.AddWithValue("@p5", existing.StuId);
                    updateCmd.ExecuteNonQuery();
                }
            }

            var markedGraduates = 0;
            var graduatesWithActiveLoans = 0;
            var activeStudentsBeforeSync = 0;
            if (markMissingAsGraduates)
            {
                var studentsWithActiveLoans = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using (var loansCmd = new OleDbCommand("SELECT DISTINCT stuID FROM [Transaction] WHERE History = '0'", connection, tx))
                using (var loansReader = loansCmd.ExecuteReader())
                {
                    while (loansReader is not null && loansReader.Read())
                    {
                        var loanStuId = Convert.ToString(loansReader["stuID"])?.Trim();
                        if (!string.IsNullOrWhiteSpace(loanStuId))
                        {
                            studentsWithActiveLoans.Add(NormalizeStudentIdForSync(loanStuId));
                        }
                    }
                }

                var incomingKeys = new HashSet<string>(
                    normalizedIncoming.Select(s => NormalizeStudentIdForSync(s.StuId)),
                    StringComparer.OrdinalIgnoreCase);
                var teacherCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "0001" };
                foreach (var pair in classNameToCode)
                {
                    if (pair.Key.Contains("מורים", StringComparison.OrdinalIgnoreCase))
                    {
                        teacherCodes.Add(pair.Value);
                    }
                }

                foreach (var pair in existingStudents)
                {
                    var currentClass = pair.Value.ClassCode;
                    if (currentClass != "9999" && !teacherCodes.Contains(currentClass))
                    {
                        activeStudentsBeforeSync++;
                    }

                    if (incomingKeys.Contains(pair.Key)
                        || currentClass == "9999"
                        || teacherCodes.Contains(currentClass))
                    {
                        continue;
                    }

                    markedGraduates++;
                    if (studentsWithActiveLoans.Contains(pair.Key))
                    {
                        graduatesWithActiveLoans++;
                    }
                    if (!dryRun)
                    {
                        using var gradCmd = new OleDbCommand("UPDATE [Student] SET stuClassID = '9999' WHERE stuID = ?", connection, tx);
                        gradCmd.Parameters.AddWithValue("@p1", pair.Value.StuId);
                        gradCmd.ExecuteNonQuery();
                    }
                }
            }

            if (!dryRun)
            {
                tx?.Commit();
            }

            return new StudentSyncResult(normalizedIncoming.Count, newStudents, updatedStudents, markedGraduates, graduatesWithActiveLoans, activeStudentsBeforeSync);
        }
        catch
        {
            if (!dryRun)
            {
                try { tx?.Rollback(); } catch { }
            }

            throw;
        }
    }

    private static DateTime GetEndOfSchoolYearReturnDate()
    {
        var today = DateTime.Today;
        bool afterCutoff = today.Month > 6 || (today.Month == 6 && today.Day >= 21);
        int returnYear = afterCutoff ? today.Year + 1 : today.Year;
        return new DateTime(returnYear, 6, 20);
    }

}
