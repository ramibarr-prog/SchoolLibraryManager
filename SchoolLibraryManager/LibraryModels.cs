using System.Globalization;

namespace SchoolLibraryManager;

public class Student
{
    public string StuId { get; set; } = string.Empty;
    public string FirstName { get; set; } = string.Empty;
    public string LastName { get; set; } = string.Empty;
    public string ClassCode { get; set; } = "0000";
    public string ClassName { get; set; } = string.Empty;
    public string? City { get; set; }
    public string? Phone { get; set; }
    public string? Notes { get; set; }

    public string FullName => $"{FirstName} {LastName}".Trim();
}

public class Payment
{
    public int PaymentId { get; set; }
    public string StuId { get; set; } = string.Empty;
    public string SchoolYear { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public string ReciptNum { get; set; } = string.Empty;
}

public class Book
{
    public int BookId { get; set; }
    public string Title { get; set; } = string.Empty;
    public string? Author { get; set; }
    public string? Subject { get; set; }
    public string? Publisher { get; set; }
    public decimal? Price { get; set; }
    public string? BookType { get; set; }
    public string? Isbn { get; set; }
}

public class BookCopy
{
    public string Barcode { get; set; } = string.Empty;
    public int BookId { get; set; }
    public DateTime EntryDate { get; set; } = DateTime.Now;
    public string Status { get; set; } = "זמין";
    public string? Location { get; set; }
}

public class LoanTransaction
{
    public int TransactionId { get; set; }
    public string StuId { get; set; } = string.Empty;
    public string Barcode { get; set; } = string.Empty;
    public DateTime DateBorrowed { get; set; } = DateTime.Now;
    public DateTime DateToReturn { get; set; } = DateTime.Now.AddDays(14);
    public DateTime? ActualReturnDate { get; set; }
    public string History { get; set; } = "0";

    public bool IsActive => History == "0";
}

public static class ClassFormatter
{
    public const string TeacherCode = "0001";
    public const string TeacherDisplay = "מורים";

    public static string ToDisplayName(string? className, string? classCode)
    {
        var name = (className ?? string.Empty).Trim();
        var code = (classCode ?? string.Empty).Trim();
        if (code == TeacherCode || name == TeacherCode || name.Contains(TeacherDisplay, StringComparison.Ordinal))
        {
            return TeacherDisplay;
        }

        return name;
    }

    private static readonly Dictionary<int, string> GradeToHebrew = new()
    {
        { 7, "ז'" },
        { 8, "ח'" },
        { 9, "ט'" },
        { 10, "י'" },
        { 11, "י\"א" },
        { 12, "י\"ב" }
    };

    private static readonly Dictionary<string, int> HebrewToGrade = GradeToHebrew
        .ToDictionary(pair => pair.Value, pair => pair.Key, StringComparer.OrdinalIgnoreCase);

    public static string ToDisplay(string? classCode)
    {
        if (string.IsNullOrWhiteSpace(classCode))
        {
            return "לא משויך";
        }

        var normalized = classCode.Trim();
        if (normalized == TeacherCode || normalized == TeacherDisplay)
        {
            return TeacherDisplay;
        }

        if (normalized == "9999")
        {
            return "בוגר / לא פעיל";
        }

        if (normalized == "0000")
        {
            return "לא משויך";
        }

        if (normalized.Length != 4 ||
            !int.TryParse(normalized[..2], NumberStyles.Integer, CultureInfo.InvariantCulture, out var grade) ||
            !int.TryParse(normalized[2..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var classNumber))
        {
            return normalized;
        }

        if (GradeToHebrew.TryGetValue(grade, out var hebrewGrade))
        {
            return $"{hebrewGrade} {classNumber}";
        }

        return $"{grade:00} {classNumber}";
    }

    public static string ToCode(string? displayClass)
    {
        if (string.IsNullOrWhiteSpace(displayClass))
        {
            return "0000";
        }

        var value = displayClass.Trim();
        if (value == "בוגר / לא פעיל")
        {
            return "9999";
        }

        if (value == "לא משויך")
        {
            return "0000";
        }

        if (value.Length == 4 && int.TryParse(value, out _))
        {
            return value;
        }

        var parts = value.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2 || !HebrewToGrade.TryGetValue(parts[0], out var grade) || !int.TryParse(parts[1], out var classNumber))
        {
            return "0000";
        }

        return BuildClassCode(grade, classNumber);
    }

    public static string BuildClassCode(int grade, int classNumber)
    {
        if (grade == 99 && classNumber == 99)
        {
            return "9999";
        }

        if (grade <= 0 || classNumber <= 0)
        {
            return "0000";
        }

        return $"{grade:00}{classNumber:00}";
    }
}