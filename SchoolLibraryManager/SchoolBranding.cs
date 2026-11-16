using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace SchoolLibraryManager;

public static class SchoolBranding
{
    public const string SchoolName = "תיכון אתגרי העתיד חריש";
    public const string Motto = "מקום שטוב לגדול בו";
    public const string Contact = "רח' יקינטון 32, חריש | טלפון: 04-6514568";
    public const string LogoUri = "pack://application:,,,/Assets/school_logo.png";

    public static readonly Brush BrandBrush = new SolidColorBrush(Color.FromRgb(0x1D, 0x4F, 0x91));

    public static void AddLetterhead(FlowDocument document)
    {
        var textPanel = new Paragraph { Margin = new Thickness(0) };
        textPanel.Inlines.Add(new Run(SchoolName) { FontSize = 16, FontWeight = FontWeights.Bold, Foreground = BrandBrush });
        textPanel.Inlines.Add(new LineBreak());
        textPanel.Inlines.Add(new Run(Motto) { FontSize = 11, FontStyle = FontStyles.Italic, Foreground = Brushes.DimGray });
        textPanel.Inlines.Add(new LineBreak());
        textPanel.Inlines.Add(new Run(Contact) { FontSize = 10, Foreground = Brushes.Gray });

        var table = new Table { Margin = new Thickness(0, 0, 0, 12) };
        table.Columns.Add(new TableColumn { Width = new GridLength(60) });
        table.Columns.Add(new TableColumn());

        var group = new TableRowGroup();
        var row = new TableRow();

        var logoCell = new TableCell { BorderBrush = BrandBrush, BorderThickness = new Thickness(0, 0, 0, 1), Padding = new Thickness(0, 0, 0, 6) };
        var logo = TryCreateLogo();
        if (logo is not null)
        {
            logoCell.Blocks.Add(new BlockUIContainer(new System.Windows.Controls.Image { Source = logo, Height = 45, HorizontalAlignment = HorizontalAlignment.Right }));
        }

        var textCell = new TableCell(textPanel) { BorderBrush = BrandBrush, BorderThickness = new Thickness(0, 0, 0, 1), Padding = new Thickness(8, 0, 0, 6) };

        row.Cells.Add(logoCell);
        row.Cells.Add(textCell);
        group.Rows.Add(row);
        table.RowGroups.Add(group);

        document.Blocks.InsertBefore(document.Blocks.FirstBlock, table);
    }

    private static BitmapImage? TryCreateLogo()
    {
        try
        {
            var image = new BitmapImage();
            image.BeginInit();
            image.UriSource = new Uri(LogoUri, UriKind.Absolute);
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.EndInit();
            image.Freeze();
            return image;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "Failed to load school logo for print");
            return null;
        }
    }
}
