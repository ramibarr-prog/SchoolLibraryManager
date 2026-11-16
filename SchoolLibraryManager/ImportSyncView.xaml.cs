using System.Windows;
using System.Windows.Controls;

namespace SchoolLibraryManager;

public partial class ImportSyncView : UserControl
{
    private readonly ImportSyncWindow _window;

    public ImportSyncView(DatabaseService databaseService)
    {
        InitializeComponent();
        _window = new ImportSyncWindow(databaseService);
        AttachWindowContent(_window);
    }

    private void AttachWindowContent(Window window)
    {
        if (window.Content is not UIElement element)
        {
            return;
        }

        window.Content = null;
        Host.Children.Add(element);
    }
}
