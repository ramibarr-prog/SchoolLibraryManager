using System.Windows;
using System.Windows.Controls;

namespace SchoolLibraryManager;

public partial class ReportsView : UserControl
{
    private readonly ReportsWindow _window;

    public ReportsView(DatabaseService databaseService)
    {
        InitializeComponent();
        _window = new ReportsWindow(databaseService);
        AttachWindowContent(_window);
        Loaded += (_, _) => _window.ReloadAll();
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
