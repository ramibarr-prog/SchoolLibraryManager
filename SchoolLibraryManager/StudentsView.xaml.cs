using System.Windows;
using System.Windows.Controls;

namespace SchoolLibraryManager;

public partial class StudentsView : UserControl
{
    private readonly StudentsWindow _window;

    public StudentsView(DatabaseService databaseService)
    {
        InitializeComponent();
        _window = new StudentsWindow(databaseService);
        AttachWindowContent(_window);
    }

    public void FocusStudentById(string stuId)
    {
        _window.FocusStudentById(stuId);
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
