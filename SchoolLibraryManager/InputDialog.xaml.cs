using System.Windows;
using System.Windows.Input;

namespace SchoolLibraryManager;

public partial class InputDialog : Window
{
    public string InputText { get; private set; } = string.Empty;

    public InputDialog(string prompt, string title, string defaultValue = "")
    {
        InitializeComponent();
        Title = title;
        TxtPrompt.Text = prompt;
        TxtInput.Text = defaultValue;
    }

    private void Window_OnContentRendered(object? sender, EventArgs e)
    {
        TxtInput.Focus();
        TxtInput.SelectAll();
    }

    private void TxtInput_OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            BtnOk_OnClick(sender, e);
        }
    }

    private void BtnOk_OnClick(object sender, RoutedEventArgs e)
    {
        InputText = TxtInput.Text.Trim();
        DialogResult = true;
    }

    private void BtnCancel_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
