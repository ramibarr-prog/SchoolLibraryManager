using System.Windows;

namespace SchoolLibraryManager;

public partial class ReturnDateDialog : Window
{
    public DateTime SelectedDate { get; private set; }

    public ReturnDateDialog(DateTime? initialDate = null)
    {
        InitializeComponent();
        DpReturnDate.DisplayDateStart = DateTime.Today;
        var date = (initialDate ?? DateTime.Today).Date;
        DpReturnDate.SelectedDate = date < DateTime.Today ? DateTime.Today : date;
        SelectedDate = DpReturnDate.SelectedDate ?? DateTime.Today;
    }

    private void BtnSave_OnClick(object sender, RoutedEventArgs e)
    {
        var pickedDate = DpReturnDate.SelectedDate;
        if (pickedDate is null)
        {
            MessageBox.Show(this, "יש לבחור תאריך החזרה.", "הארכת השאלה", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SelectedDate = pickedDate.Value.Date;
        DialogResult = true;
        Close();
    }

    private void BtnCancel_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
