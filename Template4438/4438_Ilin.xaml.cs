using System.Windows;

namespace Template4438
{
    public partial class IlinWindow : Window
    {
        public IlinWindow()
        {
            InitializeComponent();
        }

        private void backButton_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }
    }
}
