using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Template4438
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BnTask_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BnKarnaukh_Click(object sender, RoutedEventArgs e)
        {
            KarnaukhWindow window = new KarnaukhWindow();
            window.Show();
            this.Hide();
        }

        private void BnVafin_Click(object sender, RoutedEventArgs e)
        {
            _4438_Vafin vafinWindow = new _4438_Vafin();
            vafinWindow.Show();
            this.Hide();
        }

        private void BnRahimov_Click(object sender, RoutedEventArgs e)
        {
            RahimovWindow rahimov = new RahimovWindow();
            rahimov.Show();
            this.Hide();
        }
    }
}
