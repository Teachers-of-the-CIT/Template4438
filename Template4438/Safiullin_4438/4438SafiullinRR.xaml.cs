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
using System.Windows.Shapes;

namespace Template4438
{
    /// <summary>
    /// Логика взаимодействия для _4438SafiullinRR.xaml
    /// </summary>
    public partial class _4438SafiullinRR : Window
    {
        public _4438SafiullinRR()
        {
            InitializeComponent();
        }

        private void backBTN_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = new MainWindow();
            main.Show();
            this.Close();
        }

        private void importBTN_Click(object sender, RoutedEventArgs e)
        {

        }

        private void exportBTN_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
