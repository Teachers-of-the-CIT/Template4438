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
        public static VasilevWindow vasilevWindow;
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
        //private void BnRahimov_Click(object sender, RoutedEventArgs e)
        //{
        //    RahimovWindow rahimov = new RahimovWindow();
        //    RahimovWindow.Show();
        //    this.Hide();
        //}
        private void BnRahimov_Click(object sender, RoutedEventArgs e)
        {
            //RahimovWindow rahimov = new RahimovWindow();
            //RahimovWindow.Show();
            //this.Hide();
        }
        private void BtnVasilev_Click(object sender, RoutedEventArgs e)
        {
            if (vasilevWindow == null)
            {
                vasilevWindow = new VasilevWindow();
                vasilevWindow.Show();
                this.Close();
            }
        }
        private void BGallyamov_Click(object sender, RoutedEventArgs e)
        {
            GallyamovARwindow g = new GallyamovARwindow();
            g.Show();
            this.Hide();
        }
        private void BTovmach_Click(object sender, RoutedEventArgs e)
        {
            _4438_Tovmach g = new _4438_Tovmach();
            g.Show();
        }
        private void BtnSharafetdinov_Click(object sender, RoutedEventArgs e)
        {
            SharafetdinovWindow NikitaW = new SharafetdinovWindow();
            NikitaW.Show();
        }
        private void btnGalimov_Click(object sender, RoutedEventArgs e)
        {
            _4438_Galimov window = new _4438_Galimov();
            window.Show();
        }
        private void BtnNasibullin_Click(object sender, RoutedEventArgs e)
        {
            _4438_Nasibullin _4438_Nasibullin = new _4438_Nasibullin();
            _4438_Nasibullin.Show();
        }
        private void BtnShigabiev_Click(object sender, RoutedEventArgs e)
        {
            ShigabievWindow ShamilW = new ShigabievWindow();
            ShamilW.Show();
            this.Hide();
        }

        private void BtnBritvina_Click(object sender, RoutedEventArgs e)
        {
            _4438_Britvina window = new _4438_Britvina();
            window.Show();
        }
        private void BtnKhaev_Click(object sender, RoutedEventArgs e)
        {
            KhaevWindow window = new KhaevWindow();
            window.Show();
        }
        private void BtnNabiullina_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            NabiullinaWindow nw = new NabiullinaWindow();
            nw.Show();
        }
        private void BtnSafiullin_Click(object sender, RoutedEventArgs e)
        {
            _4438SafiullinRR window = new _4438SafiullinRR();
            window.Show();
        }
    }
}
