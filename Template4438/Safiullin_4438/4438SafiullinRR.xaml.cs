using Microsoft.Win32;
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
using Template4438.Safiullin_4438.Database;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Template4438
{
    /// <summary>
    /// Логика взаимодействия для _4438SafiullinRR.xaml
    /// </summary>
    public partial class _4438SafiullinRR : Window
    {
        public List<ExcelEntity> excel_data;
        public _4438SafiullinRR()
        {
            InitializeComponent();
            displayTB.Text = GetDBatSTR();
        }

        private void backBTN_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = new MainWindow();
            main.Show();
            this.Close();
        }

        private void importBTN_Click(object sender, RoutedEventArgs e)
        {
            using (ExcelEntityContainer excelEntity = new ExcelEntityContainer())
            {
                if (excelEntity.ExcelEntitySet.Count() > 0)
                {
                    MessageBox.Show("В базе данных что-то было, но уже автоматически удалено. Нажмите кнопку Импорт еще раз.");
                    foreach(var item in excelEntity.ExcelEntitySet)
                    {
                        excelEntity.ExcelEntitySet.Remove(item);
                    }
                    excelEntity.SaveChanges();
                    displayTB.Text = GetDBatSTR();
                    return;
                }
            }
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx;*.json",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            Excel_Entity data_str = GetData_ToString_FromXL(ofd.FileName);

            var str = "";
            using (ExcelEntityContainer excelEntity = new ExcelEntityContainer())
            {
                for (int i = 0; i < data_str.rows; i++)
                {
                    if (data_str.data[i, 1] == "" || data_str.data[i, 1] == " ")
                        continue;
                    if (data_str.data[i, 1] == "Наименование услуги")
                        continue;
                    excelEntity.ExcelEntitySet.Add(new ExcelEntity()
                    {
                        ServiceName = data_str.data[i, 1],
                        ServiceType = data_str.data[i, 2],
                        ServiceCode = data_str.data[i, 3],
                        ServicePrice = int.Parse(data_str.data[i, 4]),
                    });
                }
                excelEntity.SaveChanges();
                displayTB.Text = GetDBatSTR();
            }

        }
        private string GetDBatSTR()
        {
           
            string str = "Список услуг:\n";
            using (var db = new ExcelEntityContainer())
            {
                foreach(var item in db.ExcelEntitySet)
                {
                    if (item.Id == 1)
                        continue;
                    str += $"-->\t{item.ServiceName}\n";
                }
            }
            return str;
        }
        private Excel_Entity GetData_ToString_FromXL(string url)
        {
            string[,] list;

            Excel.Application ObjWorkExcel = new Excel.Application();

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(url);

            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            Excel_Entity ent =
                new Excel_Entity();

            ent.data = list;
            ent.columns = _columns;
            ent.rows = _rows;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            return
                ent;
        }
        public class Excel_Entity
        {
            public int rows { get; set; }
            public int columns { get; set; }
            public string[,] data { get; set; }

        }
        private void exportBTN_Click(object sender, RoutedEventArgs e)
        {
            using (ExcelEntityContainer excelEntity = new ExcelEntityContainer())
            {
                if (excelEntity.ExcelEntitySet.Count() < 1)
                {
                    MessageBox.Show("Добавь данные в БД, пожалуйста!");
                    return;
                }
            }

            using (var db = new ExcelEntityContainer())
            {
                excel_data = db.ExcelEntitySet.ToList().OrderBy(x=>x.ServicePrice).ToList();
            }
            var list_times = excel_data.Select(x => x.ServiceType).Distinct().ToList();
            foreach(var item in list_times)
            {
                if(item.Contains("Вид услуги"))
                {
                    list_times.Remove(item);
                    break;
                }    
            }


            var app = new Excel.Application();
            app.SheetsInNewWorkbook = list_times.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < list_times.Count(); i++)
            {
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(list_times[i]);

                int j = 1;
                worksheet.Cells[1][j] = "ID";
                worksheet.Cells[2][j] = "Название услуги";
                worksheet.Cells[3][j] = "Стоимость";
                j = 2;
                foreach (var item in excel_data)
                {
                    if (item.ServiceType == worksheet.Name)
                    {
                        worksheet.Cells[1][j] = item.Id;
                        worksheet.Cells[2][j] = item.ServiceName;
                        worksheet.Cells[3][j] = item.ServicePrice;
                        j++;
                    }
                }
            }
            app.Visible = true;
        }
    }
}
