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
using Excel = Microsoft.Office.Interop.Excel;
namespace Template4438
{
    /// <summary>
    /// Логика взаимодействия для _4438_Tovmach.xaml
    /// </summary>
    public partial class _4438_Tovmach : Window
    {
        public _4438_Tovmach()
        {
            InitializeComponent();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (import.xlsx)|*.xlsx",
                Title = "Выберите файл базы для импорта в Базу данных"
            };
            if (!(openFileDialog.ShowDialog() == true)) return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(openFileDialog.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (excelimportEntities excelimportEntities = new excelimportEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    if (list[i, 0] == "") continue;
                    excelimportEntities.Orders.Add(new Order()
                    {
                        id = int.Parse(list[i, 0]),
                        OrderCode = list[i, 1],
                        CreateDate = list[i, 2],
                        OrderTime = list[i, 3],
                        ClientCode = int.Parse(list[i, 4]),
                        Services = list[i, 5],
                        Status = list[i, 6],
                        CloseData = list[i, 7],
                        RentialTime = list[i, 8]
                    });
                    excelimportEntities.SaveChanges();
                }
            }
            MessageBox.Show("Данные успешно импортированы!");
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            List<Order> order;
            List<String> OrderTime;
            using (excelimportEntities excelimportEntities = new excelimportEntities())
            {
                order = excelimportEntities.Orders.ToList().OrderBy(x => x.OrderTime).ToList();
                OrderTime = excelimportEntities.Orders.Select(x => x.OrderTime).Distinct().ToList();
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = OrderTime.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < OrderTime.Count; i++)
                {
                    int startrowindex = 1;
                    var worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = OrderTime[i].Replace(':', '-');
                    worksheet.Cells[1][startrowindex] = "Id";
                    worksheet.Cells[2][startrowindex] = "Код заказа";
                    worksheet.Cells[3][startrowindex] = "Код клиента";
                    worksheet.Cells[4][startrowindex] = "Услуги";
                    worksheet.Cells[5][startrowindex] = "Дата создания";

                    startrowindex++;
                    foreach (var ord in order)
                    {
                            worksheet.Cells[1][startrowindex] = ord.id;
                            worksheet.Cells[2][startrowindex] = ord.OrderCode;                           
                            worksheet.Cells[3][startrowindex] = ord.ClientCode;
                            worksheet.Cells[4][startrowindex] = ord.Services;
                            worksheet.Cells[5][startrowindex] = ord.CreateDate;

                            startrowindex++;   
                    }
                    worksheet.Columns.AutoFit();
                }
                app.Visible = true;
            }
        }
    }
}
