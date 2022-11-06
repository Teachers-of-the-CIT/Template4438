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
using System.Windows.Navigation;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4438
{
    /// <summary>
    /// Логика взаимодействия для _4438_Nasibullin.xaml
    /// </summary>
    public partial class _4438_Nasibullin : Window
    {
        public _4438_Nasibullin()
        {
            InitializeComponent();
        }

        private void importBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
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
            using (BaseModelEntities baseModelEntities = new BaseModelEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    if (list[i, 0] == "") continue;
                    baseModelEntities.Orders.Add(new Order()
                    {
                        id = int.Parse(list[i, 0]),
                        OrderCode = list[i, 1],
                        CreateDate = list[i,2],
                        OrderTime = list[i, 3],
                        ClientCode = int.Parse(list[i, 4]),
                        Services = list[i, 5],
                        Status = list[i, 6],
                        CloseDate = list[i, 7],
                        RentialTime = list[i, 8]
                    });
                    baseModelEntities.SaveChanges();
                }
            }
        }

        private void exportBtn_Click(object sender, RoutedEventArgs e)
        {
            List<Order> order;
            List<String> CreateDate;
            using (BaseModelEntities baseModelEntities = new BaseModelEntities())
            {
                order = baseModelEntities.Orders.ToList().OrderBy(x => x.CreateDate).ToList();
                CreateDate = baseModelEntities.Orders.Select(x => x.CreateDate).Distinct().ToList();
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = CreateDate.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < CreateDate.Count; i++)
                {
                    int startrowindex = 1;
                    var worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = CreateDate[i];
                    worksheet.Cells[1][startrowindex] = "Id";
                    worksheet.Cells[2][startrowindex] = "Код заказа";
                    worksheet.Cells[3][startrowindex] = "Код клиента";
                    worksheet.Cells[4][startrowindex] = "Услуги";
                    startrowindex++;
                    foreach (var ord in order)
                    {
                        if (worksheet.Name == ord.CreateDate.ToString().Split(' ')[0])
                        {
                            worksheet.Cells[1][startrowindex] = ord.id;
                            worksheet.Cells[2][startrowindex] = ord.OrderCode;
                            worksheet.Cells[3][startrowindex] = ord.ClientCode;
                            worksheet.Cells[4][startrowindex] = ord.Services;
                            startrowindex++;
                        }
                    }
                    worksheet.Columns.AutoFit();
                }
                app.Visible = true;
            }
        }
    }
}
