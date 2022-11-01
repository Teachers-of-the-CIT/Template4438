using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
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
    /// Логика взаимодействия для SharafetdinovWindow.xaml
    /// </summary>
    public partial class SharafetdinovWindow : Window
    {
        public SharafetdinovWindow()
        {
            InitializeComponent();
        }

        private void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true)) return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
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
            using (ISRPO_Laba2Entities me = new ISRPO_Laba2Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    if (list[i, 0] == "") continue;
                    me.Labatwo.Add(new Labatwo()
                    {
                        Id = int.Parse(list[i, 0]),
                        OrderCode = list[i, 1],
                        Date = DateTime.Parse(list[i, 2], CultureInfo.GetCultureInfo("ru-ru")),
                        Time = TimeSpan.Parse(list[i, 3]),
                        ClientCode = int.Parse(list[i, 4]),
                        Services = list[i, 5],
                        Status = list[i, 6],
                        Closing_Date = list[i, 7] == "" ? default(DateTime) : DateTime.Parse(DateTime.Parse(list[i, 7]).ToLongDateString()),
                        Rental_Time = list[i, 8]
                    });
                    me.SaveChanges();
                }
            }
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            List<Labatwo> labatwos;
            List<string> status;
            using (ISRPO_Laba2Entities me = new ISRPO_Laba2Entities())
            {
                labatwos = me.Labatwo.ToList().OrderBy(x => x.Status).ToList();
                status = me.Labatwo.Select(x => x.Status).Distinct().ToList();
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = status.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < status.Count; i++)
                {
                    int startrowindex = 1;
                    var worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = status[i];
                    worksheet.Cells[1][startrowindex] = "Id";
                    worksheet.Cells[2][startrowindex] = "Код заказа";
                    worksheet.Cells[3][startrowindex] = "Дата создания";
                    worksheet.Cells[4][startrowindex] = "Код клиента";
                    worksheet.Cells[5][startrowindex] = "Услуги";
                    startrowindex++;
                    foreach (var order in labatwos)
                    {
                        if (worksheet.Name == order.Status)
                        {
                            worksheet.Cells[1][startrowindex] = order.Id;
                            worksheet.Cells[2][startrowindex] = order.OrderCode;
                            worksheet.Cells[3][startrowindex] = order.Date;
                            worksheet.Cells[4][startrowindex] = order.ClientCode;
                            worksheet.Cells[5][startrowindex] = order.Services;
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
