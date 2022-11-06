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
using Microsoft.Win32;

namespace Template4438
{
    /// <summary>
    /// Логика взаимодействия для _4438_Khalikov.xaml
    /// </summary>
    public partial class _4438_Khalikov : Window
    {
        public _4438_Khalikov()
        {
            InitializeComponent();
        }

        private void BtnExpot_Click(object sender, RoutedEventArgs e)
        {
            List<string> allDates = new List<string>();
            List<Orders> allOrders = new List<Orders>();
            using (LR2_ISRPOEntities LR2Entities = new LR2_ISRPOEntities())
            {
                allOrders = LR2Entities.Orders.ToList().OrderBy(s =>s.Дата_создания).ToList();
                
                for(int i = 0; i < allOrders.Count(); i++)
                {
                    bool flag = true;
                    for(int j = 0; j < allDates.Count(); j++)
                    {
                        if(allOrders[i].Дата_создания == allDates[j])
                        {
                            flag = false;
                        }
                    }
                    if (flag && allOrders[i].Дата_создания != null && allOrders[i].Дата_создания != "")
                        allDates.Add(allOrders[i].Дата_создания);
                }
            }

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allDates.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            var orderCategories = allOrders.GroupBy(s => s.Дата_создания).ToList();


            for (int i = 0; i < allDates.Count(); i++)
            {
                int startRowIndex = 2;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = allDates[i];
                worksheet.Cells[1][startRowIndex] = "Id";
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][startRowIndex] = "Код клиента";
                worksheet.Cells[4][startRowIndex] = "Услуги";
                startRowIndex++;
                foreach (var order in orderCategories)
                {
                    if (order.Key == allDates[i])
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = allDates[i];
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;
                        foreach (Orders orders in allOrders)
                        {
                            if (orders.Дата_создания == allDates[i])
                            {
                                worksheet.Cells[1][startRowIndex] = orders.ID;
                                worksheet.Cells[2][startRowIndex] = orders.Код_заказа;
                                worksheet.Cells[3][startRowIndex] = orders.Код_клиента;
                                worksheet.Cells[4][startRowIndex] = orders.Услуги;
                                startRowIndex++;
                            }
                        }
                    }
                    else
                    {
                        continue;
                    }
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ImportInDB(list, _rows);
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
        }

        private void ImportInDB(string[,] list, int _rows)
        {
            using (LR2_ISRPOEntities LR2Entities = new LR2_ISRPOEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    LR2Entities.Orders.Add(new Orders()
                    {
                        ID = i,
                        Код_заказа = list[i, 1],
                        Дата_создания = list[i, 2],
                        Время_заказа = list[i, 3],
                        Код_клиента = list[i, 4],
                        Услуги = list[i, 5],
                        Статус = list[i, 6],
                        Дата_закрытия = list[i, 7],
                        Время_проката = list[i, 8]
                    });
                }
                LR2Entities.SaveChanges();
            }
        }
    }
}
