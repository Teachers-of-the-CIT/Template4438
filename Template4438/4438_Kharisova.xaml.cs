using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;


namespace Template4438
{
    /// <summary>
    /// Логика взаимодействия для _4438_Kharisova.xaml
    /// </summary>
    public partial class _4438_Kharisova : Window
    {
        public _4438_Kharisova()
        {
            InitializeComponent();
        }


        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {

            List<string> allRent = new List<string>();
            List<Order> allOrders = new List<Order>();
            using (OrderDBEntities orderEntities = new OrderDBEntities())
            {
                allOrders = orderEntities.Order.ToList().OrderBy(s => s.rentaltime).ToList();

                for (int i = 0; i < allOrders.Count(); i++)
                {
                    bool flag = true;
                    for (int j = 0; j < allRent.Count(); j++)
                    {
                        if (allOrders[i].rentaltime == allRent[j])
                        {
                            flag = false;
                        }
                    }
                    if (flag && allOrders[i].rentaltime != null && allOrders[i].rentaltime != "")
                        allRent.Add(allOrders[i].rentaltime);
                }
            }

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allRent.Count();
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            var orderCategories = allOrders.GroupBy(s => s.rentaltime).ToList();


            for (int i = 0; i < allRent.Count(); i++)
            {
                int startRowIndex = 2;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = allRent[i];
                worksheet.Cells[1][startRowIndex] = "id";
                worksheet.Cells[2][startRowIndex] = "ordercode";
                worksheet.Cells[3][startRowIndex] = "createdate";
                worksheet.Cells[4][startRowIndex] = "clientcode";
                worksheet.Cells[5][startRowIndex] = "features";
                startRowIndex++;
                foreach (var order in orderCategories)
                {
                    if (order.Key == allRent[i])
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = allRent[i];
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;
                        foreach (Order orders in allOrders)
                        {
                            if (orders.rentaltime == allRent[i])
                            {
                                worksheet.Cells[1][startRowIndex] = orders.id;
                                worksheet.Cells[2][startRowIndex] = orders.ordercode;
                                worksheet.Cells[3][startRowIndex] = orders.createdate;
                                worksheet.Cells[4][startRowIndex] = orders.clientcode;
                                worksheet.Cells[5][startRowIndex] = orders.feauters;
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
            using (OrderDBEntities orderEntities = new OrderDBEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    orderEntities.Order.Add(new Order()
                    {
                        id = i,
                        ordercode = list[i, 1],
                        createdate = list[i, 2],
                        ordertime = list[i, 3],
                        clientcode = list[i, 4],
                        feauters = list[i, 5],
                        status = list[i, 6],
                        enddate = list[i, 7],
                        rentaltime = list[i, 8]
                    });
                }
                orderEntities.SaveChanges();
            }

        }
    }
}
