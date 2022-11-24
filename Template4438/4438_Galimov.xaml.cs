using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;
namespace Template4438
{
    /// <summary>
    /// Логика взаимодействия для _4438_Galimov.xaml
    /// </summary>
    public partial class _4438_Galimov : Window
    {
        public _4438_Galimov()
        {
            InitializeComponent();
        }

        private void importBtn_Click(object sender, RoutedEventArgs e)
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

        private void exportBtn_Click(object sender, RoutedEventArgs e)
        {
            List<Order> order;
            List<String> CreateDate;
            using (excelimportEntities excelimportEntities = new excelimportEntities())
            {
                order = excelimportEntities.Orders.ToList().OrderBy(x => x.CreateDate).ToList();
                CreateDate = excelimportEntities.Orders.Select(x => x.CreateDate).Distinct().ToList();
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

        async private void importJsonBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Spisok.json)|*.json",
                Title = "выберите файл базы данных"
            };
            if (!(dialog.ShowDialog() == true)) return;
            FileStream fileStream = new FileStream(dialog.FileName, FileMode.OpenOrCreate);
            List<JsonOrder> ordersList = await JsonSerializer.DeserializeAsync<List<JsonOrder>>(fileStream);
            using (excelimportEntities ent = new excelimportEntities())
            {
                foreach (JsonOrder o in ordersList)
                {
                    ent.Orders.Add(new Order()
                    {
                        OrderCode = o.CodeOrder,
                        CreateDate = o.CreateDate,
                        OrderTime = o.CreateTime,
                        ClientCode = Int32.Parse(o.CodeClient),
                        Services = o.Services,
                        Status = o.Status,
                        CloseData = o.CreateDate == "" ? default(DateTime).ToString() : o.CreateDate,
                        RentialTime = o.ProkatTime
                    });
                    ent.SaveChanges();
                }
            }
            MessageBox.Show("Данные успешно импортированы!");
        }

        private void exportJsonBtn_Click(object sender, RoutedEventArgs e)
        {
            using (excelimportEntities ent = new excelimportEntities())
            {
                List<Order> ordersList = ent.Orders.ToList();
                List<string> dateList = ordersList.Select(x => x.CreateDate).Distinct().ToList();

                Word.Application app = new Word.Application();

                Word.Document document = app.Documents.Add();

                for (int i = 0; i < dateList.Count(); i++)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = dateList[i].ToString();
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;


                    Word.Range cellRange;
                    int row = 1;
                    Word.Table ordersTable = document.Tables.Add(tableRange, ordersList.Select(x => DateTime.Parse(Convert.ToString(x.CreateDate)).ToShortDateString().ToString()).Where(x => range.Text.Contains(x)).ToList().Count() + 1, 4);


                    foreach (Order orders in ordersList)
                    {
                        var itemlist = DateTime.Parse(Convert.ToString(orders.CreateDate)).ToShortDateString().ToString();
                        if (range.Text.Contains(itemlist))
                        {
                            ordersTable.Borders.InsideLineStyle =
                                ordersTable.Borders.OutsideLineStyle =
                                Word.WdLineStyle.wdLineStyleSingle;
                            ordersTable.Range.Cells.VerticalAlignment =
                                Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            cellRange = ordersTable.Cell(1, 1).Range;
                            cellRange.Text = "ID";
                            cellRange = ordersTable.Cell(1, 2).Range;
                            cellRange.Text = "Код Заказа";
                            cellRange = ordersTable.Cell(1, 3).Range;
                            cellRange.Text = "Код Клиента";
                            cellRange = ordersTable.Cell(1, 4).Range;
                            cellRange.Text = "Услуги";

                            ordersTable.Rows[1].Range.Bold = 1;
                            ordersTable.Rows[1].Range.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = ordersTable.Cell(row + 1, 1).Range;
                            cellRange.Text = orders.id.ToString();
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = ordersTable.Cell(row + 1, 2).Range;
                            cellRange.Text = orders.OrderCode;
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = ordersTable.Cell(row + 1, 3).Range;
                            cellRange.Text = orders.ClientCode.ToString();
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = ordersTable.Cell(row + 1, 4).Range;
                            cellRange.Text = orders.Services.ToString();
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            row = row + 1;
                        }
                    }
                    app.Visible = true;

                }
            }
        }
    }
}
