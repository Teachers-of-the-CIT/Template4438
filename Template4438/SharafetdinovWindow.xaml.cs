using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
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
using Window = System.Windows.Window;
using Word = Microsoft.Office.Interop.Word;

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

        private void WordExportBtn_Click(object sender, RoutedEventArgs e)
        {
            using (ISRPO_Laba2Entities ent = new ISRPO_Laba2Entities())
            {
                List<Labatwo> ordersList = ent.Labatwo.ToList();
                List<string> statusList = ordersList.Select(x => x.Status).Distinct().ToList();

                Word.Application app = new Word.Application();

                Word.Document document = app.Documents.Add();

                for (int i = 0; i < statusList.Count(); i++)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = statusList[i].ToString();
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;


                    Word.Range cellRange;
                    int row = 1;
                    Word.Table ordersTable = document.Tables.Add(tableRange, ordersList.Select(x => x.Status).Where(x => range.Text.Contains(x)).ToList().Count() + 1, 5);


                    foreach (Labatwo orders in ordersList)
                    {
                        var itemlist = orders.Status;
                        if (range.Text.Contains(itemlist))
                        {
                            ordersTable.Borders.InsideLineStyle =
                                ordersTable.Borders.OutsideLineStyle =
                                Word.WdLineStyle.wdLineStyleSingle;
                            ordersTable.Range.Cells.VerticalAlignment =
                                Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            cellRange = ordersTable.Cell(1, 1).Range;
                            cellRange.Text = "Id";
                            cellRange = ordersTable.Cell(1, 2).Range;
                            cellRange.Text = "Код Заказа";
                            cellRange = ordersTable.Cell(1, 3).Range;
                            cellRange.Text = "Дата создания";
                            cellRange = ordersTable.Cell(1, 4).Range;
                            cellRange.Text = "Код Клиента";
                            cellRange = ordersTable.Cell(1, 5).Range;
                            cellRange.Text = "Услуги";

                            ordersTable.Rows[1].Range.Bold = 1;
                            ordersTable.Rows[1].Range.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = ordersTable.Cell(row + 1, 1).Range;
                            cellRange.Text = orders.Id.ToString();
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = ordersTable.Cell(row + 1, 2).Range;
                            cellRange.Text = orders.OrderCode;
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = ordersTable.Cell(row + 1, 3).Range;
                            cellRange.Text = orders.Date.ToString();
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = ordersTable.Cell(row + 1, 4).Range;
                            cellRange.Text = orders.ClientCode.ToString();
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = ordersTable.Cell(row + 1, 5).Range;
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

        private void JsonImprtoBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Spisok.json)|*.json",
                Title = "выберите файл базы данных"
            };
            if (!(dialog.ShowDialog() == true)) return;
            FileStream fileStream = new FileStream(dialog.FileName, FileMode.OpenOrCreate);
            List<JsonOrders> ordersList = JsonSerializer.Deserialize<List<JsonOrders>>(fileStream);
            using (ISRPO_Laba2Entities ent = new ISRPO_Laba2Entities())
            {
                foreach (JsonOrders o in ordersList)
                {
                    ent.Labatwo.Add(new Labatwo()
                    {
                        OrderCode = o.CodeOrder,
                        Date = DateTime.Parse(o.CreateDate),
                        Time = DateTime.Parse(o.CreateTime).TimeOfDay,
                        ClientCode = Convert.ToInt32(o.CodeClient),
                        Services = o.Services,
                        Status = o.Status,
                        Closing_Date = o.CreateDate == "" ? default(DateTime) : DateTime.Parse(o.CreateDate),
                        Rental_Time = o.ProkatTime
                    });
                    ent.SaveChanges();
                }
            }
        }
    }
}
