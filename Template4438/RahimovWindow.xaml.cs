
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
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Markup;

namespace Template4438
{
    /// <summary>
    /// Логика взаимодействия для RahimovWindow.xaml
    /// </summary>
    public partial class RahimovWindow : Window
    {
        public RahimovWindow()
        {
            InitializeComponent();
        }


        private void ExportBTN_Click(object sender, RoutedEventArgs e)
        {
            List<MainTable> allServices;

            using (var userEntities = new ISRPO1Entities1())
            {
                allServices = userEntities.MainTable.ToList().OrderBy(x => x.CreateDate).ToList();

                var datecreate = userEntities.MainTable.Select(x => x.CreateDate).Distinct().ToList();

                var app = new Excel.Application();
                app.SheetsInNewWorkbook = datecreate.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < datecreate.Count(); i++)
                {
                    int j = 1;
                    var worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = datecreate[i];

                    worksheet.Cells[1][j] = "Id";
                    worksheet.Cells[2][j] = "Код заказа";
                    worksheet.Cells[3][j] = "Код клиента";
                    worksheet.Cells[4][j] = "Услуги";
                    j = 2;
                    foreach (var services in allServices)
                    {
                        if (worksheet.Name == services.CreateDate.ToString())
                        {
                            worksheet.Cells[1][j] = services.ID;
                            worksheet.Cells[2][j] = services.CodeOrder;
                            worksheet.Cells[3][j] = services.CodeClient;
                            worksheet.Cells[4][j] = services.Services;
                            j++;
                        }

                    }
                    worksheet.Columns.AutoFit();
                }
                app.Visible = true;
            }
        }

        private void ImportBTN_Click(object sender, RoutedEventArgs e)
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
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (ISRPO1Entities1 entities = new ISRPO1Entities1())
            {
                for (int i = 1; i < _rows; i++)
                {
                    if (list[i, 1] == "" || list[i, 0] == " ")
                    {
                        continue;
                    }
                    entities.MainTable.Add(new MainTable()
                    {
                        ID = Int32.Parse(list[i, 0]),
                        CodeOrder = list[i, 1],
                        CreateDate = list[i, 2],
                        CreateTime = list[i, 3],
                        CodeClient = list[i, 4],
                        Services = list[i, 5],
                        Status = list[i, 6],
                        ClosedDate = list[i, 7],
                        ProkatTime = list[i, 8]
                    }); ;
                }
                entities.SaveChanges();
            }

        }

        private void ExportWordBTN_Click(object sender, RoutedEventArgs e)
        {
            List<MainTable> allServices;
            using (var userEntities = new ISRPO1Entities1())
            {
                allServices = userEntities.MainTable.ToList();
                var dateList = allServices.Select(x => DateTime.Parse(x.CreateDate.ToString()).ToShortDateString()).Distinct().OrderBy(x=>x).ToList();
                
                var app = new Word.Application();
                
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
                    Word.Table studentsTable = document.Tables.Add(tableRange, allServices.Select(x => DateTime.Parse(x.CreateDate).ToShortDateString().ToString()).Where(x => range.Text.Contains(x)).ToList().Count() + 1, 4);
                   

                    foreach (var services in allServices)
                    {
                        var itemlist = DateTime.Parse(services.CreateDate).ToShortDateString().ToString();
                        if (range.Text.Contains(itemlist))
                        {
                            studentsTable.Borders.InsideLineStyle = 
                                studentsTable.Borders.OutsideLineStyle = 
                                Word.WdLineStyle.wdLineStyleSingle;
                            studentsTable.Range.Cells.VerticalAlignment = 
                                Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            cellRange = studentsTable.Cell(1, 1).Range;
                            cellRange.Text = "ID";
                            cellRange = studentsTable.Cell(1, 2).Range;
                            cellRange.Text = "Код Заказа";
                            cellRange = studentsTable.Cell(1, 3).Range;
                            cellRange.Text = "Код Клиента";
                            cellRange = studentsTable.Cell(1, 4).Range;
                            cellRange.Text = "Услуги";
                            
                            studentsTable.Rows[1].Range.Bold = 1;
                            studentsTable.Rows[1].Range.ParagraphFormat.Alignment = 
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = studentsTable.Cell(row + 1, 1).Range;
                            cellRange.Text = services.ID.ToString();
                            cellRange.ParagraphFormat.Alignment = 
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = studentsTable.Cell(row + 1, 2).Range;
                            cellRange.Text = services.CodeOrder;
                            cellRange.ParagraphFormat.Alignment = 
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = studentsTable.Cell(row + 1, 3).Range;
                            cellRange.Text = services.CodeClient.ToString();
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = studentsTable.Cell(row + 1, 4).Range;
                            cellRange.Text = services.Services.ToString();
                            cellRange.ParagraphFormat.Alignment = 
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            row = row+1;
                        }

                    }
                    app.Visible = true;
           
                }
            }
        }

            private async void JSONImportBTN_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Spisok.json)|*.json",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            FileStream FS = new FileStream(ofd.FileName, FileMode.OpenOrCreate);
            var maintable = await JsonSerializer.DeserializeAsync<List<MainTable>>(FS);

            using (ISRPO1Entities1 db = new ISRPO1Entities1())
            {
                foreach (MainTable m in maintable)
                {
                    MainTable maint = new MainTable();
                    maint.ID = m.ID;
                    maint.ProkatTime = m.ProkatTime;
                    maint.CreateTime = m.CreateTime;
                    maint.CodeClient = m.CodeClient;
                    maint.CodeOrder = m.CodeOrder;
                    maint.CreateDate = m.CreateDate;
                    maint.Status = m.Status;
                    maint.ClosedDate = m.ClosedDate;
                    db.MainTable.Add(maint);
                }
                db.SaveChanges();
            }
        }
    }
}
  

