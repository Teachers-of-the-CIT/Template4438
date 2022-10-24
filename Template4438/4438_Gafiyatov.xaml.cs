using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace Template4438
{
    /// <summary>
    /// Interaction logic for _4438_Gafiyatov.xaml
    /// </summary>
    public partial class _4438_Gafiyatov : Window
    {
        public _4438_Gafiyatov()
        {
            InitializeComponent();
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
            using (LR2Entities entities = new LR2Entities())
            {
                for (int i = 1; i < _rows - 1; i++)
                {
                    entities.Services.Add(new Service()
                    {
                        ID = Int32.Parse(list[i, 0]),
                        Service_Name = list[i, 1],
                        Service_Type = list[i, 2],
                        Service_Code = list[i, 3],
                        Cost = Int32.Parse(list[i, 4])
                    });
                }
                entities.SaveChanges();
            }
        }

        private void ExportBTN_Click(object sender, RoutedEventArgs e)
        {
            List<Service> allServices;
            using (LR2Entities userEntities = new LR2Entities())
            {
                allServices = userEntities.Services.ToList().OrderBy(s => s.Service_Name).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            var studentsCategories = allServices.OrderBy(o => o.Cost).GroupBy(s => s.ID)
                    .ToDictionary(g => g.Key, g => g.Select(s => new { s.ID, s.Service_Name, s.Service_Type, s.Cost }).ToArray());
            for (int i = 0; i < 3; i++)
            {
                int startRowIndex = 1;
                var worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Категория {i + 1}";
                worksheet.Cells[1][startRowIndex] = "Id";
                worksheet.Cells[2][startRowIndex] = "Название услуги";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                startRowIndex++;

                var data = i == 0 ? studentsCategories.Where(w => w.Value.All(p => p.Service_Type.Equals("Rental")))
                         : i == 1 ? studentsCategories.Where(w => w.Value.All(p => p.Service_Type.Equals("Learning")))
                         : i == 2 ? studentsCategories.Where(w => w.Value.All(p => p.Service_Type.Equals("Rise"))) : studentsCategories;

                foreach (var students in data)
                {
                    foreach (var student in students.Value)
                    {
                        worksheet.Cells[1][startRowIndex] = student.ID;
                        worksheet.Cells[2][startRowIndex] = student.Service_Name;
                        worksheet.Cells[3][startRowIndex] = student.Cost;
                        startRowIndex++;
                    }
                }
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        private async void ImportJSON_BTN_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Spisok.json)|*.json",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate);
            var service = await JsonSerializer.DeserializeAsync<List<Service>>(fs);

            using (LR2Entities db = new LR2Entities())
            {
                foreach (Service serv in service)
                {
                    Service services = new Service();
                    services.ID = serv.ID;
                    services.Service_Name = serv.Service_Name;
                    services.Service_Type = serv.Service_Type;
                    services.Service_Code = serv.Service_Code;
                    services.Cost = serv.Cost;
                    db.Services.Add(services);
                }
                db.SaveChanges();
            }
        }

        private void ExportWordBTN_Click(object sender, RoutedEventArgs e)
        {
            List<Service> allServices;
            using (LR2Entities usersEntities = new LR2Entities())
            {
                allServices = usersEntities.Services.ToList().OrderBy(s => s.Service_Name).ToList();
            }
            var costsCategories = allServices.OrderBy(o => o.Cost).GroupBy(s => s.ID)
                    .ToDictionary(g => g.Key, g => g.Select(s => new { s.ID, s.Service_Name, s.Service_Type, s.Cost }).ToArray());
            var app = new Word.Application();
            Word.Document document = app.Documents.Add();
            for (int i = 0; i < 3; i++)
            {
                var data = i == 0 ? costsCategories.Where(w => w.Value.All(p => p.Service_Type.Equals("Rental")))
                         : i == 1 ? costsCategories.Where(w => w.Value.All(p => p.Service_Type.Equals("Learning")))
                         : i == 2 ? costsCategories.Where(w => w.Value.All(p => p.Service_Type.Equals("Rise"))) : costsCategories;
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = $"Категория {i + 1}";
                
                range.InsertParagraphAfter();
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                var studentsTable = document.Tables.Add(tableRange, data.Select(s => s.Value.Length).Sum() + 1, 3);
                studentsTable.Borders.InsideLineStyle = studentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                studentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Word.Range cellRange;
                cellRange = studentsTable.Cell(1, 1).Range;
                cellRange.Text = "Id";
                cellRange = studentsTable.Cell(1, 2).Range;
                cellRange.Text = "Название услуги";
                cellRange = studentsTable.Cell(1, 3).Range;
                cellRange.Text = "Стоимость";
                studentsTable.Rows[1].Range.Bold = 1;
                studentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                int row = 1;
                var stepSize = 1;
                foreach (var group in data)
                {
                    foreach (var currentCost in group.Value)
                    {
                        cellRange = studentsTable.Cell(row + stepSize, 1).Range;
                        cellRange.Text = currentCost.ID.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = studentsTable.Cell(row + stepSize, 2).Range;
                        cellRange.Text = currentCost.Service_Name;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = studentsTable.Cell(row + stepSize, 3).Range;
                        cellRange.Text = currentCost.Cost.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        row++;
                    }
                }
                Word.Paragraph countCostsParagraph = document.Paragraphs.Add();
                Word.Range countCostsRange = countCostsParagraph.Range;
                countCostsRange.Text = $"Количество услуг - {data.Select(s => s.Value.Length).Sum()} ";
                countCostsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                countCostsRange.InsertParagraphAfter();
                document.Words.Last.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            }
            app.Visible = true;
        }
    }
}
