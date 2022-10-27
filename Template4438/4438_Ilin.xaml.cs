using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Linq;
using System.Text.Json;
using System.IO;
using System.Text;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Template4438
{
    public partial class IlinWindow : Window
    {
        public static EntityModelContainer db = new EntityModelContainer();
        public static string projectFolder = AppDomain.CurrentDomain.BaseDirectory;

        public IlinWindow()
        {
            InitializeComponent();
        }

        private void backButton_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

        public void ImportFromExcel(string filepath)
        {
            Excel.Application app   = new Excel.Application();
            Excel.Workbook    book  = app.Workbooks.Open(filepath);
            Excel.Worksheet   sheet = (Excel.Worksheet)book.Sheets[1];

            var dbServices = db.Services.ToArray();
            var lastCell = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int rows = lastCell.Row;
            int i = 0;
            int importedServices = 0;
            int declinedServices = 0;
            List<Service> servicesToAdd = new List<Service>(12);
            for (; i < rows - 1; i++)
            {
                Service service = new Service();
                service.Id    = int.Parse(sheet.Cells[i + 2, 1].Text);
                service.Name  = sheet.Cells[i + 2, 2].Text;
                service.Type  = sheet.Cells[i + 2, 3].Text;
                service.Code  = sheet.Cells[i + 2, 4].Text;
                service.Price = double.Parse(sheet.Cells[i + 2, 5].Text);

                if (dbServices.Length > 0)
                {
                    bool found = false;
                    foreach (var dbService in dbServices)
                    {
                        if (dbService.Id == service.Id)
                        {
                            declinedServices += 1;
                            found = true;
                            break;
                        }
                    }

                    if (!found) 
                    {
                        servicesToAdd.Add(service);
                        importedServices += 1;
                    }
                }
                else
                {
                    servicesToAdd.Add(service);
                    importedServices += 1;
                }

            }

            int totalServices = i;
            db.Services.AddRange(servicesToAdd);
            db.SaveChanges();
            MessageBox.Show(string.Format("Импортировано {0}/{1} сущностей.\n({2} сущностей совпадают с уже существующими)", importedServices, totalServices, declinedServices),
                            "Импорт Excel",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
        }

        public void ImportFromJson(string filepath)
        {
            var dbServices = db.Services.ToArray();
            int importedServices = 0;
            int declinedServices = 0;
            int i = 0;
            List<Service> servicesToAdd = new List<Service>(12);
            using (StreamReader reader = new StreamReader(filepath))
            {
                string json = reader.ReadToEnd();
                JsonDocument jsonDoc = JsonDocument.Parse(json);

                foreach (var element in jsonDoc.RootElement.EnumerateArray())
                {
                    Int32  id    = element.GetProperty("IdServices"   ).GetInt32();
                    string name  = element.GetProperty("NameServices" ).GetString();
                    string type  = element.GetProperty("TypeOfService").GetString();
                    string code  = element.GetProperty("CodeService"  ).GetString();
                    double price = element.GetProperty("Cost"         ).GetDouble();

                    Service service = new Service();
                    service.Id    = id;
                    service.Name  = name;
                    service.Type  = type;
                    service.Code  = code;
                    service.Price = price;

                    if (dbServices.Length > 0)
                    {
                        bool found = false;
                        foreach (var dbService in dbServices)
                        {
                            if (dbService.Id == service.Id)
                            {
                                declinedServices += 1;
                                found = true;
                                break;
                            }

                        }

                        if (!found)
                        {
                            servicesToAdd.Add(service);
                            importedServices += 1;
                        }
                    }
                    else
                    {
                        servicesToAdd.Add(service);
                        importedServices += 1;
                    }

                    i += 1;
                }
            }

            int totalServices = i;
            db.Services.AddRange(servicesToAdd);
            db.SaveChangesAsync();
            MessageBox.Show(string.Format("Импортировано {0}/{1} сущностей.\n({2} сущностей совпадают с уже существующими)", importedServices, totalServices, declinedServices),
                            "Импорт JSON",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
        }

        public void ExportToExcel()
        {
            if (db.Services.FirstOrDefault() == null)
            {
                MessageBox.Show("Нет данных в базе данных! Пожалуйста, сначала импортируйте данные.",
                                "Экспорт Excel",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                return;
            }

            Excel.Application app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook book = app.Workbooks.Add(Type.Missing);

            List<Service> ctg1 = db.Services.Where(p => p.Price >= 0 && p.Price <= 350  ).ToList();
            List<Service> ctg2 = db.Services.Where(p => p.Price >= 250 && p.Price <= 800).ToList();
            List<Service> ctg3 = db.Services.Where(p => p.Price >= 800                  ).ToList();

            List<List<Service>> ctgs = new List<List<Service>>();
            ctgs.Add(ctg1);
            ctgs.Add(ctg2);
            ctgs.Add(ctg3);

            int sheetId = 1;
            Excel.Worksheet sheet;
            string[] sheetTitles = { "Категория №1", "Категория №2", "Категория №3" };
            foreach (var ctg in ctgs)
            {
                sheet = app.Sheets[sheetId];
                sheet.Name = sheetTitles[sheetId - 1];

                sheet.Cells[1, 1] = "Id";
                sheet.Cells[1, 2] = "Название услуги";
                sheet.Cells[1, 3] = "Вид услуги";
                sheet.Cells[1, 4] = "Стоимость";

                int i = 0;
                foreach (var item in ctg)
                {
                    sheet.Cells[i + 2, 1] = item.Id;
                    sheet.Cells[i + 2, 2] = item.Name;
                    sheet.Cells[i + 2, 3] = item.Type;
                    sheet.Cells[i + 2, 4] = item.Price;
                    i++;
                }

                Excel.Range rangeTitles = sheet.Range[sheet.Cells[1][1], sheet.Cells[4][1]];
                rangeTitles.Font.Size = 12;
                rangeTitles.Font.Bold = true;
                rangeTitles.Font.Name = "Times New Roman";

                Excel.Range rangeAll = sheet.Range[sheet.Cells[1][1], sheet.Cells[4][i + 1]];
                rangeAll.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangeAll.Borders[Excel.XlBordersIndex.xlEdgeBottom      ].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeAll.Borders[Excel.XlBordersIndex.xlEdgeLeft        ].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeAll.Borders[Excel.XlBordersIndex.xlEdgeTop         ].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeAll.Borders[Excel.XlBordersIndex.xlEdgeRight       ].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeAll.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeAll.Borders[Excel.XlBordersIndex.xlInsideVertical  ].LineStyle = Excel.XlLineStyle.xlContinuous;

                sheet.Columns.AutoFit();
                sheetId++;
            }

            app.Visible = true;
        }

        public void ExportToWord()
        {
            if (db.Services.FirstOrDefault() == null)
            {
                MessageBox.Show("Нет данных в базе данных! Пожалуйста, сначала импортируйте данные.",
                                "Экспорт Word",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                return;
            }

            List<Service> services = db.Services.ToList();

            List<Service> ctg1 = db.Services.Where(p => p.Price >= 0 && p.Price <= 350  ).ToList();
            List<Service> ctg2 = db.Services.Where(p => p.Price >= 250 && p.Price <= 800).ToList();
            List<Service> ctg3 = db.Services.Where(p => p.Price >= 800                  ).ToList();

            var word = new Word.Application();
            var doc  = word.Documents.Add();
            doc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
            AddCategoryToWordDocument(doc, "Категория №1", ctg1);
            AddCategoryToWordDocument(doc, "Категория №2", ctg2);
            AddCategoryToWordDocument(doc, "Категория №3", ctg3);

            word.Visible = true;
        }

        private void importExcelButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = projectFolder;
            dialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            dialog.FilterIndex = 1;
            dialog.RestoreDirectory = true;

            if (!(dialog.ShowDialog() == true))
                return;

            string filepath = dialog.FileName;
            Thread importThread = new Thread(() => {
                ImportFromExcel(filepath);
            });
            importThread.Start();
        }

        private void importJsonButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = projectFolder;
            dialog.Filter = "JSON files (*.json)|*.json";
            dialog.FilterIndex = 1;
            dialog.RestoreDirectory = true;

            if (!(dialog.ShowDialog() == true))
                return;

            string filepath = dialog.FileName;
            Thread importThread = new Thread(() => {
                ImportFromJson(filepath);
            });
            importThread.Start();
        }

        private void exportExcelButton_Click(object sender, RoutedEventArgs e)
        {
            Thread excelExportThread = new Thread(new ThreadStart(ExportToExcel));
            excelExportThread.Start();
        }

        private void AddCategoryToWordDocument(Word.Document doc, string title, List<Service> category)
        {
            var paragraph                    = doc.Paragraphs.Last;
            paragraph.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
            paragraph.Format.SpaceAfter      = 0;

            var range        = paragraph.Range;
            range.Text       = title;
            range.Font.Name  = "Times New Roman";
            range.Font.Size  = 14;
            range.Bold       = 1;
            range.Font.Color = Word.WdColor.wdColorBlack;
            range.InsertParagraphAfter();

            var table = doc.Paragraphs.Add();
            var serviceTable = doc.Tables.Add(table.Range, category.Count + 1, 4);
            serviceTable.Borders.InsideLineStyle = serviceTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            serviceTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            serviceTable.Cell(1, 1).Range.Text = "Id";
            serviceTable.Cell(1, 2).Range.Text = "Название услуги";
            serviceTable.Cell(1, 3).Range.Text = "Вид услуги";
            serviceTable.Cell(1, 4).Range.Text = "Стоимость";

            // serviceTable.Rows[1].Range.Bold = 1;
            // serviceTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            int i = 1;
            foreach (var service in category)
            {
                serviceTable.Cell(i + 1, 1).Range.Text = service.Id.ToString();
                serviceTable.Cell(i + 1, 2).Range.Text = service.Name;
                serviceTable.Cell(i + 1, 3).Range.Text = service.Type;
                serviceTable.Cell(i + 1, 4).Range.Text = service.Price.ToString();
                i++;
            }

            serviceTable.Range.Bold = 0;
            serviceTable.Rows[1].Range.Bold = 1;
            serviceTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
            serviceTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            range                                 = serviceTable.Range;
            range.ParagraphFormat.SpaceAfter      = 0.0f;
            range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

            paragraph = doc.Paragraphs.Add();
            paragraph = doc.Paragraphs.Add();

            range            = paragraph.Range;
            range.Text       = "Всего позиций: " + category.Count;
            range.Font.Name  = "Times New Roman";
            range.Font.Size  = 14;
            range.Bold       = 0;
            range.Font.Color = Word.WdColor.wdColorBlack;

            doc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
        }

        private void exportWordButton_Click(object sender, RoutedEventArgs e)
        {
            Thread wordExportThread = new Thread(new ThreadStart(ExportToWord));
            wordExportThread.Start();
        }

        private void showDbEntitiesButton_Click(object sender, RoutedEventArgs e)
        {
            var dbServices = db.Services.ToArray();
            var builder = new StringBuilder();
            builder.Append("Формат: [Id]: { \"Name\", \"Type\", \"Code\", Price }\n");
            int i = 0;
            for (; i < dbServices.Length; i++)
            {
                Service s = dbServices[i];
                string line = string.Format("[{0}]: {{ \"{1}\", \"{2}\", \"{3}\", {4} }}\n", i, s.Name, s.Type, s.Code, s.Price);
                builder.Append(line);
            }
            builder.Append(string.Format("(Выведено {0}/{1} сущностей)\n", i, dbServices.Length));
            MessageBox.Show(builder.ToString(),
                            "Просмотр данных БД",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
        }

        private void clearDbEntitiesButton_Click(object sender, RoutedEventArgs e)
        {
            var dbServices = db.Services.ToArray();
            int i = 0;
            for (; i < dbServices.Length; i++)
            {
                Service service = dbServices[i];
                db.Services.Remove(service);
            }

            db.SaveChangesAsync();
            MessageBox.Show(string.Format("Удалено {0}/{1} сущностей.", i, dbServices.Length),
                            "Удаление данных БД",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
        }
    }
}
