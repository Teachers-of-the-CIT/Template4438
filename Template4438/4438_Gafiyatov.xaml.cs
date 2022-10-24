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
    }
}
