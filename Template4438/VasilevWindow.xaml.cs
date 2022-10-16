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
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4438
{
	public partial class VasilevWindow : Window
	{
		public static MainWindow mainWindow;
		public VasilevWindow()
		{
			InitializeComponent();
		}

		private void Back_Click(object sender, RoutedEventArgs e)
		{
			if (mainWindow == null)
			{ 
				mainWindow = new MainWindow();
				mainWindow.Show();
				this.Close();
			}
		}

		private void ImportData_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog ofd = new OpenFileDialog()
			{
				DefaultExt = "*.xls;*.xlsx",
				Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
				Title = "Выберите файл базы данных"
			};

			if (!(ofd.ShowDialog() == true))
			{
				return;
			}

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

			ObjWorkBook.Close(false, Type.Missing, Type.Missing);
			ObjWorkExcel.Quit();

			GC.Collect();

			using (ServicesTableEntities servicesTableEntities = new ServicesTableEntities())
			{
				for (int i = 1; i < _rows; i++)
				{
					servicesTableEntities.Services.Add(new Services() {
						Name = list[i, 1],
						KindService = list[i, 2],
						CodeService = list[i, 3],
						Cost = Convert.ToInt32(list[i, 4])
					});
				}

				servicesTableEntities.SaveChanges();
			}	
		}

		private void ExportData_Click(object sender, RoutedEventArgs e)
		{
			List<Services> allServices;

			using (ServicesTableEntities servicesTableEntities = new ServicesTableEntities())
			{
				allServices = servicesTableEntities.Services.ToList().OrderBy(s => s.Cost).ToList();
			}

			const int CountCategory = 3;

			var app = new Excel.Application();
			app.SheetsInNewWorkbook = CountCategory;
			Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

			int[] array = { 0, 350, 800, 2000};

			for (int i = 0; i < CountCategory; i++)
			{
				Excel.Worksheet worksheet = app.Worksheets[i + 1];
				worksheet.Name = "Kategory_" + (i + 1);

				int rowIndex = 1;

				worksheet.Cells[1][rowIndex] = "Id";
				worksheet.Cells[2][rowIndex] = "Название услуги";
				worksheet.Cells[3][rowIndex] = "Вид услуги";
				worksheet.Cells[4][rowIndex] = "Стоимость";

				foreach (var item in allServices)
				{
					if (item.Cost > array[i] && item.Cost <= array[i + 1])
					{
						rowIndex++;

						worksheet.Cells[1][rowIndex] = Convert.ToString(item.Id);
						worksheet.Cells[2][rowIndex] = item.Name;
						worksheet.Cells[3][rowIndex] = item.KindService;
						worksheet.Cells[4][rowIndex] = item.Cost;
					}
				}
			}

			app.Visible = true;

			workbook.SaveAs(@"D:\outputFileExcel.xlsx");

			MessageBox.Show("Экспорт выполнен успешно!");
		}
	}
}
