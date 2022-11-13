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
using Newtonsoft.Json;
using System.IO;
using System.Globalization;
using Word = Microsoft.Office.Interop.Word;

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

			try
			{
				using (ServicesTableEntities servicesTableEntities = new ServicesTableEntities())
				{
					for (int i = 1; i < _rows; i++)
					{
						servicesTableEntities.Services.Add(new Services()
						{
							Name = list[i, 1],
							KindService = list[i, 2],
							CodeService = list[i, 3],
							Cost = Convert.ToInt32(list[i, 4])
						});
					}

					servicesTableEntities.SaveChanges();
				}
			}
			catch
			{
				MessageBox.Show("При попытке сохранения информации в базе данных возникла ошибка.");
			}
			
			MessageBox.Show("Импорт данных выполнен успешно.");
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

			MessageBox.Show("Экспорт выполнен успешно.");
		}

		/// <summary>
		/// Импортировать данные из JSON файла в БД.
		/// </summary>
		private void ImportJson_Click(object sender, RoutedEventArgs e)
		{
			List<Classes.SecondServicesModel> secondServices = new List<Classes.SecondServicesModel>();

			string fileName = "../../../Data.json";
			string jsonString = File.ReadAllText(fileName);

			secondServices = JsonConvert.DeserializeObject<List<Classes.SecondServicesModel>>(jsonString);

			AddDataToDb(secondServices);
		}

		private void AddDataToDb(List<Classes.SecondServicesModel> secondServicesModels)
		{
			var cultureInfo = new CultureInfo("de-DE");
			DateTime? dt = null;

			using (var entity = new ServicesTableEntities())
			{
				SecondServices secondServices;

				foreach (var item in secondServicesModels)
				{
					secondServices = new SecondServices();

					secondServices.Id = item.Id;
					secondServices.OrderCode = item.CodeOrder;
					string date = item.CreateDate.ToString();
					secondServices.CreateDate = DateTime.Parse(date, cultureInfo);
					secondServices.OrderTime = item.CreateTime;
					secondServices.UserCode = Convert.ToInt32(item.CodeClient);
					secondServices.NumberServices = item.Services;
					secondServices.Status = item.Status;
					secondServices.CloseDate = String.IsNullOrEmpty(item.ClosedDate) ? dt : DateTime.Parse(item.ClosedDate, cultureInfo);
					secondServices.RentalTime = item.ProkatTime;
					
					entity.SecondServices.Add(secondServices);
				}

				entity.SaveChanges();
				MessageBox.Show("Импортирование данных из Json файла выполнено успешно");
			}
		}

		private int GetCountListElements(List<Services> services, int i)
		{
			int count = 0;

			foreach (var item in services)
			{
				if (item.Cost > costCategory[i] && item.Cost < costCategory[i + 1])
				{
					count++;
				}
			}

			return count;
		}

		public int[] costCategory = { 0, 350, 800, 2000 };

		private void ExportWord_Click(object sender, RoutedEventArgs e)
		{
			var Services = new List<Services>();

			using (var entity = new ServicesTableEntities())
			{
				Services = entity.Services.ToList();

				var app = new Word.Application();
				Word.Document document = app.Documents.Add();

				

				for (int i = 0; i < costCategory.Count() - 1; i++)
				{
					Word.Paragraph paragraph = document.Paragraphs.Add();
					Word.Range range = paragraph.Range;

					range.Text = "Группа " + (i + 1);
					range.set_Style("Заголовок 1");
					range.InsertParagraphAfter();

					Word.Paragraph tableParagrath = document.Paragraphs.Add();
					Word.Range tableRange = tableParagrath.Range;

					Word.Table serviceTable = document.Tables.Add(tableRange, GetCountListElements(Services, i) + 1, 5);

					serviceTable.Borders.InsideLineStyle = serviceTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
					serviceTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

					Word.Range cellRange = serviceTable.Cell(1, 1).Range;
					cellRange.Text = "Идентификатор";

					cellRange = serviceTable.Cell(1, 2).Range;
					cellRange.Text = "Наименование";

					cellRange = serviceTable.Cell(1, 3).Range;
					cellRange.Text = "Вид услуги";

					cellRange = serviceTable.Cell(1, 4).Range;
					cellRange.Text = "Код услуги";

					cellRange = serviceTable.Cell(1, 5).Range;
					cellRange.Text = "Стоимость";

					serviceTable.Rows[1].Range.Bold = 1;
					serviceTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

					int index = 1;

					foreach (var item in Services)
					{
						if (item.Cost > costCategory[i] && item.Cost < costCategory[i + 1])
						{
							cellRange = serviceTable.Cell(index + 1, 1).Range;
							cellRange.Text = item.Id.ToString();
							cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

							cellRange = serviceTable.Cell(index + 1, 2).Range;
							cellRange.Text = item.Name.ToString();
							cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

							cellRange = serviceTable.Cell(index + 1, 3).Range;
							cellRange.Text = item.KindService.ToString();
							cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

							cellRange = serviceTable.Cell(index + 1, 4).Range;
							cellRange.Text = item.CodeService.ToString();
							cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

							cellRange = serviceTable.Cell(index + 1, 5).Range;
							cellRange.Text = item.Cost.ToString();
							cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

							index++;
						}
					}

					app.Visible = true;
					//document.SaveAs2(@"D:\outputFileWord.docx");
					//document.SaveAs2(@"D:\outputFileWord.pdf", Word.WdExportFormat.wdExportFormatPDF);
				}
			}
		}
	}
}
