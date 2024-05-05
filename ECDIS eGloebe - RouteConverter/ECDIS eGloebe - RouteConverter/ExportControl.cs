using Aspose.Words.Tables;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using ECDIS_eGloebe___RouteConverter.ExportRoute;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace ECDIS_eGloebe___RouteConverter
{
	public partial class ExportControl : UserControl
	{
		SaveFileDialog saveFileDialog = new SaveFileDialog()
		{
			// Задаване на филтър за разширения на файловете
			Filter = "Текстови файлове (*.txt)|*.txt|Всички файлове (*.*)|*.*",
			FilterIndex = 1,
			RestoreDirectory = true
		};

		string xmlFilePath = "eGlobeTemplate-Excel.xml";
		string xlsxFilePath = "test.xlsx";



		public ExportControl()
		{
			InitializeComponent();
		}

		private void btnExportHomeInfo_Click(object sender, EventArgs e)
		{
			if (File.Exists(xmlFilePath))
			{
				File.Delete(xmlFilePath);
			}
			// Създаване на нов Excel пакет с помощта на EPPlus
			using (ExcelPackage package = new ExcelPackage())
			{
				// Добавяне на нов лист
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Данни");

				// Автоматично настройване на ширината на колоните
				worksheet.Cells.AutoFitColumns();

				// Задаване на ширина на колоните
				SetColumnWidth(worksheet);
				SetRowHeight(worksheet);

				// Добавяне на заглавки на колоните

				SetFirstRow(worksheet);
				//worksheet.Cells[1, 1].Value = "Име";
				////worksheet.Cells[1, 2].Value = "Фамилия";
				//worksheet.Cells[1, 3].Value = "Години";

				// Стилизиране на заглавките
				using (ExcelRange range = worksheet.Cells["A2:C2"])
				{
					range.Style.Font.Bold = true;
					range.Style.Fill.PatternType = ExcelFillStyle.Solid;
					range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
				}

				// Примерни данни
				string[,] data = {
				{"Иван", "Петров", "30"},
				{"Мария", "Иванова", "25"},
				{"Петър", "Георгиев", "35"}
			};

				// Попълване на данните в таблицата
				for (int i = 5; i < data.GetLength(0) + 5; i++)
				{
					for (int j = 0; j < data.GetLength(1); j++)
					{
						worksheet.Cells[i + 2, j + 1].Value = data[i - 5, j];
					}
				}



				// Добавяне на граници на клетки
				AddBorders(worksheet, "A1", "B3");

				// Обединяване на клетките
				worksheet.Cells["A5:B5"].Merge = true;

				// Запазване на Excel файл
				package.SaveAs(new FileInfo(xlsxFilePath));
			}


			MessageBox.Show("Таблицата е създадена успешно в XLSX.");

			File.Open(xlsxFilePath, FileMode.Open);
		}

		private static void AddBorders(ExcelWorksheet worksheet, string cell1, string cell2)
		{
			using (ExcelRange range = worksheet.Cells[$"{cell1}:{cell2}"])
			{
				range.Style.Border.Top.Style = ExcelBorderStyle.Double;
				range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
			}
		}

		//private void SetFirstRow(ExcelWorksheet worksheet)
		//{
		//	worksheet.Cells[1, 2].Value = "Rev.  0" + Environment.NewLine + "Data rev. 28/02/2023";
		//}

		private void SetColumnWidth(ExcelWorksheet worksheet)
		{
			worksheet.Column(1).Width = 3;
			worksheet.Column(2).Width = 16.17;
			worksheet.Column(3).Width = 6.67;
			worksheet.Column(4).Width = 6.5;
			worksheet.Column(5).Width = 5.5;
			worksheet.Column(6).Width = 5.5;
			worksheet.Column(7).Width = 3;
			worksheet.Column(8).Width = 8.5;
			worksheet.Column(9).Width = 8.5;
			worksheet.Column(10).Width = 6;
			worksheet.Column(11).Width = 3.67;
			worksheet.Column(12).Width = 6;
			worksheet.Column(13).Width = 4.83;
			worksheet.Column(14).Width = 6.17;
			worksheet.Column(15).Width = 3;
			worksheet.Column(16).Width = 8.17;
			worksheet.Column(17).Width = 2.83;
			worksheet.Column(18).Width = 3.83;
			worksheet.Column(19).Width = 3.67;
			worksheet.Column(20).Width = 3.33;
			worksheet.Column(21).Width = 5.5;
			worksheet.Column(22).Width = 5.33;
			worksheet.Column(23).Width = 9.33;
			worksheet.Column(24).Width = 12;
			worksheet.Column(25).Width = 10.17;
		}

		private void SetRowHeight(ExcelWorksheet worksheet)
		{
			worksheet.Row(1).Height = 28.5;
			//worksheet.Row(2).Height = 16.17;
			//worksheet.Row(3).Height = 6.67;
			//worksheet.Row(4).Height = 6.5;
			//worksheet.Row(5).Height = 5.5;
			//worksheet.Row(6).Height = 5.5;
			//worksheet.Row(7).Height = 3;
			//worksheet.Row(8).Height = 8.5;
			//worksheet.Row(9).Height = 8.5;
			//worksheet.Row(10).Height = 6;
			//worksheet.Row(11).Height = 3.67;
			//worksheet.Row(12).Height = 6;
			//worksheet.Row(13).Height = 4.83;
			//worksheet.Row(14).Height = 6.17;
			//worksheet.Row(15).Height = 3;
			//worksheet.Row(16).Height = 8.17;
			//worksheet.Row(17).Height = 2.83;
			//worksheet.Row(18).Height = 3.83;
			//worksheet.Row(19).Height = 3.67;
			//worksheet.Row(20).Height = 3.33;
			//worksheet.Row(21).Height = 5.5;
			//worksheet.Row(22).Height = 5.33;
			//worksheet.Row(23).Height = 9.33;
			//worksheet.Row(24).Height = 12;
			//worksheet.Row(25).Height = 10.17;
		}

		private void SetFirstRow(ExcelWorksheet worksheet)
		{
			worksheet.Cells[1, 2].Value = @"Rev.  0" + Environment.NewLine + "Data rev. 28/02/2023";

			MergeCellsByIndex("B1", "D1", worksheet);

			using (ExcelRange range = worksheet.Cells[$"B1:D1"])
			{
				range.Style.Fill.PatternType = ExcelFillStyle.Gray125;
				range.Style.Fill.BackgroundColor.Tint = 000000;
				range.Style.WrapText = true;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Gray);
				range.Style.Font.Size = 10;
				range.Style.Font.Name = "Times New Roman";
				range.Style.Font.Bold = true;
			}
		}

		private void MergeCellsByIndex(string startCell, string endCell, ExcelWorksheet worksheet)
		{
			worksheet.Cells[startCell + ":" + endCell].Merge = true;
		}
	}
}
