using Aspose.Words;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using static ECDIS_eGloebe___RouteConverter.Common.Common;

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

				SetFirstRows(worksheet);
				//worksheet.Cells[1, 1].Value = "Име";
				////worksheet.Cells[1, 2].Value = "Фамилия";
				//worksheet.Cells[1, 3].Value = "Години";

				//// Стилизиране на заглавките
				//using (ExcelRange range = worksheet.Cells["A2:C2"])
				//{
				//	range.Style.Font.Bold = true;
				//	range.Style.Fill.PatternType = ExcelFillStyle.Solid;
				//	range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
				//}

				// Примерни данни
				//	string[,] data = {
				//	{"Иван", "Петров", "30"},
				//	{"Мария", "Иванова", "25"},
				//	{"Петър", "Георгиев", "35"}
				//};

				//	// Попълване на данните в таблицата
				//	for (int i = 5; i < data.GetLength(0) + 5; i++)
				//	{
				//		for (int j = 0; j < data.GetLength(1); j++)
				//		{
				//			worksheet.Cells[i + 2, j + 1].Value = data[i - 5, j];
				//		}
				//	}



				// Добавяне на граници на клетки
				AddBorders(worksheet, "A1", "B3");

				// Set Print Area
				SetPrintArea(worksheet);

				// Запазване на Excel файл
				package.SaveAs(new FileInfo(xlsxFilePath));
			}


			MessageBox.Show("Таблицата е създадена успешно в XLSX.");
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
			worksheet.Row(3).Height = 16;
			worksheet.Row(4).Height = 14.25;
			worksheet.Row(5).Height = 15;
			worksheet.Row(6).Height = 15;
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

		private void SetFirstRows(ExcelWorksheet worksheet)
		{
			//FirstRow FirstCellHeader
			worksheet.Cells[1, 2].Value = @"Rev.  0" + Environment.NewLine + "Data rev. 28/02/2023";

			using (ExcelRange range = worksheet.Cells[$"B1:D1"])
			{
				range.Merge = true;
				range.Style.Fill.PatternType = ExcelFillStyle.Gray125;
				range.Style.Fill.BackgroundColor.Tint = 000000;
				range.Style.WrapText = true;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Gray);
				range.Style.Font.Size = 10;
				range.Style.Font.Name = "Times New Roman";
				range.Style.Font.Bold = true;
			}

			//FirstRow LastCellHeader
			worksheet.Cells[1, 22].Value = @"SMS-F-041";

			using (ExcelRange range = worksheet.Cells[$"V1:X1"])
			{
				range.Merge = true;
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
				range.Style.Fill.PatternType = ExcelFillStyle.Gray125;
				range.Style.Fill.BackgroundColor.Tint = 000000;
				range.Style.WrapText = true;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Gray);
				range.Style.Font.Size = 10;
				range.Style.Font.Name = "Times New Roman";
				range.Style.Font.Bold = true;
			}

			//FirstRow CenterCellHeader
			worksheet.Cells[1, 10].Value = @"Voyage Planning";

			using (ExcelRange range = worksheet.Cells[$"J1:O1"])
			{
				range.Merge = true;
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				range.Style.WrapText = true;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Gray);
				range.Style.Font.Size = 11;
				range.Style.Font.Name = "Arial";
				range.Style.Font.Bold = true;
			}

			//Third Row
			worksheet.Cells[3, 1].Value = @"SAFETY MANAGEMENT SYSTEM";

			using (ExcelRange range = worksheet.Cells[$"A3:Y3"])
			{
				range.Merge = true;
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				range.Style.WrapText = true;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
				range.Style.Font.Size = 12;
				range.Style.Font.Name = "Arial";
				range.Style.Font.Bold = true;
			}

			//Fourth Row
			worksheet.Cells[4, 8].Value = $"Ship’s Name {HomeInfo.VesselName}";

			using (ExcelRange range = worksheet.Cells[$"H4:R4"])
			{
				range.Merge = true;
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				range.Style.WrapText = true;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
				range.Style.Font.Size = 10;
				range.Style.Font.Name = "Times";
				range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
			}

			//Fifth Row
			worksheet.Cells[5, 2].Value =
				$"Voy nr. {HomeInfo.Voyage} From : {HomeInfo.PortFrom} To : {HomeInfo.PortTo}  " +
				$"Ets : {HomeInfo.Ets} Eta : {HomeInfo.Eta}   Safety contour/Safety depth : {HomeInfo.SafetyCountDepth}";

			using (ExcelRange range = worksheet.Cells[$"B5:Y5"])
			{
				range.Merge = true;
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				range.Style.WrapText = true;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
				range.Style.Font.Size = 10;
				range.Style.Font.Name = "Times";
			}

			//Sixth Row
			worksheet.Cells[6, 2].Value = $"Draft  : Fwd  {HomeInfo.draftFWD}m. Centre  {HomeInfo.draftMiddle}m.   Aft  {HomeInfo.draftAFT}m.";

			using (ExcelRange range = worksheet.Cells[$"B5:Y5"])
			{
				range.Merge = true;
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				range.Style.WrapText = true;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
				range.Style.Font.Size = 10;
				range.Style.Font.Name = "Times";
			}

			AddTableHeader(7, worksheet);
		}

		

		private void AddTableHeader(int initialRow, ExcelWorksheet worksheet)
		{
			// Default Style

			worksheet.Row(initialRow).Height = 24;
			worksheet.Row(initialRow + 1).Height = 19.5;
			worksheet.Row(initialRow + 2).Height = 105;

			using (ExcelRange range = worksheet.Cells[$"A{initialRow}:Y{initialRow + 2}"])
			{
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
				range.Style.Font.Size = 8;
				range.Style.Font.Name = "Arial Narrow";
				range.Style.TextRotation = 90;
				range.Style.Border.Top.Style = ExcelBorderStyle.Double;
				range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				range.Style.WrapText = true;
			}

			//A Column
			worksheet.Cells[initialRow, 1].Value = $"Way Points nr.";
			MergeCellsByIndex($"A{initialRow}", $"A{initialRow + 2}", worksheet);

			//B Column
			worksheet.Cells[initialRow, 2].Value = $"Position – Lat.Long.";
			MergeCellsByIndex($"B{initialRow}", $"B{initialRow + 2}", worksheet);

			//C Column
			worksheet.Cells[initialRow, 3].Value = $"Date";
			MergeCellsByIndex($"C{initialRow}", $"C{initialRow + 2}", worksheet);


			//D Column
			worksheet.Cells[initialRow, 4].Value = $"Course";
			MergeCellsByIndex($"D{initialRow}", $"D{initialRow + 2}", worksheet);

			//E Column
			worksheet.Cells[initialRow, 5].Value = $"Estimated Speed (safe speed)";
			MergeCellsByIndex($"E{initialRow}", $"E{initialRow + 2}", worksheet);

			//F Column
			worksheet.Cells[initialRow, 6].Value = $"Miles between wayponts";
			MergeCellsByIndex($"F{initialRow}", $"F{initialRow + 2}", worksheet);

			//G Column
			worksheet.Cells[initialRow, 7].Value = $"Miles to final destination";
			MergeCellsByIndex($"G{initialRow}", $"G{initialRow + 2}", worksheet);

			//H Column
			worksheet.Cells[initialRow, 8].Value = $"Chart to be used nr.";
			MergeCellsByIndex($"H{initialRow}", $"H{initialRow + 2}", worksheet);

			//I Column
			worksheet.Cells[initialRow, 9].Value = $"Ship’s reporting System (Ares, etc.)";
			MergeCellsByIndex($"I{initialRow}", $"I{initialRow + 2}", worksheet);

			//J Column
			worksheet.Cells[initialRow, 10].Value = $"VTS station and  VHF Channel";
			MergeCellsByIndex($"J{initialRow}", $"J{initialRow + 2}", worksheet);

			//K Column
			worksheet.Cells[initialRow, 11].Value = $"Marpol special areas Yes/No";
			MergeCellsByIndex($"K{initialRow}", $"K{initialRow + 2}", worksheet);

			//L Column
			worksheet.Cells[initialRow, 12].Value = $"Squat effect";
			MergeCellsByIndex($"L{initialRow}", $"L{initialRow + 2}", worksheet);

			//M Column
			worksheet.Cells[initialRow, 13].Value = $"Minimum underkeel clearance  included squat effect";
			MergeCellsByIndex($"M{initialRow}", $"M{initialRow + 2}", worksheet);

			//N Column
			worksheet.Cells[initialRow, 14].Value = $"Safe distance from obstacles";
			MergeCellsByIndex($"N{initialRow}", $"N{initialRow + 2}", worksheet);

			//O Column
			worksheet.Cells[initialRow, 15].Value = $"Ship positioning system (Gps, etc.)";
			MergeCellsByIndex($"O{initialRow}", $"O{initialRow + 2}", worksheet);

			//P Column
			worksheet.Cells[initialRow, 16].Value = $"Current Direction/Speed";
			MergeCellsByIndex($"P{initialRow}", $"P{initialRow + 2}", worksheet);

			//Q Column
			worksheet.Cells[initialRow, 17].Value = $"Stand by in the engine Yes/No";
			MergeCellsByIndex($"Q{initialRow}", $"Q{initialRow + 2}", worksheet);

			//R:V Column
			worksheet.Cells[initialRow, 18].Value = $"From pilot station to the key or  v.v.";
			MergeCellsByIndex($"R{initialRow}", $"V{initialRow}", worksheet);
			using (ExcelRange range = worksheet.Cells[$"R{initialRow}:V{initialRow}"])
			{
				range.Style.TextRotation = 0;
			}

			//R:T Column
			worksheet.Cells[initialRow + 1, 18].Value = $"Tide time";
			MergeCellsByIndex($"R{initialRow + 1}", $"T{initialRow + 1}", worksheet);
			using (ExcelRange range = worksheet.Cells[$"R{initialRow + 1}:T{initialRow + 1}"])
			{
				range.Style.TextRotation = 0;
			}

			//R3 Column
			worksheet.Cells[initialRow + 2, 18].Value = $"Low";

			//S3 Column
			worksheet.Cells[initialRow + 2, 19].Value = $"High";

			//T3 Column
			worksheet.Cells[initialRow + 2, 20].Value = $"Slack";

			//U2 Column
			worksheet.Cells[initialRow + 1, 21].Value = $"Miles pilot/berth (or  v.v.)";
			MergeCellsByIndex($"U{initialRow + 1}", $"U{initialRow + 2}", worksheet);

			//V2 Column
			worksheet.Cells[initialRow + 1, 22].Value = $"Miles to the lock";
			MergeCellsByIndex($"V{initialRow + 1}", $"V{initialRow + 2}", worksheet);

			using (ExcelRange range = worksheet.Cells[$"R{initialRow + 1}:V{initialRow + 2}"])
			{
				range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			}

			//W Column
			worksheet.Cells[initialRow, 23].Value = $"Refuge port of emergency anchorage";
			MergeCellsByIndex($"W{initialRow}", $"W{initialRow + 2}", worksheet);

			//X Column
			worksheet.Cells[initialRow, 24].Value = $"Nautical publications to be used";
			MergeCellsByIndex($"X{initialRow}", $"X{initialRow + 2}", worksheet);

			//Y Column
			worksheet.Cells[initialRow, 25].Value = $"Notes";
			MergeCellsByIndex($"Y{initialRow}", $"Y{initialRow + 2}", worksheet);
		}

		private void MergeCellsByIndex(string startCell, string endCell, ExcelWorksheet worksheet)
		{
			worksheet.Cells[startCell + ":" + endCell].Merge = true;
		}

		private void SetPrintArea(ExcelWorksheet worksheet)
		{
			int maxColumnIndex = worksheet.Dimension.End.Column;
			int maxRowIndex = worksheet.Dimension.End.Row;
			string printArea = $"A1:{ExcelColumnFromNumber(maxColumnIndex)}{maxRowIndex}";
			worksheet.PrinterSettings.PrintArea = worksheet.Cells[printArea];

			// Задаване на реда, който да се повтаря при принтиране на всяка страница
			worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:1"];

			worksheet.PrinterSettings.Orientation = eOrientation.Landscape;
			worksheet.PrinterSettings.Scale = 80;
		}
		// Метод за преобразуване на номера на колона към буква
		private string ExcelColumnFromNumber(int column)
		{
			int dividend = column;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0)
			{
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}
	}
}
