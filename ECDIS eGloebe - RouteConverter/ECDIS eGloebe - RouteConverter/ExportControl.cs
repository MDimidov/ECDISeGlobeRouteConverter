//using Aspose.Words;
//using Aspose.Words.Tables;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using ECDIS_eGloebe___RouteConverter.DTOs;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Linq;
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

			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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


				// Задаване брояч на текущия ред

				SetFirstRows(worksheet);

				if (RouteDto.Waipoints.Any())
				{
					for (int i = 0; i < RouteDto.Waipoints.Count; i++)
					{
						AddWayPoint(worksheet, RouteDto.Waipoints[i], i + 1);
					}
				}

				// Добавяне на граници на клетки
				//AddBorders(worksheet, "A1", "B3");

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
			worksheet.Column(7).Width = 6.86;
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
			worksheet.Cells[++rowNo, 2].Value = @"Rev.  0" + Environment.NewLine + "Data rev. 28/02/2023";
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
			worksheet.Cells[rowNo, 22].Value = @"SMS-F-041";
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
			worksheet.Cells[rowNo++, 10].Value = @"Voyage Planning";
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
			worksheet.Cells[++rowNo, 1].Value = @"SAFETY MANAGEMENT SYSTEM";
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
			worksheet.Cells[++rowNo, 8].Value = $"Ship’s Name {HomeInfo.VesselName}";
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
			worksheet.Cells[++rowNo, 2].Value =
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
			worksheet.Cells[++rowNo, 2].Value = $"Draft  : Fwd  {HomeInfo.draftFWD}m. Centre  {HomeInfo.draftMiddle}m.   Aft  {HomeInfo.draftAFT}m.";
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

			AddTableHeader(worksheet);
		}

		

		private void AddTableHeader(ExcelWorksheet worksheet)
		{
			// Default Style

			worksheet.Row(rowNo).Height = 24;
			worksheet.Row(rowNo + 1).Height = 19.5;
			worksheet.Row(rowNo + 2).Height = 105;

			using (ExcelRange range = worksheet.Cells[$"A{rowNo}:Y{rowNo + 2}"])
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
			worksheet.Cells[rowNo, 1].Value = $"Way Points nr.";
			MergeCellsByIndex($"A{rowNo}", $"A{rowNo + 2}", worksheet);

			//B Column
			worksheet.Cells[rowNo, 2].Value = $"Position – Lat.Long.";
			MergeCellsByIndex($"B{rowNo}", $"B{rowNo + 2}", worksheet);

			//C Column
			worksheet.Cells[rowNo, 3].Value = $"Date";
			MergeCellsByIndex($"C{rowNo}", $"C{rowNo + 2}", worksheet);


			//D Column
			worksheet.Cells[rowNo, 4].Value = $"Course";
			MergeCellsByIndex($"D{rowNo}", $"D{rowNo + 2}", worksheet);

			//E Column
			worksheet.Cells[rowNo, 5].Value = $"Estimated Speed (safe speed)";
			MergeCellsByIndex($"E{rowNo}", $"E{rowNo + 2}", worksheet);

			//F Column
			worksheet.Cells[rowNo, 6].Value = $"Miles between wayponts";
			MergeCellsByIndex($"F{rowNo}", $"F{rowNo + 2}", worksheet);

			//G Column
			worksheet.Cells[rowNo, 7].Value = $"Miles to final destination";
			MergeCellsByIndex($"G{rowNo}", $"G{rowNo + 2}", worksheet);

			//H Column
			worksheet.Cells[rowNo, 8].Value = $"Chart to be used nr.";
			MergeCellsByIndex($"H{rowNo}", $"H{rowNo + 2}", worksheet);

			//I Column
			worksheet.Cells[rowNo, 9].Value = $"Ship’s reporting System (Ares, etc.)";
			MergeCellsByIndex($"I{rowNo}", $"I{rowNo + 2}", worksheet);

			//J Column
			worksheet.Cells[rowNo, 10].Value = $"VTS station and  VHF Channel";
			MergeCellsByIndex($"J{rowNo}", $"J{rowNo + 2}", worksheet);

			//K Column
			worksheet.Cells[rowNo, 11].Value = $"Marpol special areas Yes/No";
			MergeCellsByIndex($"K{rowNo}", $"K{rowNo + 2}", worksheet);

			//L Column
			worksheet.Cells[rowNo, 12].Value = $"Squat effect";
			MergeCellsByIndex($"L{rowNo}", $"L{rowNo + 2}", worksheet);

			//M Column
			worksheet.Cells[rowNo, 13].Value = $"Minimum underkeel clearance  included squat effect";
			MergeCellsByIndex($"M{rowNo}", $"M{rowNo + 2}", worksheet);

			//N Column
			worksheet.Cells[rowNo, 14].Value = $"Safe distance from obstacles";
			MergeCellsByIndex($"N{rowNo}", $"N{rowNo + 2}", worksheet);

			//O Column
			worksheet.Cells[rowNo, 15].Value = $"Ship positioning system (Gps, etc.)";
			MergeCellsByIndex($"O{rowNo}", $"O{rowNo + 2}", worksheet);

			//P Column
			worksheet.Cells[rowNo, 16].Value = $"Current Direction/Speed";
			MergeCellsByIndex($"P{rowNo}", $"P{rowNo + 2}", worksheet);

			//Q Column
			worksheet.Cells[rowNo, 17].Value = $"Stand by in the engine Yes/No";
			MergeCellsByIndex($"Q{rowNo}", $"Q{rowNo + 2}", worksheet);

			//R:V Column
			worksheet.Cells[rowNo, 18].Value = $"From pilot station to the key or  v.v.";
			MergeCellsByIndex($"R{rowNo}", $"V{rowNo}", worksheet);
			using (ExcelRange range = worksheet.Cells[$"R{rowNo}:V{rowNo}"])
			{
				range.Style.TextRotation = 0;
			}

			//R:T Column
			worksheet.Cells[rowNo + 1, 18].Value = $"Tide time";
			MergeCellsByIndex($"R{rowNo + 1}", $"T{rowNo + 1}", worksheet);
			using (ExcelRange range = worksheet.Cells[$"R{rowNo + 1}:T{rowNo + 1}"])
			{
				range.Style.TextRotation = 0;
			}

			//R3 Column
			worksheet.Cells[rowNo + 2, 18].Value = $"Low";

			//S3 Column
			worksheet.Cells[rowNo + 2, 19].Value = $"High";

			//T3 Column
			worksheet.Cells[rowNo + 2, 20].Value = $"Slack";

			//U2 Column
			worksheet.Cells[rowNo + 1, 21].Value = $"Miles pilot/berth (or  v.v.)";
			MergeCellsByIndex($"U{rowNo + 1}", $"U{rowNo + 2}", worksheet);

			//V2 Column
			worksheet.Cells[rowNo + 1, 22].Value = $"Miles to the lock";
			MergeCellsByIndex($"V{rowNo + 1}", $"V{rowNo + 2}", worksheet);

			using (ExcelRange range = worksheet.Cells[$"R{rowNo + 1}:V{rowNo + 2}"])
			{
				range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			}

			//W Column
			worksheet.Cells[rowNo, 23].Value = $"Refuge port of emergency anchorage";
			MergeCellsByIndex($"W{rowNo}", $"W{rowNo + 2}", worksheet);

			//X Column
			worksheet.Cells[rowNo, 24].Value = $"Nautical publications to be used";
			MergeCellsByIndex($"X{rowNo}", $"X{rowNo + 2}", worksheet);

			//Y Column
			worksheet.Cells[rowNo, 25].Value = $"Notes";
			MergeCellsByIndex($"Y{rowNo}", $"Y{rowNo + 2}", worksheet);

			rowNo += 3;
		}

		private void AddWayPoint(ExcelWorksheet worksheet, ImportWaipointDto wp, int wpNo)
		{
			// Default Style

			worksheet.Row(rowNo).Height = 20.25;
			worksheet.Row(rowNo + 1).Height = 20.25;

			using (ExcelRange range = worksheet.Cells[$"A{rowNo}:Y{rowNo + 1}"])
			{
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
				range.Style.Font.Size = 8;
				range.Style.Font.Name = "Arial Narrow";
				range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
				range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
				range.Style.WrapText = true;
			}

			//A Column
			worksheet.Cells[rowNo, 1].Value = $"{wpNo}";
			MergeCellsByIndex($"A{rowNo}", $"A{rowNo + 1}", worksheet);

			//B1 Column
			worksheet.Cells[rowNo, 2].Value = $"{wp.Position.LatDegrees}° {wp.Position.LatMinutes.ToString("F3")}' {wp.Position.LatDir}";
			MergeCellsByIndex($"B{rowNo}", $"B{rowNo}", worksheet);

			//B2 Column
			worksheet.Cells[rowNo + 1, 2].Value = $"{wp.Position.LongDegrees}° {wp.Position.LongMinutes.ToString("F3")}' {wp.Position.LongDir}";
			MergeCellsByIndex($"B{rowNo + 1}", $"B{rowNo + 1}", worksheet);

			//C Column
			worksheet.Cells[rowNo, 3].Value = $"##Date##";
			MergeCellsByIndex($"C{rowNo}", $"C{rowNo + 1}", worksheet);


			//D Column
			worksheet.Cells[rowNo, 4].Value = $"{wp.Course}°";
			MergeCellsByIndex($"D{rowNo}", $"D{rowNo + 1}", worksheet);

			//E Column
			worksheet.Cells[rowNo, 5].Value = $"##Estimated Speed##";
			MergeCellsByIndex($"E{rowNo}", $"E{rowNo + 1}", worksheet);

			//F Column
			worksheet.Cells[rowNo, 6].Value = $"{wp.DistanceFromLastWp}";
			MergeCellsByIndex($"F{rowNo}", $"F{rowNo + 1}", worksheet);

			//G Column
			worksheet.Cells[rowNo, 7].Value = $"{wp.DistanceToGo}";
			MergeCellsByIndex($"G{rowNo}", $"G{rowNo + 1}", worksheet);

			//H Column
			worksheet.Cells[rowNo, 8].Value = $"##Chart to be used##";
			MergeCellsByIndex($"H{rowNo}", $"H{rowNo + 1}", worksheet);

			//I Column
			worksheet.Cells[rowNo, 9].Value = $"##Ship’s reporting System##";
			MergeCellsByIndex($"I{rowNo}", $"I{rowNo + 1}", worksheet);

			//J Column
			worksheet.Cells[rowNo, 10].Value = $"##VTS station and VHF##";
			MergeCellsByIndex($"J{rowNo}", $"J{rowNo + 1}", worksheet);

			//K Column
			worksheet.Cells[rowNo, 11].Value = $"No";
			MergeCellsByIndex($"K{rowNo}", $"K{rowNo + 1}", worksheet);

			//L Column
			worksheet.Cells[rowNo, 12].Value = $"##wp1Squat##";
			MergeCellsByIndex($"L{rowNo}", $"L{rowNo + 1}", worksheet);

			//M Column
			worksheet.Cells[rowNo, 13].Value = $"##wp1UKC##";
			MergeCellsByIndex($"M{rowNo}", $"M{rowNo + 1}", worksheet);

			//N Column
			worksheet.Cells[rowNo, 14].Value = $"##wp1SafeDist##";
			MergeCellsByIndex($"N{rowNo}", $"N{rowNo + 1}", worksheet);

			//O Column
			worksheet.Cells[rowNo, 15].Value = $"GPS";
			MergeCellsByIndex($"O{rowNo}", $"O{rowNo + 1}", worksheet);

			//P Column
			worksheet.Cells[rowNo, 16].Value = $"##wp1CurrentDir##";
			MergeCellsByIndex($"P{rowNo}", $"P{rowNo + 1}", worksheet);

			//Q Column
			worksheet.Cells[rowNo, 17].Value = $"Yes";
			MergeCellsByIndex($"Q{rowNo}", $"Q{rowNo + 1}", worksheet);

			//R Column
			worksheet.Cells[rowNo, 18].Value = $"##wp1TideLow##";
			MergeCellsByIndex($"R{rowNo}", $"R{rowNo + 1}", worksheet);


			//S Column
			worksheet.Cells[rowNo, 19].Value = $"##wp1TideHigh##";
			MergeCellsByIndex($"S{rowNo}", $"S{rowNo + 1}", worksheet);

			//T Column
			worksheet.Cells[rowNo, 20].Value = $"##wp1TideSlack##";
			MergeCellsByIndex($"T{rowNo}", $"T{rowNo + 1}", worksheet);

			//U Column
			worksheet.Cells[rowNo + 1, 21].Value = $"##wp1PilotMiles##";
			MergeCellsByIndex($"U{rowNo}", $"U{rowNo + 1}", worksheet);

			//V Column
			worksheet.Cells[rowNo + 1, 22].Value = $"##wp1MilesToLock##";
			MergeCellsByIndex($"V{rowNo}", $"V{rowNo + 1}", worksheet);

			//W Column
			worksheet.Cells[rowNo, 23].Value = $"##wp1EmAnchorage##";
			MergeCellsByIndex($"W{rowNo}", $"W{rowNo + 1}", worksheet);

			//X Column
			worksheet.Cells[rowNo, 24].Value = $"##wp1NauticalPublications##";
			MergeCellsByIndex($"X{rowNo}", $"X{rowNo + 1}", worksheet);

			//Y Column
			worksheet.Cells[rowNo, 25].Value = $"##wp1Notes##";
			MergeCellsByIndex($"Y{rowNo}", $"Y{rowNo + 1}", worksheet);

			rowNo += 2;
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
