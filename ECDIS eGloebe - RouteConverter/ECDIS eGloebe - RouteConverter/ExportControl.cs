//using Aspose.Words;
//using Aspose.Words.Tables;
using Aspose.Words;
using Aspose.Words.Replacing;
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

						if (i == 7 || (RouteDto.Waipoints.Count - 1 < 7 && i == RouteDto.Waipoints.Count - 1))
						{
							AddNameSide(worksheet);
							InsertPageBreak(worksheet);
							AddEmptyRow(worksheet);
							AddTableHeader(worksheet);
						}
						else if (((i + 3) % 10 == 0) 
								&& i > 9 
								&& RouteDto.Waipoints.Count - 1 != i)
						{
							AddNumberPage(worksheet);
							InsertPageBreak(worksheet);
							AddEmptyRow(worksheet);
							AddTableHeader(worksheet);
						}
						else if(RouteDto.Waipoints.Count - 1 == i)
						{
							AddLastBorder(worksheet);
							AddNumberPage(worksheet);
						}
					}
				}

				// Set Print Area
				SetPrintArea(worksheet);

				ReplaceText(worksheet, "C1", $"E{RowNo}", "##lastPage##", PageNo.ToString());

				// Запазване на Excel файл
				package.SaveAs(new FileInfo(xlsxFilePath));
			}


			MessageBox.Show("Таблицата е създадена успешно в XLSX.");
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
			worksheet.Column(15).Width = 4;
			worksheet.Column(16).Width = 8.17;
			worksheet.Column(17).Width = 4;
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
			worksheet.Cells[++RowNo, 2].Value = @"Rev.  0" + Environment.NewLine + "Data rev. 28/02/2023";
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
			worksheet.Cells[RowNo, 22].Value = @"SMS-F-041";
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
			worksheet.Cells[RowNo++, 10].Value = @"Voyage Planning";
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
			worksheet.Cells[++RowNo, 1].Value = @"SAFETY MANAGEMENT SYSTEM";
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
			worksheet.Cells[++RowNo, 8].Value = $"Ship’s Name {HomeInfo.VesselName}";
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
			worksheet.Cells[++RowNo, 2].Value =
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
			worksheet.Cells[++RowNo, 2].Value = $"Draft  : Fwd  {HomeInfo.draftFWD} m. Centre  {HomeInfo.draftMiddle} m.   Aft  {HomeInfo.draftAFT} m.";
			using (ExcelRange range = worksheet.Cells[$"B6:Y6"])
			{
				range.Merge = true;
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				range.Style.WrapText = true;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
				range.Style.Font.Size = 10;
				range.Style.Font.Name = "Times";
			}
			RowNo++;

			AddTableHeader(worksheet);
		}



		private void AddTableHeader(ExcelWorksheet worksheet)
		{
			// Default Style

			worksheet.Row(RowNo).Height = 24;
			worksheet.Row(RowNo + 1).Height = 19.5;
			worksheet.Row(RowNo + 2).Height = 105;

			using (ExcelRange range = worksheet.Cells[$"A{RowNo}:Y{RowNo + 2}"])
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
			worksheet.Cells[RowNo, 1].Value = $"Way Points nr.";
			MergeCellsByIndex($"A{RowNo}", $"A{RowNo + 2}", worksheet);

			//B Column
			worksheet.Cells[RowNo, 2].Value = $"Position – Lat.Long.";
			MergeCellsByIndex($"B{RowNo}", $"B{RowNo + 2}", worksheet);

			//C Column
			worksheet.Cells[RowNo, 3].Value = $"Date";
			MergeCellsByIndex($"C{RowNo}", $"C{RowNo + 2}", worksheet);


			//D Column
			worksheet.Cells[RowNo, 4].Value = $"Course";
			MergeCellsByIndex($"D{RowNo}", $"D{RowNo + 2}", worksheet);

			//E Column
			worksheet.Cells[RowNo, 5].Value = $"Estimated Speed (safe speed)";
			MergeCellsByIndex($"E{RowNo}", $"E{RowNo + 2}", worksheet);

			//F Column
			worksheet.Cells[RowNo, 6].Value = $"Miles between wayponts";
			MergeCellsByIndex($"F{RowNo}", $"F{RowNo + 2}", worksheet);

			//G Column
			worksheet.Cells[RowNo, 7].Value = $"Miles to final destination";
			MergeCellsByIndex($"G{RowNo}", $"G{RowNo + 2}", worksheet);

			//H Column
			worksheet.Cells[RowNo, 8].Value = $"Chart to be used nr.";
			MergeCellsByIndex($"H{RowNo}", $"H{RowNo + 2}", worksheet);

			//I Column
			worksheet.Cells[RowNo, 9].Value = $"Ship’s reporting System (Ares, etc.)";
			MergeCellsByIndex($"I{RowNo}", $"I{RowNo + 2}", worksheet);

			//J Column
			worksheet.Cells[RowNo, 10].Value = $"VTS station and  VHF Channel";
			MergeCellsByIndex($"J{RowNo}", $"J{RowNo + 2}", worksheet);

			//K Column
			worksheet.Cells[RowNo, 11].Value = $"Marpol special areas Yes/No";
			MergeCellsByIndex($"K{RowNo}", $"K{RowNo + 2}", worksheet);

			//L Column
			worksheet.Cells[RowNo, 12].Value = $"Squat effect";
			MergeCellsByIndex($"L{RowNo}", $"L{RowNo + 2}", worksheet);

			//M Column
			worksheet.Cells[RowNo, 13].Value = $"Minimum underkeel clearance  included squat effect";
			MergeCellsByIndex($"M{RowNo}", $"M{RowNo + 2}", worksheet);

			//N Column
			worksheet.Cells[RowNo, 14].Value = $"Safe distance from obstacles";
			MergeCellsByIndex($"N{RowNo}", $"N{RowNo + 2}", worksheet);

			//O Column
			worksheet.Cells[RowNo, 15].Value = $"Ship positioning system (Gps, etc.)";
			MergeCellsByIndex($"O{RowNo}", $"O{RowNo + 2}", worksheet);

			//P Column
			worksheet.Cells[RowNo, 16].Value = $"Current Direction/Speed";
			MergeCellsByIndex($"P{RowNo}", $"P{RowNo + 2}", worksheet);

			//Q Column
			worksheet.Cells[RowNo, 17].Value = $"Stand by in the engine Yes/No";
			MergeCellsByIndex($"Q{RowNo}", $"Q{RowNo + 2}", worksheet);

			//R:V Column
			worksheet.Cells[RowNo, 18].Value = $"From pilot station to the key or  v.v.";
			MergeCellsByIndex($"R{RowNo}", $"V{RowNo}", worksheet);
			using (ExcelRange range = worksheet.Cells[$"R{RowNo}:V{RowNo}"])
			{
				range.Style.TextRotation = 0;
			}

			//R:T Column
			worksheet.Cells[RowNo + 1, 18].Value = $"Tide time";
			MergeCellsByIndex($"R{RowNo + 1}", $"T{RowNo + 1}", worksheet);
			using (ExcelRange range = worksheet.Cells[$"R{RowNo + 1}:T{RowNo + 1}"])
			{
				range.Style.TextRotation = 0;
			}

			//R3 Column
			worksheet.Cells[RowNo + 2, 18].Value = $"Low";

			//S3 Column
			worksheet.Cells[RowNo + 2, 19].Value = $"High";

			//T3 Column
			worksheet.Cells[RowNo + 2, 20].Value = $"Slack";

			//U2 Column
			worksheet.Cells[RowNo + 1, 21].Value = $"Miles pilot/berth (or  v.v.)";
			MergeCellsByIndex($"U{RowNo + 1}", $"U{RowNo + 2}", worksheet);

			//V2 Column
			worksheet.Cells[RowNo + 1, 22].Value = $"Miles to the lock";
			MergeCellsByIndex($"V{RowNo + 1}", $"V{RowNo + 2}", worksheet);

			using (ExcelRange range = worksheet.Cells[$"R{RowNo + 1}:V{RowNo + 2}"])
			{
				range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			}

			//W Column
			worksheet.Cells[RowNo, 23].Value = $"Refuge port of emergency anchorage";
			MergeCellsByIndex($"W{RowNo}", $"W{RowNo + 2}", worksheet);

			//X Column
			worksheet.Cells[RowNo, 24].Value = $"Nautical publications to be used";
			MergeCellsByIndex($"X{RowNo}", $"X{RowNo + 2}", worksheet);

			//Y Column
			worksheet.Cells[RowNo, 25].Value = $"Notes";
			MergeCellsByIndex($"Y{RowNo}", $"Y{RowNo + 2}", worksheet);

			RowNo += 3;
		}

		private void AddWayPoint(ExcelWorksheet worksheet, ImportWaipointDto wp, int wpNo)
		{
			// Default Style

			worksheet.Row(RowNo).Height = 20.25;
			worksheet.Row(RowNo + 1).Height = 20.25;

			using (ExcelRange range = worksheet.Cells[$"A{RowNo}:Y{RowNo + 1}"])
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
			worksheet.Cells[RowNo, 1].Value = $"{wpNo}";
			MergeCellsByIndex($"A{RowNo}", $"A{RowNo + 1}", worksheet);

			//B1 Column
			worksheet.Cells[RowNo, 2].Value = $"{wp.Position.LatDegrees}° {wp.Position.LatMinutes.ToString("F3")}' {wp.Position.LatDir}";
			MergeCellsByIndex($"B{RowNo}", $"B{RowNo}", worksheet);

			//B2 Column
			worksheet.Cells[RowNo + 1, 2].Value = $"{wp.Position.LongDegrees}° {wp.Position.LongMinutes.ToString("F3")}' {wp.Position.LongDir}";
			MergeCellsByIndex($"B{RowNo + 1}", $"B{RowNo + 1}", worksheet);

			//C Column
			worksheet.Cells[RowNo, 3].Value = $"##Date##";
			MergeCellsByIndex($"C{RowNo}", $"C{RowNo + 1}", worksheet);


			//D Column
			worksheet.Cells[RowNo, 4].Value = $"{wp.Course}°";
			MergeCellsByIndex($"D{RowNo}", $"D{RowNo + 1}", worksheet);

			//E Column
			worksheet.Cells[RowNo, 5].Value = $"##Estimated Speed##";
			MergeCellsByIndex($"E{RowNo}", $"E{RowNo + 1}", worksheet);

			//F Column
			worksheet.Cells[RowNo, 6].Value = $"{wp.DistanceFromLastWp}";
			MergeCellsByIndex($"F{RowNo}", $"F{RowNo + 1}", worksheet);

			//G Column
			worksheet.Cells[RowNo, 7].Value = $"{wp.DistanceToGo}";
			MergeCellsByIndex($"G{RowNo}", $"G{RowNo + 1}", worksheet);

			//H Column
			worksheet.Cells[RowNo, 8].Value = $"##Chart to be used##";
			MergeCellsByIndex($"H{RowNo}", $"H{RowNo + 1}", worksheet);

			//I Column
			worksheet.Cells[RowNo, 9].Value = $"##Ship’s reporting System##";
			MergeCellsByIndex($"I{RowNo}", $"I{RowNo + 1}", worksheet);

			//J Column
			worksheet.Cells[RowNo, 10].Value = $"##VTS station and VHF##";
			MergeCellsByIndex($"J{RowNo}", $"J{RowNo + 1}", worksheet);

			//K Column
			worksheet.Cells[RowNo, 11].Value = $"No";
			MergeCellsByIndex($"K{RowNo}", $"K{RowNo + 1}", worksheet);

			//L Column
			worksheet.Cells[RowNo, 12].Value = $"##wp1Squat##";
			MergeCellsByIndex($"L{RowNo}", $"L{RowNo + 1}", worksheet);

			//M Column
			worksheet.Cells[RowNo, 13].Value = $"##wp1UKC##";
			MergeCellsByIndex($"M{RowNo}", $"M{RowNo + 1}", worksheet);

			//N Column
			worksheet.Cells[RowNo, 14].Value = $"##wp1SafeDist##";
			MergeCellsByIndex($"N{RowNo}", $"N{RowNo + 1}", worksheet);

			//O Column
			worksheet.Cells[RowNo, 15].Value = $"GPS";
			MergeCellsByIndex($"O{RowNo}", $"O{RowNo + 1}", worksheet);

			//P Column
			worksheet.Cells[RowNo, 16].Value = $"##wp1CurrentDir##";
			MergeCellsByIndex($"P{RowNo}", $"P{RowNo + 1}", worksheet);

			//Q Column
			worksheet.Cells[RowNo, 17].Value = $"Yes";
			MergeCellsByIndex($"Q{RowNo}", $"Q{RowNo + 1}", worksheet);

			//R Column
			worksheet.Cells[RowNo, 18].Value = $"##wp1TideLow##";
			MergeCellsByIndex($"R{RowNo}", $"R{RowNo + 1}", worksheet);


			//S Column
			worksheet.Cells[RowNo, 19].Value = $"##wp1TideHigh##";
			MergeCellsByIndex($"S{RowNo}", $"S{RowNo + 1}", worksheet);

			//T Column
			worksheet.Cells[RowNo, 20].Value = $"##wp1TideSlack##";
			MergeCellsByIndex($"T{RowNo}", $"T{RowNo + 1}", worksheet);

			//U Column
			worksheet.Cells[RowNo + 1, 21].Value = $"##wp1PilotMiles##";
			MergeCellsByIndex($"U{RowNo}", $"U{RowNo + 1}", worksheet);

			//V Column
			worksheet.Cells[RowNo + 1, 22].Value = $"##wp1MilesToLock##";
			MergeCellsByIndex($"V{RowNo}", $"V{RowNo + 1}", worksheet);

			//W Column
			worksheet.Cells[RowNo, 23].Value = $"##wp1EmAnchorage##";
			MergeCellsByIndex($"W{RowNo}", $"W{RowNo + 1}", worksheet);

			//X Column
			worksheet.Cells[RowNo, 24].Value = $"##wp1NauticalPublications##";
			MergeCellsByIndex($"X{RowNo}", $"X{RowNo + 1}", worksheet);

			//Y Column
			worksheet.Cells[RowNo, 25].Value = $"##wp1Notes##";
			MergeCellsByIndex($"Y{RowNo}", $"Y{RowNo + 1}", worksheet);

			RowNo += 2;
			//RemovePageBreak(worksheet);
		}

		private void AddNameSide(ExcelWorksheet worksheet)
		{
			worksheet.Row(RowNo).Height = 13.5;
			worksheet.Row(RowNo + 1).Height = 12.75;
			worksheet.Row(RowNo + 2).Height = 13.5;
			worksheet.Row(RowNo + 3).Height = 16.5;

			using (ExcelRange range = worksheet.Cells[$"A{RowNo}:Y{RowNo + 4}"])
			{
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
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

			//A1:B1 Column
			worksheet.Cells[RowNo, 1].Value = $"date :";
			MergeCellsByIndex($"A{RowNo}", $"B{RowNo}", worksheet);

			//A2:B2 Column
			worksheet.Cells[RowNo + 1, 1].Value = $"      Slide nr.:  ";
			MergeCellsByIndex($"A{RowNo + 1}", $"B{RowNo + 1}", worksheet);

			//C1:F1 Column
			worksheet.Cells[RowNo, 3].Value = $"{HomeInfo.CreationDate.ToString("dd.MM.yyyy")}";
			MergeCellsByIndex($"C{RowNo}", $"F{RowNo}", worksheet);

			//C2:F2 Column
			worksheet.Cells[RowNo + 1, 3].Value = $" {PageNo} of : ##lastPage##";
			MergeCellsByIndex($"C{RowNo + 1}", $"F{RowNo + 1}", worksheet);

			//G1:J2 Column
			worksheet.Cells[RowNo, 7].Value = $"Delagate Officer";
			MergeCellsByIndex($"G{RowNo}", $"J{RowNo + 1}", worksheet);


			//K1:Q1 Column
			worksheet.Cells[RowNo, 11].Value = $"Name {HomeInfo.DelegateOfficerName}";
			MergeCellsByIndex($"K{RowNo}", $"Q{RowNo}", worksheet);

			//K2:Q2 Column
			worksheet.Cells[RowNo + 1, 11].Value = $"Signature";
			MergeCellsByIndex($"K{RowNo + 1}", $"Q{RowNo + 1}", worksheet);

			//R1:U2 Column
			worksheet.Cells[RowNo, 18].Value = $"Master";
			MergeCellsByIndex($"R{RowNo}", $"U{RowNo + 1}", worksheet);


			//V1:Y1 Column
			worksheet.Cells[RowNo, 22].Value = $"Name {HomeInfo.MasterName}";
			MergeCellsByIndex($"V{RowNo}", $"Y{RowNo}", worksheet);

			//V2:Y2 Column
			worksheet.Cells[RowNo + 1, 22].Value = $"Signature";
			MergeCellsByIndex($"V{RowNo + 1}", $"Y{RowNo + 1}", worksheet);

			using (ExcelRange range = worksheet.Cells[$"A{RowNo}:Y{RowNo}"])
			{
				range.Style.Border.Top.Style = ExcelBorderStyle.Double;
			}

			//Increase rows count
			RowNo += 2;

			//A3:D4 Column
			worksheet.Cells[RowNo, 1].Value = $"Duty Officer for for acknowledge";
			MergeCellsByIndex($"A{RowNo}", $"D{RowNo + 1}", worksheet);

			//E3:H3 Column
			worksheet.Cells[RowNo, 5].Value = $"Name {HomeInfo.ChiefMateName}";
			MergeCellsByIndex($"E{RowNo}", $"H{RowNo}", worksheet);

			//E4:H4 Column
			worksheet.Cells[RowNo + 1, 5].Value = $"Signature";
			MergeCellsByIndex($"E{RowNo + 1}", $"H{RowNo + 1}", worksheet);

			//I3:K4 Column
			worksheet.Cells[RowNo, 9].Value = $"Duty Officer for for acknowledge";
			MergeCellsByIndex($"I{RowNo}", $"K{RowNo + 1}", worksheet);

			//L3:Q3 Column
			worksheet.Cells[RowNo, 12].Value = $"Name {HomeInfo.WKO1Name}";
			MergeCellsByIndex($"L{RowNo}", $"Q{RowNo}", worksheet);

			//L4:Q4 Column
			worksheet.Cells[RowNo + 1, 12].Value = $"Signature";
			MergeCellsByIndex($"L{RowNo + 1}", $"Q{RowNo + 1}", worksheet);

			//R3:V4 Column
			worksheet.Cells[RowNo, 18].Value = $"Duty Officer for for acknowledge";
			MergeCellsByIndex($"R{RowNo}", $"V{RowNo + 1}", worksheet);

			//W3:Y3 Column
			worksheet.Cells[RowNo, 23].Value = $"Name {HomeInfo.WKO2Name}";
			MergeCellsByIndex($"W{RowNo}", $"Y{RowNo}", worksheet);

			//W4:Y4 Column
			worksheet.Cells[RowNo + 1, 23].Value = $"Signature";
			MergeCellsByIndex($"W{RowNo + 1}", $"Y{RowNo + 1}", worksheet);


			using (ExcelRange range = worksheet.Cells[$"A{RowNo + 1}:Y{RowNo + 1}"])
			{
				range.Style.Border.Bottom.Style = ExcelBorderStyle.Double;
			}

			RowNo += 2;
		}

		private void AddNumberPage(ExcelWorksheet worksheet)
		{
			worksheet.Row(RowNo).Height = 14.25;

			using (ExcelRange range = worksheet.Cells[$"A{RowNo}:H{RowNo}"])
			{
				range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
				range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
				range.Style.Font.Size = 8;
				range.Style.Font.Name = "Arial Narrow";
				range.Style.Border.Top.Style = ExcelBorderStyle.Double;
				range.Style.Border.Bottom.Style = ExcelBorderStyle.Double;
				range.Style.Border.Left.Style = ExcelBorderStyle.Double;
				range.Style.Border.Right.Style = ExcelBorderStyle.Double;
			}

			//A1:D1 Column
			worksheet.Cells[RowNo, 1].Value = $"     Slide nr.:  ";
			MergeCellsByIndex($"A{RowNo}", $"D{RowNo}", worksheet);

			//E1:H1 Column
			worksheet.Cells[RowNo, 5].Value = $" {PageNo} of : ##lastPage##";
			MergeCellsByIndex($"E{RowNo}", $"H{RowNo}", worksheet);

			RowNo++;
		}

		private void AddEmptyRow(ExcelWorksheet worksheet)
		{

			worksheet.Row(RowNo).Height = 44.25;

			//A1:D1 Column
			worksheet.Cells[RowNo, 1].Value = $"";
			MergeCellsByIndex($"A{RowNo}", $"Y{RowNo}", worksheet);

			using (ExcelRange range = worksheet.Cells[$"A{RowNo}:Y{RowNo}"])
			{

				range.Style.Border.Top.Style = ExcelBorderStyle.None;
				range.Style.Border.Bottom.Style = ExcelBorderStyle.None;
				range.Style.Border.Left.Style = ExcelBorderStyle.None;
				range.Style.Border.Right.Style = ExcelBorderStyle.None;
			}

			RowNo++;
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

			worksheet.PrinterSettings.LeftMargin = 0.1M;
			worksheet.PrinterSettings.RightMargin = 0.1M;
			worksheet.PrinterSettings.TopMargin = 0.1M;
			worksheet.PrinterSettings.BottomMargin = 0.5M;

			worksheet.PrinterSettings.Orientation = eOrientation.Landscape;
			worksheet.PrinterSettings.Scale = 84;
			worksheet.PrinterSettings.HorizontalCentered = true;

			for (int i = 1; i <= maxRowIndex; i++)
			{
				RemovePageBreak(worksheet, i);
			}
		}

		private void InsertPageBreak(ExcelWorksheet worksheet)
		{
			worksheet.Row(RowNo - 1).PageBreak = true;
			PageNo++;
		}

		private void RemovePageBreak(ExcelWorksheet worksheet, int row)
		{
			worksheet.Row(RowNo).PageBreak = false;
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

		private void AddLastBorder(ExcelWorksheet worksheet)
		{
			using (ExcelRange range = worksheet.Cells[$"A{RowNo}:Y{RowNo}"])
			{
				range.Style.Border.Top.Style = ExcelBorderStyle.Double;
			}
		}

		private void ReplaceText(ExcelWorksheet worksheet, string startCell, string endCell, string searchText, string replaceText)
		{
			ExcelRangeBase range = worksheet.Cells[$"{startCell}:{endCell}"];

			foreach (ExcelRangeBase cell in range)
			{
				if (cell.Text.Contains(searchText))
				{
					cell.Value = cell.Text.Replace(searchText, replaceText);
				}
			}
		}
	}
}
