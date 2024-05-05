using Aspose.Words.Tables;
using Aspose.Words;
using System.IO;
using System.Windows.Forms;
using System;
using ECDIS_eGloebe___RouteConverter.DTOs;

namespace ECDIS_eGloebe___RouteConverter.Common
{
	public static class Common
	{
		public static HomeInfoDto HomeInfo = new HomeInfoDto();

		public static ImportRouteDto RouteDto = new ImportRouteDto();

		public static int RowNo = 0;

		public static int PageNo = 1;


		//public static void ExportToWordDoc()
		//{
		//	Document doc = new Document();
		//	DocumentBuilder builder = new DocumentBuilder();

		//	// Specify font formatting
		//	Font font = builder.Font;
		//	font.Size = 32;
		//	font.Bold = true;
		//	font.Color = System.Drawing.Color.Black;
		//	font.Name = "Arial";
		//	font.Underline = Underline.Single;

		//	builder.Document = doc;

		//	// Insert text
		//	builder.Writeln("This is the first page.");
		//	builder.Writeln();

		//	// Change formatting for next elements.
		//	font.Underline = Underline.None;
		//	font.Size = 10;
		//	font.Color = System.Drawing.Color.Blue;

		//	builder.Writeln("This following is a table");
		//	// Insert a table
		//	Table table = builder.StartTable();
		//	// Insert a cell
		//	builder.InsertCell();
		//	// Use fixed column widths.
		//	table.AutoFit(AutoFitBehavior.AutoFitToContents);
		//	builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
		//	builder.Write("This is row 1 cell 1");
		//	// Insert a cell
		//	builder.InsertCell();
		//	builder.Write("This is row 1 cell 2");
		//	builder.EndRow();
		//	builder.InsertCell();
		//	builder.Write("This is row 2 cell 1");
		//	builder.InsertCell();
		//	builder.Write("This is row 2 cell 2");
		//	builder.EndRow();
		//	builder.EndTable();
		//	builder.Writeln();

		//	// Insert image
		//	//builder.InsertImage("image.png");

		//	// Insert page break 
		//	builder.InsertBreak(BreakType.PageBreak);
		//	// all the elements after page break will be inserted to next page.

		//	// Save the document


		//	// Създаване на SaveFileDialog
		//	SaveFileDialog saveFileDialog1 = new SaveFileDialog();

		//	// Задаване на филтър за разширения на файловете
		//	saveFileDialog1.Filter = "Word document (*.docx)|*.docx|Всички файлове (*.*)|*.*";
		//	saveFileDialog1.FilterIndex = 1;
		//	saveFileDialog1.RestoreDirectory = true;

		//	// Ако потребителят избере място и натисне "Запис", ще продължим
		//	if (saveFileDialog1.ShowDialog() == DialogResult.OK)
		//	{
		//		try
		//		{
		//			doc.Save(saveFileDialog1.FileName);
		//			MessageBox.Show("Your file is succesfully saved");
		//		}
		//		catch (Exception ex)
		//		{
		//			MessageBox.Show("Eror while saving" + ex.Message);
		//		}
		//	}

		//}
	}
}
