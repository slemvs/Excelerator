using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excelerator.Export;

namespace Excelerator.OpenXml.Export
{
	public class OpenXmlExcelGenerator<TModel> : IExcelGenerator<TModel>
		where TModel : class
	{
		public string WorksheetName { get; set; }

		public MemoryStream Generate(List<ExcelColumnMetadata<TModel>> columns, IEnumerable<ExcelRowModel<TModel>> data)
		{
			var colsCount = columns.Count;
			//var rowsCount = data.Count;

			var ms = new MemoryStream();
			using (SpreadsheetDocument myDoc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
			{
				WorkbookPart workbookPart = myDoc.AddWorkbookPart();
				//WorksheetPart worksheetPart = workbookPart.AddNewPart<>() WorksheetParts.First();
				//string origninalSheetId = workbookPart.GetIdOfPart(worksheetPart);

				WorksheetPart replacementPart = workbookPart.AddNewPart<WorksheetPart>();
				string replacementPartId = workbookPart.GetIdOfPart(replacementPart);

				//OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
				OpenXmlWriter writer = OpenXmlWriter.Create(replacementPart);

				Row r = new Row();
				Cell c = new Cell();
				CellFormula f = new CellFormula();
				f.CalculateCell = true;
				f.Text = "RAND()";
				c.Append(f);
				CellValue v = new CellValue();
				c.Append(v);


				writer.WriteStartElement(new SheetData());
				foreach (var el in data)
				{
					writer.WriteStartElement(r);
					for (int col = 0; col < colsCount; col++)
					{
						writer.WriteElement(c);
					}
					writer.WriteEndElement();
					
				}
				writer.WriteEndElement();

				//while (reader.Read())
				//{
				//	if (reader.ElementType == typeof(SheetData))
				//	{
				//		if (reader.IsEndElement)
				//			continue;
				//		writer.WriteStartElement(new SheetData());

				//		for (int row = 0; row < rowsCount; row++)
				//		{
				//			writer.WriteStartElement(r);
				//			for (int col = 0; col < colsCount; col++)
				//			{
				//				writer.WriteElement(c);
				//			}
				//			writer.WriteEndElement();
				//		}

				//		writer.WriteEndElement();
				//	}
				//	else
				//	{
				//		if (reader.IsStartElement)
				//		{
				//			writer.WriteStartElement(reader);
				//		}
				//		else if (reader.IsEndElement)
				//		{
				//			writer.WriteEndElement();
				//		}
				//	}
				//}
				//reader.Close();
				writer.Close();

				//var sheet = workbookPart.Workbook.Descendants<Sheet>().First(s => s.Id.Value.Equals(origninalSheetId));
				//sheet.Id.Value = replacementPartId;
				//workbookPart.DeletePart(worksheetPart);

			}
			ms.Position = 0;
			return ms;
		}
	}
}
