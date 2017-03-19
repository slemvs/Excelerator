using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excelerator.Common.Export.Metadata;
using Excelerator.Export;

namespace Excelerator.OpenXml.Export
{
	public class OpenXmlExcelGenerator<T> : IExcelGenerator<T>
		where T : class
	{
		public MemoryStream Generate(WorksheetMetadata<T> wsMetadata, IEnumerable<T> data)
		{
			var colsCount = wsMetadata.ColumnsMetadata.Count;
			//var rowsCount = data.Count;

			var ms = new MemoryStream();
			using (var myDoc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
			{
				var workbookPart = myDoc.AddWorkbookPart();
				//WorksheetPart worksheetPart = workbookPart.AddNewPart<>() WorksheetParts.First();
				//string origninalSheetId = workbookPart.GetIdOfPart(worksheetPart);

				var replacementPart = workbookPart.AddNewPart<WorksheetPart>();
				var replacementPartId = workbookPart.GetIdOfPart(replacementPart);

				//OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
				var writer = OpenXmlWriter.Create(replacementPart);

				var r = new Row();
				var c = new Cell();
				var f = new CellFormula();
				f.CalculateCell = true;
				f.Text = "RAND()";
				c.Append(f);
				var v = new CellValue();
				c.Append(v);


				writer.WriteStartElement(new SheetData());
				foreach (var el in data)
				{
					writer.WriteStartElement(r);
					for (var col = 0; col < colsCount; col++)
						writer.WriteElement(c);
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