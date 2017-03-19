using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using Excelerator.Common.Import;
using Excelerator.Examples.Import.Example1.Model;

namespace Excelerator.Examples.Import.Example1
{
	public class ImportExample1 : ImportExampleBase
	{
		private readonly IExcelImporter<Example1RowModel> _importer;

		public ImportExample1(IExcelImporter<Example1RowModel> importer)
		{
			_importer = importer;
		}

		public void Execute()
		{
			var md = new List<ExcelColumnImportMetadata>
			{
				new ExcelColumnImportMetadata<Example1RowModel, string>("Value1") {CellValue = GetString, SetModelPropertyValue = (model, value) => model.Value1 = value},
				new ExcelColumnImportMetadata<Example1RowModel, long>("Value2") {CellValue = obj => Convert.ToInt64(obj), SetModelPropertyValue = (model, value) => model.Value2 = value}
			};
			var wb = new XLWorkbook("Import\\Example1\\Data\\example1.xlsx");
			var ms = new MemoryStream();
			wb.SaveAs(ms);
			ms.Position = 0;
			_importer.Import(md, ms);
		}

		private string GetString(object obj)
		{
			var res = Convert.ToString(obj);
			return res.Trim();
		}
	}
}