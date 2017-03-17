using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excelerator.Common.Import;
using Excelerator.Examples.Import.Example1.Model;

namespace Excelerator.Examples.Import.Example1
{
	public class Example1 : ImportExampleBase
	{
		private readonly IExcelImporter<Example1RowModel> _importer;
		public Example1(IExcelImporter<Example1RowModel> importer)
		{
			_importer = importer;
		}

		public void Execute()
		{
			var md = new List<ExcelColumnImportMetadata>()
			{
				new ExcelColumnImportMetadata<Example1RowModel, string>("Фамилия") { CellValue = GetString, SetModelPropertyValue = (model, value) => model.Value1 = value },
				new ExcelColumnImportMetadata<Example1RowModel, long>("Имя") { CellValue = (obj)=> Convert.ToInt64(obj), SetModelPropertyValue = (model, value) => model.Value2 = value },
			};


			
			var ms = new MemoryStream();
			_importer.Import(md, ms);

		}
		private string GetString(object obj)
		{
			var res = Convert.ToString(obj);
			return res.Trim();
		}
	}
}
