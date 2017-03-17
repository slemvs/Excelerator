using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelerator.Common.Import
{
	public class ExcelColumnImportMetadata
	{
		public ExcelColumnImportMetadata() { }
		public ExcelColumnImportMetadata(string title)
		{
			Title = title;
		}
		public string Title { get; set; }
		public Func<dynamic, dynamic> CellValue { get; set; }
		public Action<dynamic, dynamic> SetModelPropertyValue { get; set; }

	}

	public class ExcelColumnImportMetadata<T, TColumnType> : ExcelColumnImportMetadata
	{
		public ExcelColumnImportMetadata() : base() { }
		public ExcelColumnImportMetadata(string title) : base(title) { }
		public new Func<object, TColumnType> CellValue { get; set; }
		public new Action<T, TColumnType> SetModelPropertyValue { get; set; }
	}
}
