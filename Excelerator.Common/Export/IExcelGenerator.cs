using System;
using System.Collections.Generic;
using System.IO;

namespace Excelerator.Export
{
	public interface IExcelGenerator<TModel>
		where TModel : class
	{
		string WorksheetName { get; set; }
		MemoryStream Generate(List<ExcelColumnMetadata<TModel>> columnsMetadata, IEnumerable<ExcelRowModel<TModel>> data);
	}
}
