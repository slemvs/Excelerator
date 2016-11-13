using System.Collections.Generic;
using System.IO;

namespace Excelerator.Export
{
	public interface IExcelGenerator<TModel>
		where TModel : class
	{
		string WorksheetName { get; set; }
		MemoryStream Generate(List<ExcelColumnMetadata<TModel>> columns, List<ExcelRowModel<TModel>> data);
	}
}
