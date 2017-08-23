using System.Collections.Generic;

namespace Excelerator.Common.Export.Metadata
{
	public class WorksheetMetadata<T> where T : class
	{
		public WorksheetMetadata()
		{
			ColumnsMetadata = new List<ColumnMetadata<T>>();
		}

		public string Name { get; set; }
		public int StartRow { get; set; }
		public int StartColumn { get; set; }
		public bool FormatAsTable { get; set; } = true;
		public bool GenerateHeader { get; set; }
		public List<ColumnMetadata<T>> ColumnsMetadata { get; set; }
	}
}
