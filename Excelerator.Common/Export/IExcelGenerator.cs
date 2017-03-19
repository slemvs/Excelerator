using System;
using System.Collections.Generic;
using System.IO;
using Excelerator.Common.Export.Metadata;

namespace Excelerator.Export
{
	public interface IExcelGenerator<T>
		where T : class
	{
		string WorksheetName { get; set; }
		MemoryStream Generate(WorksheetMetadata<T> wsMetadata, IEnumerable<T> data);
	}
}
