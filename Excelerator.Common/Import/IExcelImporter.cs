using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelerator.Common.Import
{
	public interface IExcelImporter<T>
		where T : class, IRowModelBase, new()
	{
		ImportResults<T> Import(List<ExcelColumnImportMetadata> md, MemoryStream ms);
	}
	public class ImportResults<T> : ImportResults
		where T : IRowModelBase
	{
		public ImportResults()
		{
			Values = new List<T>();
			Errors = new List<ImportError<T>>();
		}
		public List<T> Values { get; set; }
		public List<ImportError<T>> Errors { get; set; }
	}
}
