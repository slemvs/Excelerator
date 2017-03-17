using System;

namespace Excelerator.Common.Import.Exceptions
{
	public class ExcelImporterException : ApplicationException
	{
		public ExcelImporterException(string message) : base(message)
		{
		}
	}
}