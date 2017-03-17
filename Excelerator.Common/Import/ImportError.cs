using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelerator.Common.Import
{
	public class ImportError<T>
	{
		public T Value { get; set; }
		public ValidationResult ValidationResult { get; set; }
	}
}
