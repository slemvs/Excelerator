using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelerator.Common.Models
{
	public class ColumnAddress
	{
		public ColumnAddress(string column)
		{
			Column = column;
		}
		public string Column { get; set; }
	}
}
