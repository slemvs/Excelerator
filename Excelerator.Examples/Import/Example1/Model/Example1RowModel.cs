using Excelerator.Common.Import;

namespace Excelerator.Examples.Import.Example1.Model
{
	public class Example1RowModel : IRowModelBase
	{
		public string Value1 { get; set; }
		public long Value2 { get; set; }
		public int Row { get; set; }
	}
}