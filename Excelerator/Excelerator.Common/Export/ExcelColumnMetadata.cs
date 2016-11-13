using Excelerator.Enums;
using System;

namespace Excelerator.Export
{
	public abstract class ExcelColumnMetadata
	{
		public string Header { get; set; }
		public double Width { get; set; } = 20;
		public HorizontalCellAlignmentValues HorizontalAlignment { get; set; } = HorizontalCellAlignmentValues.General;
		public VerticalCellAlignmentValues VerticalAlignment { get; set; } = VerticalCellAlignmentValues.Center;
	}

	public class ExcelColumnMetadata<TModel> : ExcelColumnMetadata
		where TModel : class
	{
		public Func<TModel, string> Value;
	}
}
