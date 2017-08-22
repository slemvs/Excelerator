using System;
using Excelerator.Common.Models;
using Excelerator.Enums;

namespace Excelerator.Common.Export.Metadata
{
	public abstract class ColumnMetadata
	{
		public string Header { get; set; }
		public double Width { get; set; } = 20;
		public HorizontalCellAlignmentValues HorizontalAlignment { get; set; } = HorizontalCellAlignmentValues.General;
		public VerticalCellAlignmentValues VerticalAlignment { get; set; } = VerticalCellAlignmentValues.Center;
		public ColumnAddress ColumnAddress { get; set; }
	}

	public class ColumnMetadata<T> : ColumnMetadata
		where T : class
	{
		public Func<T, string> Value;
	}
}
