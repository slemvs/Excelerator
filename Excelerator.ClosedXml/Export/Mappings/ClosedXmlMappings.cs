using ClosedXML.Excel;
using Excelerator.Enums;

namespace Excelerator.ClosedXml.Export.Mappings
{
	public static class ClosedXmlMappings
	{
		public static XLAlignmentHorizontalValues Map(this HorizontalCellAlignmentValues alignment)
		{
			switch (alignment)
			{
				case HorizontalCellAlignmentValues.General: return XLAlignmentHorizontalValues.General;
				case HorizontalCellAlignmentValues.Center: return XLAlignmentHorizontalValues.Center;
				case HorizontalCellAlignmentValues.CenterContinuous: return XLAlignmentHorizontalValues.CenterContinuous;
				case HorizontalCellAlignmentValues.Distributed: return XLAlignmentHorizontalValues.Distributed;
				case HorizontalCellAlignmentValues.Fill: return XLAlignmentHorizontalValues.Fill;
				case HorizontalCellAlignmentValues.Justify: return XLAlignmentHorizontalValues.Justify;
				case HorizontalCellAlignmentValues.Left: return XLAlignmentHorizontalValues.Left;
				case HorizontalCellAlignmentValues.Right: return XLAlignmentHorizontalValues.Right;
				default: return XLAlignmentHorizontalValues.General;
			}
		}
		public static XLAlignmentVerticalValues Map(this VerticalCellAlignmentValues alignment)
		{
			switch (alignment)
			{
				case VerticalCellAlignmentValues.Center: return XLAlignmentVerticalValues.Center;
				case VerticalCellAlignmentValues.Top: return XLAlignmentVerticalValues.Top;
				case VerticalCellAlignmentValues.Bottom: return XLAlignmentVerticalValues.Bottom;
				case VerticalCellAlignmentValues.Distributed: return XLAlignmentVerticalValues.Distributed;
				case VerticalCellAlignmentValues.Justify: return XLAlignmentVerticalValues.Justify;
				default:
					return XLAlignmentVerticalValues.Top;
			}
		}
	}
}
