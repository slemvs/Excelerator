using Excelerator.Enums;
using NPOI.SS.UserModel;

namespace Excelerator.NPOI.Export.Mappings
{
	public static class NpoiMappings
	{
		public static HorizontalAlignment Map(this HorizontalCellAlignmentValues alignment)
		{
			switch (alignment)
			{
				case HorizontalCellAlignmentValues.General: return HorizontalAlignment.General;
				case HorizontalCellAlignmentValues.Center: return HorizontalAlignment.Center;
				case HorizontalCellAlignmentValues.CenterContinuous: return HorizontalAlignment.CenterSelection;
				case HorizontalCellAlignmentValues.Distributed: return HorizontalAlignment.Distributed;
				case HorizontalCellAlignmentValues.Fill: return HorizontalAlignment.Fill;
				case HorizontalCellAlignmentValues.Justify: return HorizontalAlignment.Justify;
				case HorizontalCellAlignmentValues.Left: return HorizontalAlignment.Left;
				case HorizontalCellAlignmentValues.Right: return HorizontalAlignment.Right;
				default: return HorizontalAlignment.General;
			}
		}

		public static VerticalAlignment Map(this VerticalCellAlignmentValues alignment)
		{
			switch (alignment)
			{
				case VerticalCellAlignmentValues.Center: return VerticalAlignment.Center;
				case VerticalCellAlignmentValues.Top: return VerticalAlignment.Top;
				case VerticalCellAlignmentValues.Bottom: return VerticalAlignment.Bottom;
				case VerticalCellAlignmentValues.Distributed: return VerticalAlignment.Distributed;
				case VerticalCellAlignmentValues.Justify: return VerticalAlignment.Justify;
				default:
					return VerticalAlignment.Top;
			}
		}
	}
}