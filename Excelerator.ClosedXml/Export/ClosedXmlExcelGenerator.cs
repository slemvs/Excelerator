using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Excelerator.ClosedXml.Export.Mappings;
using Excelerator.Common.Export.Metadata;
using Excelerator.Common.Import;
using Excelerator.Export;

namespace Excelerator.ClosedXml.Export
{
	public class ClosedXmlExcelGenerator<T> : IExcelGenerator<T>
		where T : class
	{
		private readonly int _startCol = 1;
		private readonly int _startRow = 1;

		protected XLWorkbook Workbook { get; set; }

		public MemoryStream Generate(WorksheetMetadata<T> wsMetadata, IEnumerable<T> data)
		{
			var ms = new MemoryStream();
			var workbook = Workbook ?? new XLWorkbook(XLEventTracking.Disabled);
			var worksheet = workbook.Worksheets.Any() ? workbook.Worksheet(1) : workbook.Worksheets.Add(string.IsNullOrEmpty(wsMetadata.Name) ? "Sheet1" : wsMetadata.Name);
			var row = wsMetadata.StartRow == 0 ? _startRow : wsMetadata.StartRow;
			var rangeStartRow = row;
			var rangeStartColumn = wsMetadata.StartColumn == 0 ? _startCol : wsMetadata.StartColumn;

			// Set headers
			Process((i, j, md) =>
			{
				worksheet.Column(j).Width = md.Width;
				worksheet.Cell(i, j).Value = md.Header;
				worksheet.Cell(i, j).Style.Alignment.Horizontal = md.HorizontalAlignment.Map();
				worksheet.Cell(i, j).Style.Alignment.Vertical = md.VerticalAlignment.Map();
			}, row, wsMetadata);
			row++;
			if (data != null)
				foreach (var val in data)
				{
					Process((i, j, md) =>
					{
						var value = md.Value(val);
						var cell = worksheet.Cell(i, j);
						cell.Value = value ?? string.Empty;
						cell.Style.Alignment.Horizontal = md.HorizontalAlignment.Map();
						cell.Style.Alignment.Vertical = md.VerticalAlignment.Map();
					}, row, wsMetadata);
					row++;
				}

			if (wsMetadata.FormatAsTable)
				worksheet.Range(rangeStartRow, rangeStartColumn, row - 1, rangeStartColumn + wsMetadata.ColumnsMetadata.Count - 1).CreateTable();

			workbook.SaveAs(ms);
			ms.Position = 0;
			return ms;
		}

		private void Process(Action<int, int, ColumnMetadata<T>> action, int currentRow, WorksheetMetadata<T> md)
		{
			var startCol = md.StartColumn == 0 ? _startCol : md.StartColumn;
			for (var i = 1; i <= md.ColumnsMetadata.Count; i++)
				action(currentRow, startCol + i - 1, md.ColumnsMetadata.ElementAt(i - 1));
		}
	}
}