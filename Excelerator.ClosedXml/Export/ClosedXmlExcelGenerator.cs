﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Excelerator.ClosedXml.Export.Mappings;
using Excelerator.Common.Export.Metadata;
using Excelerator.Export;

namespace Excelerator.ClosedXml.Export
{
	public class ClosedXmlExcelGenerator<T> : IExcelGenerator<T>
		where T : class
	{
		private readonly int _startCol = 1;
		private readonly int _startRow = 1;

		protected XLWorkbook Workbook { get; set; }

		protected int StartRow { get; set; }
		protected int StartColumn { get; set; }
		public string WorksheetName { get; set; }

		public MemoryStream Generate(WorksheetMetadata<T> wsMetadata, IEnumerable<T> data)
		{
			var ms = new MemoryStream();
			var workbook = Workbook ?? new XLWorkbook(XLEventTracking.Disabled);
			var worksheet = workbook.Worksheets.Any() ? workbook.Worksheet(1) : workbook.Worksheets.Add(string.IsNullOrEmpty(WorksheetName) ? "Sheet1" : WorksheetName);
			var row = StartRow == 0 ? _startRow : StartRow;

			// Set headers
			Process((i, j, md) =>
			{
				worksheet.Column(j).Width = md.Width;
				worksheet.Cell(i, j).Value = md.Header;
				worksheet.Cell(i, j).Style.Alignment.Horizontal = md.HorizontalAlignment.Map();
				worksheet.Cell(i, j).Style.Alignment.Vertical = md.VerticalAlignment.Map();
			}, row, wsMetadata.ColumnsMetadata);
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
					}, row, wsMetadata.ColumnsMetadata);
					row++;
				}
			workbook.SaveAs(ms);
			ms.Position = 0;
			return ms;
		}

		private void Process(Action<int, int, ColumnMetadata<T>> action, int currentRow, List<ColumnMetadata<T>> columns)
		{
			var startCol = StartColumn == 0 ? _startCol : StartColumn;
			for (var i = _startCol; i <= columns.Count; i++)
				action(currentRow, i, columns.ElementAt(i - 1));
		}
	}
}