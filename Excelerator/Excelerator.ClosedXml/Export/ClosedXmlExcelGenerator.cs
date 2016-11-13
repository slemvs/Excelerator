using ClosedXML.Excel;
using Excelerator.ClosedXml.Export.Mappings;
using Excelerator.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Excelerator.ClosedXml.Export
{
	public class ClosedXmlExcelGenerator<TModel> : IExcelGenerator<TModel>
		where TModel : class
	{
		private int _startCol = 1;
		private int _startRow = 1;

		protected XLWorkbook Workbook { get; set; }
		public string WorksheetName { get; set; }

		protected int StartRow { get; set; }
		protected int StartColumn { get; set; }

		protected virtual XLWorkbook GenerateWorkbook(List<ExcelColumnMetadata<TModel>> columnsMd, List<ExcelRowModel<TModel>> data)
		{
			var workbook = Workbook ?? new XLWorkbook();
			var worksheet = workbook.Worksheets.Any() ? workbook.Worksheet(1) : workbook.Worksheets.Add(string.IsNullOrEmpty(WorksheetName) ? "Sheet1" : WorksheetName);

			var row = StartRow == 0 ? _startRow : StartRow;

			// Set headers
			Process((i, j, md) =>
			{
				worksheet.Column(j).Width = md.Width;
				worksheet.Cell(i, j).Value = md.Header;
				worksheet.Cell(i, j).Style.Alignment.Horizontal = md.HorizontalAlignment.Map();
				worksheet.Cell(i, j).Style.Alignment.Vertical = md.VerticalAlignment.Map();
			}, row, columnsMd);
			row++;
			if (data != null)
			{
				foreach (var val in data)
				{
					Process((i, j, md) =>
					{
						var value = md.Value(val.Data);
						worksheet.Cell(i, j).Value = value ?? string.Empty;
						worksheet.Cell(i, j).Style.Alignment.Horizontal = md.HorizontalAlignment.Map();
						worksheet.Cell(i, j).Style.Alignment.Vertical = md.VerticalAlignment.Map();
					}, row, columnsMd);
					row++;
				}
			}
			return workbook;
		}

		public virtual MemoryStream Generate(List<ExcelColumnMetadata<TModel>> columnsMd, List<ExcelRowModel<TModel>> data)
		{
			var ms = new MemoryStream();
			var workbook = GenerateWorkbook(columnsMd, data);
			workbook.SaveAs(ms);
			ms.Position = 0;
			return ms;
		}


		private void Process(Action<int, int, ExcelColumnMetadata<TModel>> action, int currentRow, List<ExcelColumnMetadata<TModel>> columns)
		{
			var startCol = StartColumn == 0 ? _startCol : StartColumn;
			for (int i = _startCol; i <= columns.Count; i++)
				action(currentRow, i, columns.ElementAt(i - 1));
		}


	}
}
