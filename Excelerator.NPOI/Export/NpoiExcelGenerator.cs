using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excelerator.Common.Export.Metadata;
using Excelerator.Export;
using Excelerator.NPOI.Export.Mappings;
using Excelerator.NPOI.Extensions;
using NPOI.HSSF.UserModel;

namespace Excelerator.NPOI.Export
{
	public class NpoiExcelGenerator<T> : IExcelGenerator<T>
		where T : class
	{
		private readonly int _startCol = 1;
		private readonly int _startRow = 1;

		public MemoryStream Generate(WorksheetMetadata<T> wsMetadata, IEnumerable<T> data)
		{
			var dataArray = data as T[] ?? data.ToArray();
			if (wsMetadata?.ColumnsMetadata == null || !wsMetadata.ColumnsMetadata.Any() || !dataArray.Any())
				return null;
			var ms = new MemoryStream();
			var wb = new HSSFWorkbook();
			var sh = (HSSFSheet) wb.CreateSheet();
			//var headerCellStyle = (HSSFCellStyle)wb.CreateCellStyle();
			//var dataCellStyle = (HSSFCellStyle)wb.CreateCellStyle();
			//var font = (HSSFFont)wb.CreateFont();


			//worksheet.Cell(i, j).Style.Alignment.Horizontal = md.HorizontalAlignment.Map();
			//worksheet.Cell(i, j).Style.Alignment.Vertical = md.VerticalAlignment.Map();


			var row = wsMetadata.StartRow == 0 ? _startRow : wsMetadata.StartRow;
			var startCol = wsMetadata.StartColumn == 0 ? _startCol - 1 : wsMetadata.StartColumn - 1;

			// Set headers
			var headerRow = (HSSFRow) sh.CreateRow(row);
			for (var i = 0; i < wsMetadata.ColumnsMetadata.Count; i++)
			{
				var cmd = wsMetadata.ColumnsMetadata[i];
				var value = cmd.Header;
				var address = cmd.ColumnAddress;
				var cellCol = address?.Column.GetColumnNumberFromLetter() ?? startCol + i;
				var cell = headerRow.CreateCell(cellCol);
				cell.SetCellValue(value);
				cell.CellStyle.Alignment = cmd.HorizontalAlignment.Map();
				cell.CellStyle.VerticalAlignment = cmd.VerticalAlignment.Map();
			}
			row++;

			foreach (var val in dataArray)
			{
				var r = (HSSFRow) sh.CreateRow(row);
				for (var i = 0; i < wsMetadata.ColumnsMetadata.Count; i++)
				{
					var cmd = wsMetadata.ColumnsMetadata[i];
					var value = cmd.Value(val);
					var address = cmd.ColumnAddress;
					var cellCol = address?.Column.GetColumnNumberFromLetter() ?? startCol + i;
					var cell = r.CreateCell(cellCol);
					cell.SetCellValue(value);
					cell.CellStyle.Alignment = cmd.HorizontalAlignment.Map();
					cell.CellStyle.VerticalAlignment = cmd.VerticalAlignment.Map();
				}
				row++;
			}
			wb.Write(ms);
			return ms;
		}

		public MemoryStream Generate(MemoryStream templateStream, WorksheetMetadata<T> wsMetadata, IEnumerable<T> data)
		{
			var dataArray = data as T[] ?? data.ToArray();
			if (wsMetadata?.ColumnsMetadata == null || !wsMetadata.ColumnsMetadata.Any() || !dataArray.Any())
				return null;
			var ms = new MemoryStream();
			var wb = new HSSFWorkbook(templateStream);
			var sh = (HSSFSheet)wb.GetSheetAt(0);
			//var headerCellStyle = (HSSFCellStyle)wb.CreateCellStyle();
			//var dataCellStyle = (HSSFCellStyle)wb.CreateCellStyle();
			//var font = (HSSFFont)wb.CreateFont();
			
			//worksheet.Cell(i, j).Style.Alignment.Horizontal = md.HorizontalAlignment.Map();
			//worksheet.Cell(i, j).Style.Alignment.Vertical = md.VerticalAlignment.Map();
			
			var row = wsMetadata.StartRow == 0 ? _startRow : wsMetadata.StartRow;
			var startCol = wsMetadata.StartColumn == 0 ? _startCol - 1 : wsMetadata.StartColumn - 1;

			// Set headers
			var headerRow = (HSSFRow)sh.CreateRow(row);
			for (var i = 0; i < wsMetadata.ColumnsMetadata.Count; i++)
			{
				var cmd = wsMetadata.ColumnsMetadata[i];
				var value = cmd.Header;
				var address = cmd.ColumnAddress;
				var cellCol = address?.Column.GetColumnNumberFromLetter() ?? startCol + i;
				var cell = headerRow.CreateCell(cellCol);
				cell.SetCellValue(value);
				cell.CellStyle.Alignment = cmd.HorizontalAlignment.Map();
				cell.CellStyle.VerticalAlignment = cmd.VerticalAlignment.Map();
			}
			row++;

			foreach (var val in dataArray)
			{
				var r = (HSSFRow)sh.CreateRow(row);
				for (var i = 0; i < wsMetadata.ColumnsMetadata.Count; i++)
				{
					var cmd = wsMetadata.ColumnsMetadata[i];
					var value = cmd.Value(val);
					var address = cmd.ColumnAddress;
					var cellCol = address?.Column.GetColumnNumberFromLetter() ?? startCol + i;
					var cell = r.CreateCell(cellCol);
					cell.SetCellValue(value);
					cell.CellStyle.Alignment = cmd.HorizontalAlignment.Map();
					cell.CellStyle.VerticalAlignment = cmd.VerticalAlignment.Map();
				}
				row++;
			}
			wb.Write(ms);
			return ms;
		}
	}
}