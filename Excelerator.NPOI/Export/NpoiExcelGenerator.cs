using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excelerator.Common.Export.Metadata;
using Excelerator.Export;
using Excelerator.NPOI.Export.Mappings;
using Excelerator.NPOI.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

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
			var sh = (HSSFSheet)wb.CreateSheet();
			//var headerCellStyle = (HSSFCellStyle)wb.CreateCellStyle();
			//var dataCellStyle = (HSSFCellStyle)wb.CreateCellStyle();
			//var font = (HSSFFont)wb.CreateFont();


			//worksheet.Cell(i, j).Style.Alignment.Horizontal = md.HorizontalAlignment.Map();
			//worksheet.Cell(i, j).Style.Alignment.Vertical = md.VerticalAlignment.Map();


			var row = wsMetadata.StartRow == 0 ? _startRow : wsMetadata.StartRow;
			var startCol = wsMetadata.StartColumn == 0 ? _startCol - 1 : wsMetadata.StartColumn - 1;

			// Set headers
			if (wsMetadata.GenerateHeader)
			{
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
			}
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

			var row = wsMetadata.StartRow == 0 ? _startRow - 1 : wsMetadata.StartRow - 1;
			var startCol = wsMetadata.StartColumn == 0 ? _startCol - 1 : wsMetadata.StartColumn - 1;

			// Set headers
			if (wsMetadata.GenerateHeader)
			{
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
			}

			var sourceRow = (HSSFRow)sh.GetRow(row);
			var rowsCount = dataArray.Count();
			sh.ShiftRows(sourceRow.RowNum + 1, sourceRow.RowNum + 1 + rowsCount, rowsCount, true, false);

			foreach (var val in dataArray)
			{
				var newRow = CopyRow(sourceRow, sh, row);
				for (var i = 0; i < wsMetadata.ColumnsMetadata.Count; i++)
				{
					var cmd = wsMetadata.ColumnsMetadata[i];
					var value = cmd.Value(val);
					var address = cmd.ColumnAddress;
					var cellCol = address?.Column.GetColumnNumberFromLetter() ?? startCol + i;
					var cell = newRow.GetCell(cellCol);
					cell.SetCellValue(value);
					cell.CellStyle.Alignment = cmd.HorizontalAlignment.Map();
					cell.CellStyle.VerticalAlignment = cmd.VerticalAlignment.Map();
				}
				row++;
			}
			wb.Write(ms);
			return ms;
		}

		private HSSFRow CopyRow(HSSFRow sourceRow, HSSFSheet sheet, int rowNum)
		{
			var newRowNum = rowNum;
			var res = (HSSFRow)sheet.CreateRow(rowNum);
			sourceRow.Cells.ForEach(sourceCell =>
			{
				var newCell = res.CreateCell(sourceCell.ColumnIndex);
				newCell.CellStyle = sourceCell.CellStyle;

			});
			var rowMergedRegions = GetMergeRegions(sourceRow.RowNum, sheet);

			foreach (var sourceMergedRegion in rowMergedRegions)
			{
				var cra = new CellRangeAddress(newRowNum, newRowNum, sourceMergedRegion.FirstColumn, sourceMergedRegion.LastColumn);
				sheet.AddMergedRegion(cra);
			}

			return res;
		}

		private List<CellRangeAddress> GetMergeRegions(int row, HSSFSheet sheet)
		{
			var res = new List<CellRangeAddress>();
			for (int i = 0; i < sheet.NumMergedRegions; i++)
			{
				var mr = sheet.GetMergedRegion(i);
				if (mr.FirstRow == row && mr.LastRow == row && !res.Any(r => r.FirstRow == mr.FirstRow && r.LastRow == mr.LastRow && r.FirstColumn == mr.FirstColumn && r.LastColumn == mr.LastColumn))
				{
					res.Add(mr);
				}
			}
			return res;
		}
	}
}