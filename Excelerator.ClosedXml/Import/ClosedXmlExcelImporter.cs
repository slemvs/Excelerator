using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Excelerator.Common.Import;
using Excelerator.Common.Import.Exceptions;

namespace Excelerator.ClosedXml.Import
{
	public class ClosedXmlExcelImporter<T> : IExcelImporter<T> where T : class, IRowModelBase, new()
	{
		public ImportResults<T> Import(List<ExcelColumnImportMetadata> md, MemoryStream ms)
		{
			var res = new ImportResults<T>();
			var wb = new XLWorkbook(ms);
			if (!wb.Worksheets.Any())
				throw new ExcelImporterException("Workbook doesn't contain any worksheet");

			var sh = wb.Worksheet(1);
			var headerCell = SearchHeader(sh, md);
			var colsOrder = GetColumnsOrder(sh, md);

			ProcessRows(row =>
			{
				var item = new T {Row = row};
				md.ForEach(metadata =>
				{
					dynamic cmd = md;
					var cellValue = sh.Cell(row, colsOrder[metadata.Title]).Value;
					cmd.SetModelPropertyValue(item, cmd.CellValue(cellValue));
				});
				res.Values.Add(item);
			}, headerCell.Address.RowNumber + 1, GetRowsCount(sh));
			return res;
		}

		private IXLCell GetStartCell(IXLWorksheet sheet)
		{
			return sheet.RangeUsed().FirstCell();
		}

		private int GetColumnsCount(IXLWorksheet sheet)
		{
			return sheet.RangeUsed().ColumnCount();
		}

		private int GetRowsCount(IXLWorksheet sheet)
		{
			return sheet.RangeUsed().RowCount();
		}

		private IXLCell SearchHeader(IXLWorksheet sheet, List<ExcelColumnImportMetadata> mdList)
		{
			var startCell = GetStartCell(sheet);
			var startRow = startCell.Address.RowNumber;
			var startCol = startCell.Address.ColumnNumber;
			var headers = mdList.Select(md => md.Title).ToList();
			var colsCount = GetColumnsCount(sheet);
			var rowsCount = GetRowsCount(sheet);

			var headerRow = -1;
			var headerCol = -1;
			ProcessRows(row =>
			{
				if (headerRow > 0) return;
				var isHeaderRow = true;
				ProcessCols((i, j) =>
				{
					var cellVal = sheet.Cell(i, j).GetValue<string>();
					if (!isHeaderRow || string.IsNullOrEmpty(cellVal))
						return;
					var isCellContainsAnyHeader = headers.Any(h => h.Equals(cellVal, StringComparison.InvariantCultureIgnoreCase));
					if (headerCol == -1 && isCellContainsAnyHeader)
						headerCol = j;
					if (!isCellContainsAnyHeader)
						isHeaderRow = false;
				}, row, startCol, colsCount);
				if (isHeaderRow) headerRow = row + 1;
			}, startRow, rowsCount);

			return sheet.Cell(headerRow, startCol);
		}

		private Dictionary<string, int> GetColumnsOrder(IXLWorksheet sheet, List<ExcelColumnImportMetadata> mdList)
		{
			var res = new Dictionary<string, int>();
			var headerCell = SearchHeader(sheet, mdList);
			var startCell = GetStartCell(sheet);
			var startCol = startCell.Address.ColumnNumber;
			var colsCount = GetColumnsCount(sheet);
			var headers = mdList.Select(md => md.Title).ToList();
			ProcessCols((i, j) =>
			{
				var cellVal = sheet.Cell(i, j).GetValue<string>();
				if (string.IsNullOrEmpty(cellVal)) return;
				var targetHeader = headers.FirstOrDefault(h => h.Equals(cellVal, StringComparison.InvariantCultureIgnoreCase));
				if (targetHeader != null && !res.ContainsKey(targetHeader))
					res.Add(targetHeader, j);
			}, headerCell.Address.RowNumber, startCol, colsCount);
			return res;
		}

		private void ProcessRows(Action<int> action, int startRow, int rowsCount)
		{
			for (var i = startRow; i <= rowsCount; i++) action(i);
		}

		private void ProcessCols(Action<int, int> action, int row, int startColumn, int columnsCount)
		{
			for (var i = startColumn; i <= startColumn + columnsCount; i++) action(row, i);
		}
	}
}