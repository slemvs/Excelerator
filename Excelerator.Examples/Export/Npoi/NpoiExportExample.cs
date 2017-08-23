using System;
using System.Collections.Generic;
using Excelerator.Common.Export.Metadata;
using Excelerator.Common.Models;
using Excelerator.Enums;
using Excelerator.Examples.Export.Npoi.Model;
using Excelerator.Examples.Extensions;
using Excelerator.Export;

namespace Excelerator.Examples.Export.Npoi
{
	public class NpoiExportExample : ExportExampleBase
	{
		private readonly IExcelGenerator<NpoiExampleModel> _npoiGenerator;

		public NpoiExportExample(IExcelGenerator<NpoiExampleModel> npoiGenerator)
		{
			_npoiGenerator = npoiGenerator;
		}

		public void Execute()
		{
			var wsMd = new WorksheetMetadata<NpoiExampleModel>
			{
				Name = "Example1",
				StartColumn = 1,
				StartRow = 4,
				FormatAsTable = true,
				ColumnsMetadata = new List<ColumnMetadata<NpoiExampleModel>>
				{
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop1", Value = _ => _.Prop1, ColumnAddress = new ColumnAddress("B")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop2", Value = _ => _.Prop2, ColumnAddress = new ColumnAddress("E")}
				}
			};

			var data = new List<NpoiExampleModel>();
			Iterate(i => data.Add(new NpoiExampleModel { Prop1 = $"prop1_{i}", Prop2 = $"prop2_{i}" }), 20);
			_npoiGenerator.Generate(wsMd, data)
				.SaveToFile($"{ExportExamplesPath}\\example1.xls");
		}
	}
}