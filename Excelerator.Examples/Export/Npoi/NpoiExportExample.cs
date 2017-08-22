using System;
using System.Collections.Generic;
using Excelerator.Common.Export.Metadata;
using Excelerator.Enums;
using Excelerator.Examples.Export.Npoi.Model;
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
				StartColumn = 5,
				StartRow = 3,
				FormatAsTable = true,
				ColumnsMetadata = new List<ColumnMetadata<NpoiExampleModel>>
				{
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop1", Value = _ => _.Prop1},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop2", Value = _ => _.Prop2}
				}
			};

			var data = new List<NpoiExampleModel>();
			Iterate(i => data.Add(new NpoiExampleModel { Prop1 = $"prop1_{i}", Prop2 = $"prop2_{i}" }), 20);
			SaveToFile(_npoiGenerator.Generate(wsMd, data), $"{ExportExamplesPath}\\example1.xls");
		}


		private void Iterate(Action<int> action, int maxValue)
		{
			for (var i = 0; i < maxValue; i++)
				action(i);
		}
	}
}