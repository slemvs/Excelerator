using System;
using System.Collections.Generic;
using Excelerator.Common.Export.Metadata;
using Excelerator.Examples.Export.Example1.Model;
using Excelerator.Export;

namespace Excelerator.Examples.Export.Example1
{
	public class ExportExample1 : ExportExampleBase
	{
		private readonly IExcelGenerator<Example1Model> _generator;

		public ExportExample1(IExcelGenerator<Example1Model> generator)
		{
			_generator = generator;
		}

		public void Execute()
		{
			var wsMd = new WorksheetMetadata<Example1Model>
			{
				ColumnsMetadata = new List<ColumnMetadata<Example1Model>>
				{
					new ColumnMetadata<Example1Model> {Header = "Prop1", Value = _ => _.Prop1},
					new ColumnMetadata<Example1Model> {Header = "Prop2", Value = _ => _.Prop2}
				}
			};

			var data = new List<Example1Model>();
			Iterate(i => data.Add(new Example1Model {Prop1 = $"prop1_{i}", Prop2 = $"prop2_{i}"}), 20);
			_generator.WorksheetName = "Example1";
			SaveToFile(_generator.Generate(wsMd, data), $"{ExportExamplesPath}\\example1.xlsx");
		}

		#region Helpers

		private void Iterate(Action<int> action, int maxValue)
		{
			for (var i = 0; i < maxValue; i++)
				action(i);
		}

		#endregion
	}
}