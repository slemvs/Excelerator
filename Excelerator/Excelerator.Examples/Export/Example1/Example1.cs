using Excelerator.Examples.Export.Example1.Model;
using Excelerator.Export;
using System;
using System.Collections.Generic;

namespace Excelerator.Examples.Export.Example1
{
	public class Example1 : ExportExampleBase
	{
		private readonly IExcelGenerator<Example1Model> _generator;

		public Example1(IExcelGenerator<Example1Model> generator)
		{
			_generator = generator;
		}

		public void Execute()
		{
			var metadata = new List<ExcelColumnMetadata<Example1Model>>()
			{
				new ExcelColumnMetadata<Example1Model>() {Header = "Prop1", Value = _ => _.Prop1},
				new ExcelColumnMetadata<Example1Model>() {Header = "Prop2", Value = _ => _.Prop2},
			};
			var data = new List<ExcelRowModel<Example1Model>>();
			Iterate(i => data.Add(new ExcelRowModel<Example1Model>()
			{
				Data = new Example1Model() { Prop1 = $"prop1_{i}", Prop2 = $"prop2_{i}" }
			}), 20);

			_generator.WorksheetName = "Example1";
			SaveToFile(_generator.Generate(metadata, data), $"{ExportExamplesPath}\\example1.xlsx");
		}

		#region Helpers
		private void Iterate(Action<int> action, int maxValue)
		{
			for (int i = 0; i < maxValue; i++)
			{
				action(i);
			}
		}
		#endregion
	}
}
