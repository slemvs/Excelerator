using System;
using System.Collections.Generic;
using Excelerator.Common.Export.Metadata;
using Excelerator.Examples.Export.Example1.Model;
using Excelerator.Examples.Export.Example2.Model;
using Excelerator.Examples.Extensions;
using Excelerator.Export;

namespace Excelerator.Examples.Export.Example2
{
	public class ExportExample2 : ExportExampleBase
	{
		//private readonly IExcelGenerator<Example2Model> _generator;
		private readonly IExcelGenerator<Example21Model> _openXmlGenerator;

		public ExportExample2(IExcelGenerator<Example21Model> openXmlGenerator)
		{
			//_generator = generator;
			_openXmlGenerator = openXmlGenerator;
		}

		public void Execute()
		{
			var wsmd1 = new WorksheetMetadata<Example2Model>
			{
				ColumnsMetadata = new List<ColumnMetadata<Example2Model>>
				{
					new ColumnMetadata<Example2Model> {Header = "Prop1", Value = _ => _.Prop1},
					new ColumnMetadata<Example2Model> {Header = "Prop2", Value = _ => _.Prop2},
					new ColumnMetadata<Example2Model> {Header = "Prop3", Value = _ => _.Prop3},
					new ColumnMetadata<Example2Model> {Header = "Prop4", Value = _ => _.Prop4},
					new ColumnMetadata<Example2Model> {Header = "Prop5", Value = _ => _.Prop5},
					new ColumnMetadata<Example2Model> {Header = "Prop6", Value = _ => _.Prop6},
					new ColumnMetadata<Example2Model> {Header = "Prop7", Value = _ => _.Prop7},
					new ColumnMetadata<Example2Model> {Header = "Prop8", Value = _ => _.Prop8},
					new ColumnMetadata<Example2Model> {Header = "Prop9", Value = _ => _.Prop9},
					new ColumnMetadata<Example2Model> {Header = "Prop10", Value = _ => _.Prop10},
					new ColumnMetadata<Example2Model> {Header = "Prop11", Value = _ => _.Prop11},
					new ColumnMetadata<Example2Model> {Header = "Prop12", Value = _ => _.Prop12},
					new ColumnMetadata<Example2Model> {Header = "Prop13", Value = _ => _.Prop13},
					new ColumnMetadata<Example2Model> {Header = "Prop14", Value = _ => _.Prop14},
					new ColumnMetadata<Example2Model> {Header = "Prop15", Value = _ => _.Prop15},
					new ColumnMetadata<Example2Model> {Header = "Prop16", Value = _ => _.Prop16},
					new ColumnMetadata<Example2Model> {Header = "Prop17", Value = _ => _.Prop17},
					new ColumnMetadata<Example2Model> {Header = "Prop18", Value = _ => _.Prop18},
					new ColumnMetadata<Example2Model> {Header = "Prop19", Value = _ => _.Prop19},
					new ColumnMetadata<Example2Model> {Header = "Prop20", Value = _ => _.Prop20},
					new ColumnMetadata<Example2Model> {Header = "Prop21", Value = _ => _.Prop21},
					new ColumnMetadata<Example2Model> {Header = "Prop22", Value = _ => _.Prop22},
					new ColumnMetadata<Example2Model> {Header = "Prop23", Value = _ => _.Prop23},
					new ColumnMetadata<Example2Model> {Header = "Prop24", Value = _ => _.Prop24},
					new ColumnMetadata<Example2Model> {Header = "Prop25", Value = _ => _.Prop25},
					new ColumnMetadata<Example2Model> {Header = "Prop26", Value = _ => _.Prop26},
					new ColumnMetadata<Example2Model> {Header = "Prop27", Value = _ => _.Prop27},
					new ColumnMetadata<Example2Model> {Header = "Prop28", Value = _ => _.Prop28},
					new ColumnMetadata<Example2Model> {Header = "Prop29", Value = _ => _.Prop29},
					new ColumnMetadata<Example2Model> {Header = "Prop30", Value = _ => _.Prop30}
				}
			};
			var wsmd2 = new WorksheetMetadata<Example21Model>
			{
				ColumnsMetadata = new List<ColumnMetadata<Example21Model>>
				{
					new ColumnMetadata<Example21Model> {Header = "Prop1", Value = _ => _.Prop1},
					new ColumnMetadata<Example21Model> {Header = "Prop2", Value = _ => _.Prop2},
					new ColumnMetadata<Example21Model> {Header = "Prop3", Value = _ => _.Prop3},
					new ColumnMetadata<Example21Model> {Header = "Prop4", Value = _ => _.Prop4},
					new ColumnMetadata<Example21Model> {Header = "Prop5", Value = _ => _.Prop5},
					new ColumnMetadata<Example21Model> {Header = "Prop6", Value = _ => _.Prop6},
					new ColumnMetadata<Example21Model> {Header = "Prop7", Value = _ => _.Prop7},
					new ColumnMetadata<Example21Model> {Header = "Prop8", Value = _ => _.Prop8},
					new ColumnMetadata<Example21Model> {Header = "Prop9", Value = _ => _.Prop9},
					new ColumnMetadata<Example21Model> {Header = "Prop10", Value = _ => _.Prop10},
					new ColumnMetadata<Example21Model> {Header = "Prop11", Value = _ => _.Prop11},
					new ColumnMetadata<Example21Model> {Header = "Prop12", Value = _ => _.Prop12},
					new ColumnMetadata<Example21Model> {Header = "Prop13", Value = _ => _.Prop13},
					new ColumnMetadata<Example21Model> {Header = "Prop14", Value = _ => _.Prop14},
					new ColumnMetadata<Example21Model> {Header = "Prop15", Value = _ => _.Prop15},
					new ColumnMetadata<Example21Model> {Header = "Prop16", Value = _ => _.Prop16},
					new ColumnMetadata<Example21Model> {Header = "Prop17", Value = _ => _.Prop17},
					new ColumnMetadata<Example21Model> {Header = "Prop18", Value = _ => _.Prop18},
					new ColumnMetadata<Example21Model> {Header = "Prop19", Value = _ => _.Prop19},
					new ColumnMetadata<Example21Model> {Header = "Prop20", Value = _ => _.Prop20},
					new ColumnMetadata<Example21Model> {Header = "Prop21", Value = _ => _.Prop21},
					new ColumnMetadata<Example21Model> {Header = "Prop22", Value = _ => _.Prop22},
					new ColumnMetadata<Example21Model> {Header = "Prop23", Value = _ => _.Prop23},
					new ColumnMetadata<Example21Model> {Header = "Prop24", Value = _ => _.Prop24},
					new ColumnMetadata<Example21Model> {Header = "Prop25", Value = _ => _.Prop25},
					new ColumnMetadata<Example21Model> {Header = "Prop26", Value = _ => _.Prop26},
					new ColumnMetadata<Example21Model> {Header = "Prop27", Value = _ => _.Prop27},
					new ColumnMetadata<Example21Model> {Header = "Prop28", Value = _ => _.Prop28},
					new ColumnMetadata<Example21Model> {Header = "Prop29", Value = _ => _.Prop29},
					new ColumnMetadata<Example21Model> {Header = "Prop30", Value = _ => _.Prop30}
				}
			};

			var data = new List<Example2Model>();
			var data2 = new List<Example21Model>();
			var total = 1000000;
			for (var i = 0; i <= total; i++)
			{
			}


			Iterate(i =>
			{
				data.Add(GetModel(i));
				data2.Add(GetModel2(i));
				if (i % 100 == 0)
					Console.WriteLine(i);
			}, 1000000);

			//_generator.WorksheetName = "Example2";
			//SaveToFile(_generator.Generate(metadata, data), $"{ExportExamplesPath}\\example2.xlsx");
			//var ms = _generator.Generate(metadata, data);

			var ms = _openXmlGenerator.Generate(wsmd2, Iterate(GetModel2, 1000000));


			//Iterate(i =>
			//{

			//	if (i % 100 == 0) Console.WriteLine(i);
			//	yield return new ExcelRowModel<Example21Model>() { Data = GetModel2(i) };
			//}, 1000000));
			ms.SaveToFile($"{ExportExamplesPath}\\example21.xlsx");
		}

		#region Helpers

		private IEnumerable<Example21Model> Iterate(Func<int, Example21Model> action, int maxValue)
		{
			for (var i = 0; i < maxValue; i++)
				yield return action(i);
		}

		private void Iterate(Action<int> action, int maxValue)
		{
			for (var i = 0; i < maxValue; i++)
				action(i);
		}

		private Example2Model GetModel(int i)
		{
			return new Example2Model
			{
				Prop1 = $"prop1_{i}",
				Prop2 = $"prop2_{i}",
				Prop3 = $"prop3_{i}",
				Prop4 = $"prop4_{i}",
				Prop5 = $"prop5_{i}",
				Prop6 = $"prop6_{i}",
				Prop7 = $"prop7_{i}",
				Prop8 = $"prop8_{i}",
				Prop9 = $"prop9_{i}",
				Prop10 = $"prop10_{i}",
				Prop11 = $"prop11_{i}",
				Prop12 = $"prop12_{i}",
				Prop13 = $"prop13_{i}",
				Prop14 = $"prop14_{i}",
				Prop15 = $"prop15_{i}",
				Prop16 = $"prop16_{i}",
				Prop17 = $"prop17_{i}",
				Prop18 = $"prop18_{i}",
				Prop19 = $"prop19_{i}",
				Prop20 = $"prop20_{i}",
				Prop21 = $"prop21_{i}",
				Prop22 = $"prop22_{i}",
				Prop23 = $"prop23_{i}",
				Prop24 = $"prop24_{i}",
				Prop25 = $"prop25_{i}",
				Prop26 = $"prop26_{i}",
				Prop27 = $"prop27_{i}",
				Prop28 = $"prop28_{i}",
				Prop29 = $"prop29_{i}",
				Prop30 = $"prop30_{i}"
			};
		}

		private Example21Model GetModel2(int i)
		{
			return new Example21Model
			{
				Prop1 = $"prop1_{i}",
				Prop2 = $"prop2_{i}",
				Prop3 = $"prop3_{i}",
				Prop4 = $"prop4_{i}",
				Prop5 = $"prop5_{i}",
				Prop6 = $"prop6_{i}",
				Prop7 = $"prop7_{i}",
				Prop8 = $"prop8_{i}",
				Prop9 = $"prop9_{i}",
				Prop10 = $"prop10_{i}",
				Prop11 = $"prop11_{i}",
				Prop12 = $"prop12_{i}",
				Prop13 = $"prop13_{i}",
				Prop14 = $"prop14_{i}",
				Prop15 = $"prop15_{i}",
				Prop16 = $"prop16_{i}",
				Prop17 = $"prop17_{i}",
				Prop18 = $"prop18_{i}",
				Prop19 = $"prop19_{i}",
				Prop20 = $"prop20_{i}",
				Prop21 = $"prop21_{i}",
				Prop22 = $"prop22_{i}",
				Prop23 = $"prop23_{i}",
				Prop24 = $"prop24_{i}",
				Prop25 = $"prop25_{i}",
				Prop26 = $"prop26_{i}",
				Prop27 = $"prop27_{i}",
				Prop28 = $"prop28_{i}",
				Prop29 = $"prop29_{i}",
				Prop30 = $"prop30_{i}"
			};
		}

		#endregion
	}
}