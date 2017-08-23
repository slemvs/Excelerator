using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;
using Excelerator.Common.Export.Metadata;
using Excelerator.Common.Models;
using Excelerator.Examples.Export.Npoi.Model;
using Excelerator.Examples.Extensions;
using Excelerator.Export;

namespace Excelerator.Examples.Export.Npoi
{
	public class NpoiTemplateExportExample : ExportExampleBase
	{
		private readonly IExcelGenerator<NpoiExampleModel> _npoiGenerator;

		public NpoiTemplateExportExample(IExcelGenerator<NpoiExampleModel> npoiGenerator)
		{
			_npoiGenerator = npoiGenerator;
		}

		public void Execute()
		{
			var wsMd = new WorksheetMetadata<NpoiExampleModel>
			{
				Name = "Example1",
				StartColumn = 1,
				StartRow = 19,
				FormatAsTable = true,
				GenerateHeader = false,
				ColumnsMetadata = new List<ColumnMetadata<NpoiExampleModel>>
				{
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop1", Value = _ => _.Prop1, ColumnAddress = new ColumnAddress("A")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop2", Value = _ => _.Prop2, ColumnAddress = new ColumnAddress("U")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop3", Value = _ => _.Prop3, ColumnAddress = new ColumnAddress("AO")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop4", Value = _ => _.Prop4, ColumnAddress = new ColumnAddress("CC")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop5", Value = _ => _.Prop5, ColumnAddress = new ColumnAddress("CN")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop6", Value = _ => _.Prop6, ColumnAddress = new ColumnAddress("DA")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop7", Value = _ => _.Prop7, ColumnAddress = new ColumnAddress("DL")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop8", Value = _ => _.Prop8, ColumnAddress = new ColumnAddress("DW")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop9", Value = _ => _.Prop9, ColumnAddress = new ColumnAddress("EJ")},
					new ColumnMetadata<NpoiExampleModel> {Header = "Prop10", Value = _ => _.Prop10, ColumnAddress = new ColumnAddress("EX")},
				}
			};

			#region Fill data
			var data = new List<NpoiExampleModel>();
			Iterate(i => data.Add(new NpoiExampleModel
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
				
			}), 20);
			#endregion s
			var template = new FileInfo("c:\\tmp\\График-отпусков.xls");
			using (var ms = new MemoryStream())
			{
				using (var fs = template.OpenRead())
				{
					fs.CopyTo(ms);
					_npoiGenerator.Generate(ms, wsMd, data)
						.SaveToFile($"{ExportExamplesPath}\\example1.xls");
				}

			}

		}
	}
}
