using Excelerator.Common.Import;
using Excelerator.Examples.Export.Example1.Generator;
using Excelerator.Examples.Export.Example1.Model;
using Excelerator.Examples.Export.Example2.Generator;
using Excelerator.Examples.Export.Example2.Model;
using Excelerator.Examples.Import.Example1.Importer;
using Excelerator.Examples.Import.Example1.Model;
using Excelerator.Export;
using Ninject.Modules;

namespace Excelerator.Examples.DependencyResolver
{
	public class DefaultModule : NinjectModule
	{
		public override void Load()
		{
			//Kernel.Rebind<IExcelGenerator<Example2Model>>().To<Example1Generator>();
			Kernel.Rebind<IExcelGenerator<Example21Model>>().To<Example2OpenXmlGenerator>();
			Kernel.Rebind<IExcelImporter<Example1RowModel>>().To<Example1Importer>();
		}
	}
}
