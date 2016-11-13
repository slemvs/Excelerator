using Excelerator.Examples.Export.Example1.Generator;
using Excelerator.Examples.Export.Example1.Model;
using Excelerator.Export;
using Ninject.Modules;

namespace Excelerator.Examples.DependencyResolver
{
	public class DefaultModule : NinjectModule
	{
		public override void Load()
		{
			Kernel.Rebind<IExcelGenerator<Example1Model>>().To<Example1Generator>();
		}
	}
}
