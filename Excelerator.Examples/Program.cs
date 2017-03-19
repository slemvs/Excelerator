using System.IO;
using Excelerator.Examples.DependencyResolver;
using Excelerator.Examples.Export.Example1;
using Excelerator.Examples.Import.Example1;
using Ninject;

namespace Excelerator.Examples
{
	public static class Program
	{
		private static IKernel _kernel;

		public static void Main(string[] args)
		{
			CheckDirs();
			_kernel = new StandardKernel(new DefaultModule());

			Get<ExportExample1>().Execute();

			//Get<Example2>().Execute();
			//Get<Example1>().Execute();
		}

		#region Helpers

		private static T Get<T>()
		{
			return _kernel.Get<T>();
		}


		private static void CheckDirs()
		{
			if (!Directory.Exists("Examples\\Export")) Directory.CreateDirectory("Examples\\Export");
			if (!Directory.Exists("Examples\\Import")) Directory.CreateDirectory("Examples\\Import");
		}

		#endregion
	}
}