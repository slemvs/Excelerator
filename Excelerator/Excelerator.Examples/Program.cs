using Excelerator.Examples.DependencyResolver;
using Excelerator.Examples.Export.Example1;
using Ninject;
using System.IO;

namespace Excelerator.Examples
{
	public class Program
	{

		private static IKernel _kernel;
		static void Main(string[] args)
		{
			CheckDirs();
			_kernel = new StandardKernel(new DefaultModule());

			// Example1
			Get<Example1>().Execute();
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
