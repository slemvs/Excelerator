using System;
using System.IO;

namespace Excelerator.Examples.Export
{
	public class ExportExampleBase
	{
		protected readonly string ExportExamplesPath = "Examples\\Export";
		
		protected void Iterate(Action<int> action, int maxValue)
		{
			for (var i = 0; i < maxValue; i++)
				action(i);
		}
	}
}
