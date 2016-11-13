using System.IO;

namespace Excelerator.Examples.Export
{
	public class ExportExampleBase
	{
		protected readonly string ExportExamplesPath = "Examples\\Export";
		public void SaveToFile(MemoryStream ms, string fileName)
		{
			var f = new FileInfo(fileName);
			//if (!f.Exists) f.cre File.Create(fileName);
			using (var fs = f. OpenWrite())
			using (var bw = new BinaryWriter(fs))
			{
				ms.Position = 0;
				bw.Write(ms.ToArray());
				bw.Flush();
				bw.Close();
			}
		}
	}
}
