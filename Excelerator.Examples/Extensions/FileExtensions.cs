using System.IO;

namespace Excelerator.Examples.Extensions
{
	public static class FileExtensions
	{
		public static void SaveToFile(this MemoryStream ms, string fileName)
		{
			var f = new FileInfo(fileName);
			//if (!f.Exists) f.cre File.Create(fileName);
			using (var fs = f.OpenWrite())
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