using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelerator.NPOI.Extensions
{
	public static class AddressExtensions
	{
		internal static readonly NumberStyles NumberStyle = NumberStyles.Float;
		internal static readonly CultureInfo ParseCulture = CultureInfo.InvariantCulture;
		private static readonly string[] letters = new string[26] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
		public static int GetColumnNumberFromLetter(this string columnLetter)
		{
			if (string.IsNullOrEmpty(columnLetter))
				throw new ArgumentNullException("columnLetter");
			columnLetter = columnLetter.ToUpper();
			if ((int)columnLetter[0] <= 57)
				return int.Parse(columnLetter, NumberStyle, ParseCulture);
			int num = 0;
			for (int index = 0; index < columnLetter.Length; ++index)
				num = num * 26 + ((int)columnLetter[index] - 65 + 1);
			return num-1;
		}

		public static string GetColumnLetterFromNumber(this int columnNumber)
		{
			--columnNumber;
			if (columnNumber <= 25)
				return letters[columnNumber];
			return GetColumnLetterFromNumber(columnNumber / 26) + GetColumnLetterFromNumber(columnNumber % 26 + 1);
		}
	}
}
