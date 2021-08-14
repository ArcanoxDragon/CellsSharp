using System.Collections.Generic;
using CellsSharp.Cells;

namespace CellsSharp.Extensions
{
	public static class CellReferenceExtensions
	{
		public static IEnumerator<CellAddress> GetEnumerator(this (CellAddress Start, CellAddress End) cellPair)
			=> ( (CellRange) cellPair ).GetEnumerator();
	}
}