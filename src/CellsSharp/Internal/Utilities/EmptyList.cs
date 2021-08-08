using System.Collections.Generic;

namespace CellsSharp.Internal.Utilities
{
	sealed class EmptyList<T>
	{
		internal static readonly IList<T> Instance = new List<T>();
	}
}