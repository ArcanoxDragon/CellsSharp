using System.Collections.Generic;
using System.Linq;

namespace CellsSharp.Extensions
{
	static class LinqExtensions
	{
		/// <summary>
		/// Flattens a structure of nested enumerables into a single-dimensional enumerable by
		/// using SelectMany on each <typeparamref name="T"/> in <paramref name="source"/>.
		/// </summary>
		internal static IEnumerable<T> Flatten<T>(this IEnumerable<IEnumerable<T>> source)
			=> source.SelectMany(item => item);
	}
}