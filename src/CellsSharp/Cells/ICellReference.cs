using System;
using System.Collections.Generic;
using JetBrains.Annotations;

namespace CellsSharp.Cells
{
	[PublicAPI]
	public interface ICellReference : IEnumerable<CellAddress>, IEquatable<ICellReference>
	{
		/// <summary>
		/// Gets whether or not this <see cref="ICellReference"/> is a valid cell reference.
		/// </summary>
		bool IsValid { get; }

		/// <summary>
		/// Gets whether or not this <see cref="ICellReference"/> refers to exactly one cell.
		/// </summary>
		bool IsSingleCell { get; }

		/// <summary>
		/// Gets the cell address within this <see cref="ICellReference"/> that is closest to the
		/// top-left corner of the worksheet.
		/// </summary>
		CellAddress TopLeft { get; }

		/// <summary>
		/// Returns whether or not the provided <paramref name="cellAddress"/> falls within this <see cref="ICellReference"/>.
		/// </summary>
		bool Contains(CellAddress cellAddress);

		/// <summary>
		/// Returns whether or not this <see cref="ICellReference"/> intersects with the provided <paramref name="cellReference"/>.
		/// </summary>
		bool IntersectsWith(ICellReference cellReference);

		/// <summary>
		/// Returns whether or not the provided <paramref name="cellReference"/> is fully contained within this <see cref="ICellReference"/>
		/// </summary>
		bool FullyContains(ICellReference cellReference);
	}
}