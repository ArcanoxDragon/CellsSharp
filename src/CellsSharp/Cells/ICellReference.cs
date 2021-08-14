using System;
using JetBrains.Annotations;

namespace CellsSharp.Cells
{
	[PublicAPI]
	public interface ICellReference : IEquatable<ICellReference>
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
		/// Gets the cell address within this <see cref="ICellReference"/> that is closest to the
		/// top-right corner of the worksheet.
		/// </summary>
		CellAddress TopRight { get; }

		/// <summary>
		/// Gets the cell address within this <see cref="ICellReference"/> that is closest to the
		/// bottom-left corner of the worksheet.
		/// </summary>
		CellAddress BottomLeft { get; }

		/// <summary>
		/// Gets the cell address within this <see cref="ICellReference"/> that is closest to the
		/// bottom-right corner of the worksheet.
		/// </summary>
		CellAddress BottomRight { get; }

		/// <summary>
		/// Gets whether or not this <see cref="ICellReference"/> refers to an empty range of cells.
		/// </summary>
		bool IsEmpty { get; }

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

		/// <summary>
		/// Returns a concrete <see cref="CellReference"/> that represents the same collection
		/// of cells as this <see cref="ICellReference"/>.
		/// </summary>
		CellReference ToCellReference();
	}
}