using CellsSharp.Cells;
using JetBrains.Annotations;

namespace CellsSharp.Worksheets
{
	[PublicAPI]
	public interface ISheetView
	{
		/// <summary>
		/// Gets or sets the text of the cells in this <see cref="ISheetView"/>.
		/// </summary>
		string CellText { get; set; }

		/// <summary>
		/// Gets or sets the numeric value of the cells in this <see cref="ISheetView"/>.
		/// </summary>
		double CellValue { get; set; }

		/// <summary>
		/// Gets whether or not the cell range represented by this <see cref="ISheetView"/>'s
		/// <see cref="CellReference"/> is a merged cell range.
		/// </summary>
		/// <remarks>
		/// <para>
		/// If this <see cref="ISheetView"/>'s <see cref="CellReference"/> is not a single-range
		/// cell reference, or if it only refers to a single cell, this property is false.
		/// </para>
		/// <para>
		/// If this <see cref="ISheetView"/>'s <see cref="CellReference"/> refers to a portion
		/// of a merged cell range, or a range which overlaps with a merge cell range, but the
		/// range is not exactly equal to the merged range, this property is false.
		/// </para>
		/// </remarks>
		bool IsMerged { get; }

		/// <summary>
		/// Clears the values of all cells in this <see cref="ISheetView"/>.
		/// </summary>
		void ClearValue();

		/// <summary>
		/// Clears the formatting of all cells in this <see cref="ISheetView"/>.
		/// </summary>
		void ClearFormatting();

		/// <summary>
		/// Clears the values and formatting of all cells in this <see cref="ISheetView"/>.
		/// </summary>
		void ClearAll();

		/// <summary>
		/// Merges all cell ranges in this <see cref="ISheetView"/>'s <see cref="CellReference"/>
		/// such that each range becomes a single cell.
		/// </summary>
		void Merge();

		/// <summary>
		/// For each cell range in this <see cref="ISheetView"/>'s <see cref="CellReference"/>,
		/// unmerges any merged cells that intersect with that range.
		/// </summary>
		void Unmerge();
	}
}