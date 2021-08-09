using CellsSharp.Cells;
using JetBrains.Annotations;

namespace CellsSharp.Worksheets
{
	[PublicAPI]
	public interface IWorksheet
	{
		/// <summary>
		/// Gets an <see cref="ISheetView"/> instance representing a view of the worksheet in the
		/// range provided by <paramref name="cellReference"/>.
		/// </summary>
		ISheetView this[ICellReference cellReference] { get; }

		/// <summary>
		/// Gets an <see cref="ISheetView"/> instance representing a view of the worksheet in the
		/// range provided by <paramref name="cellReference"/>.
		/// </summary>
		ISheetView this[string cellReference] { get; }

		/// <summary>
		/// Gets an <see cref="ISheetView"/> instance representing a view of the cell at the
		/// provided <paramref name="row"/> and <paramref name="column"/>.
		/// </summary>
#if NET5_0_OR_GREATER
		ISheetView this[uint row, uint column] => this[new CellAddress(row, column)];
#else
		ISheetView this[uint row, uint column] { get; }
#endif

		/// <summary>
		/// Gets an <see cref="ISheetView"/> instance representing a view of the cells in the
		/// range specified by <paramref name="topLeftRow"/>, <paramref name="topLeftColumn"/>,
		/// <paramref name="bottomRightRow"/>, and <paramref name="bottomRightColumn"/>.
		/// </summary>
#if NET5_0_OR_GREATER
		ISheetView this[uint topLeftRow, uint topLeftColumn, uint bottomRightRow, uint bottomRightColumn]
			=> this[new CellRange(new CellAddress(topLeftRow, topLeftColumn),
								  new CellAddress(bottomRightRow, bottomRightColumn))];
#else
		ISheetView this[uint topLeftRow, uint topLeftColumn, uint bottomRightRow, uint bottomRightColumn] { get; }
#endif

		/// <summary>
		/// Saves the worksheet and all related parts.
		/// </summary>
		void Save();
	}
}