using System;
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

		/// <inheritdoc cref="this[ICellReference]"/>
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
			=> this[new CellRange(topLeftRow, topLeftColumn, bottomRightRow, bottomRightColumn)];
#else
		ISheetView this[uint topLeftRow, uint topLeftColumn, uint bottomRightRow, uint bottomRightColumn] { get; }
#endif

		/// <summary>
		/// Opens an <see cref="ISheetView"/> for the cells represented by the provided
		/// <paramref name="cellReference"/> and calls <paramref name="action"/> with the
		/// <see cref="ISheetView"/>.
		/// </summary>
#if NET5_0_OR_GREATER
		void ForRange(ICellReference cellReference, Action<ISheetView> action) => action(this[cellReference]);
#else
		void ForRange(ICellReference cellReference, Action<ISheetView> action);
#endif

		/// <inheritdoc cref="ForRange(ICellReference,Action{ISheetView})"/>
#if NET5_0_OR_GREATER
		void ForRange(string cellReference, Action<ISheetView> action) => action(this[cellReference]);
#else
		void ForRange(string cellReference, Action<ISheetView> action);
#endif

		/// <summary>
		/// Opens an <see cref="ISheetView"/> for the cell at the provided <paramref name="row"/>
		/// and <paramref name="column"/> and calls <paramref name="action"/> with the <see cref="ISheetView"/>.
		/// </summary>
#if NET5_0_OR_GREATER
		void ForRange(uint row, uint column, Action<ISheetView> action) => action(this[row, column]);
#else
		void ForRange(uint row, uint column, Action<ISheetView> action);
#endif

		/// <summary>
		/// Opens an <see cref="ISheetView"/> for the cells in the range specified by
		/// <paramref name="topLeftRow"/>, <paramref name="topLeftColumn"/>,
		/// <paramref name="bottomRightRow"/>, and <paramref name="bottomRightColumn"/>
		/// and calls <paramref name="action"/> with the <see cref="ISheetView"/>.
		/// </summary>
#if NET5_0_OR_GREATER
		void ForRange(uint topLeftRow, uint topLeftColumn,
					  uint bottomRightRow, uint bottomRightColumn,
					  Action<ISheetView> action)
			=> action(this[topLeftRow, topLeftColumn, bottomRightRow, bottomRightColumn]);
#else
		void ForRange(uint topLeftRow, uint topLeftColumn,
					  uint bottomRightRow, uint bottomRightColumn,
					  Action<ISheetView> action);
#endif

		/// <summary>
		/// Saves the worksheet and all related parts.
		/// </summary>
		void Save();
	}
}