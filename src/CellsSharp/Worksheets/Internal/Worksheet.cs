#if !NET5_0_OR_GREATER
using System;
#endif

using System.Collections.Generic;
using CellsSharp.Cells;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using CellsSharp.Workbooks;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Worksheets.Internal
{
	sealed class Worksheet : PartRootBase, IWorksheetImpl
	{
		public Worksheet(
			IEnumerable<IWorksheetSaveLoadHandler> saveLoadHandlers,
			IChangeNotifier changeNotifier,
			// ReSharper disable once SuggestBaseTypeForParameterInConstructor
			WorksheetPart worksheetPart,
			SheetData sheetData,
			IStringTable strings,
			MergedCellsCollection mergedCells
		) : base(saveLoadHandlers, changeNotifier)
		{
			// For new worksheets, they should immediately be marked as dirty.
			// For existing ones, Load() will be called which in turn calls
			// LoadParts(), which marks the change notifier as clean.
			changeNotifier.NotifyOfChange(this, worksheetPart);

			SheetData = sheetData;
			Strings = strings;
			MergedCells = mergedCells;
		}

		private IStringTable          Strings     { get; }
		private SheetData             SheetData   { get; }
		private MergedCellsCollection MergedCells { get; }

		#region IWorksheetImpl

		public ISheetView this[ICellReference cellReference] => new SheetView(cellReference, Strings, SheetData, MergedCells);

		public ISheetView this[string cellReference] => new SheetView(CellReference.Parse(cellReference), Strings, SheetData, MergedCells);

#if !NET5_0_OR_GREATER
		public ISheetView this[uint row, uint column] => this[new CellAddress(row, column)];

		public ISheetView this[uint topLeftRow, uint topLeftColumn, uint bottomRightRow, uint bottomRightColumn]
			=> this[new CellRange(topLeftRow, topLeftColumn, bottomRightRow, bottomRightColumn)];

		public void ForRange(ICellReference cellReference, Action<ISheetView> action)
			=> action(this[cellReference]);

		public void ForRange(string cellReference, Action<ISheetView> action)
			=> action(this[cellReference]);

		public void ForRange(uint row, uint column, Action<ISheetView> action)
			=> action(this[row, column]);

		public void ForRange(uint topLeftRow, uint topLeftColumn, uint bottomRightRow, uint bottomRightColumn, Action<ISheetView> action)
			=> action(this[topLeftRow, topLeftColumn, bottomRightRow, bottomRightColumn]);
#endif

		public void Save()
		{
			CheckDisposed();
			SaveParts();
		}

		public void Load()
		{
			CheckDisposed();
			LoadParts();
		}

		#endregion
	}
}