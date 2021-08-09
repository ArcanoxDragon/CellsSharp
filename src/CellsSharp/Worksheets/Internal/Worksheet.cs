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
			IStringTable strings
		) : base(saveLoadHandlers, changeNotifier)
		{
			// For new worksheets, they should immediately be marked as dirty.
			// For existing ones, Load() will be called which in turn calls
			// LoadParts(), which marks the change notifier as clean.
			changeNotifier.NotifyOfChange(this, worksheetPart);

			SheetData = sheetData;
			Strings = strings;
		}

		private SheetData    SheetData { get; }
		private IStringTable Strings   { get; }

		#region IWorksheetImpl

		/// <inheritdoc />
		public ISheetView this[ICellReference cellReference] => new SheetView(SheetData, Strings, cellReference);

		/// <inheritdoc />
		public ISheetView this[string cellReference] => new SheetView(SheetData, Strings, CellReference.Parse(cellReference));

#if !NET5_0_OR_GREATER
		/// <inheritdoc />
		public ISheetView this[uint row, uint column] => this[new CellAddress(row, column)];

		/// <inheritdoc />
		public ISheetView this[uint topLeftRow, uint topLeftColumn, uint bottomRightRow, uint bottomRightColumn]
			=> this[new CellRange(new CellAddress(topLeftRow, topLeftColumn),
								  new CellAddress(bottomRightRow, bottomRightColumn))];
#endif

		/// <inheritdoc />
		public void Save()
		{
			CheckDisposed();
			SaveParts();
		}

		/// <inheritdoc />
		public void Load()
		{
			CheckDisposed();
			LoadParts();
		}

		#endregion
	}
}