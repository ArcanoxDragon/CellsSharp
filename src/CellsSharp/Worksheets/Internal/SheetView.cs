using System;
using CellsSharp.Cells;
using CellsSharp.Workbooks;

namespace CellsSharp.Worksheets.Internal
{
	sealed class SheetView : ISheetView
	{
		public SheetView(SheetData sheetData, IStringTable strings, ICellReference cellReference)
		{
			SheetData = sheetData;
			Strings = strings;
			CellReference = cellReference;
		}

		private SheetData      SheetData     { get; }
		private IStringTable   Strings       { get; }
		private ICellReference CellReference { get; }

		#region ISheetView

		public string CellText
		{
			get => FromFirstCell(address => {
				if (!SheetData.TryGetStringIndex(address, out var stringIndex))
					return null;
				if (!Strings.TryGetValue(stringIndex, out var text))
					return null;

				return text;
			}, string.Empty);
			set
			{
				uint stringIndex = Strings.GetOrInsertValue(value);

				ForAllCells(address => SheetData.PutStringIndex(address, stringIndex));
			}
		}

		public double CellValue
		{
			get => FromFirstCell(address => {
				if (!SheetData.TryGetNumericValue(address, out var numericValue))
					return default;

				return numericValue;
			}, default);
			set => ForAllCells(address => SheetData.PutNumericValue(address, value));
		}

		public void ClearValue() => ForAllCells(SheetData.ClearValue);

		public void ClearFormatting() => ForAllCells(SheetData.ClearStyleIndex);

		public void ClearAll() => ForAllCells(SheetData.ClearCell);

		#endregion

		private T FromFirstCell<T>(Func<CellAddress, T?> action, T defaultValue)
		{
			if (CellReference.IsEmpty)
				return defaultValue;

			var value = action(CellReference.TopLeft);

			return value ?? defaultValue;
		}

		private void ForAllCells(Action<CellAddress> action)
		{
			if (CellReference.IsEmpty)
				return;

			foreach (var address in CellReference)
				action(address);
		}
	}
}