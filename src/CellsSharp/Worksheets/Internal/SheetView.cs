using System;
using System.Collections.Generic;
using System.Linq;
using CellsSharp.Cells;
using CellsSharp.Workbooks;

namespace CellsSharp.Worksheets.Internal
{
	sealed class SheetView : ISheetView
	{
		#region Constructors

		public SheetView(
			ICellReference cellReference,
			IStringTable strings,
			SheetData sheetData,
			MergedCellsCollection mergedCells
		)
		{
			CellReference = cellReference.ToCellReference();
			Strings = strings;
			SheetData = sheetData;
			MergedCells = mergedCells;
		}

		#endregion

		#region Properties

		private CellReference         CellReference { get; }
		private IStringTable          Strings       { get; }
		private SheetData             SheetData     { get; }
		private MergedCellsCollection MergedCells   { get; }

		#endregion

		#region ISheetView Implementation

		public string CellText
		{
			get => FromFirstCell(address => {
				if (!SheetData.TryGetStringIndex(address, out var stringIndex))
					return null;

				return Strings.GetValue(stringIndex);
			}, string.Empty);
			set
			{
				if (string.IsNullOrEmpty(value))
				{
					ForAllCells(address => SheetData.ClearValue(address));
				}
				else
				{
					uint stringIndex = Strings.GetOrInsertValue(value);

					ForAllCells(address => SheetData.PutStringIndex(address, stringIndex));
				}
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

		public bool IsMerged
		{
			get
			{
				if (CellReference.Ranges.Count() != 1)
					return false;

				var cellRange = CellReference.Ranges.Single();

				if (cellRange.IsSingleCell)
					return false;

				return MergedCells.IsMerged(cellRange);
			}
		}

		public void ClearValue() => ForAllCells(SheetData.ClearValue);
		public void ClearFormatting() => ForAllCells(SheetData.ClearStyleIndex);
		public void ClearAll() => ForAllCells(SheetData.ClearCell);

		public void Merge()
		{
			foreach (var range in CellReference.Ranges)
				MergedCells.Merge(range);
		}

		public void Unmerge()
		{
			foreach (var range in CellReference.Ranges)
				MergedCells.Unmerge(range);
		}

		#endregion

		#region Private Methods

		private T FromFirstCell<T>(Func<CellAddress, T?> action, T defaultValue)
		{
			if (CellReference.IsEmpty)
				return defaultValue;

			var address = CellReference.TopLeft;

			if (MergedCells.IsMerged(address, out var mergedRange))
				address = mergedRange.TopLeft;

			var value = action(address);

			return value ?? defaultValue;
		}

		private void ForAllCells(Action<CellAddress> action)
		{
			if (CellReference.IsEmpty)
				return;

			var visited = new HashSet<CellAddress>();

			foreach (var address in CellReference)
			{
				if (visited.Contains(address))
					continue;

				if (MergedCells.IsMerged(address, out var mergedRange))
				{
					action(mergedRange.TopLeft);

					// Don't want to re-visit any other cells in the merged range
					foreach (var rangeAddress in mergedRange)
						visited.Add(rangeAddress);
				}
				else
				{
					action(address);
					visited.Add(address);
				}
			}
		}

		#endregion
	}
}