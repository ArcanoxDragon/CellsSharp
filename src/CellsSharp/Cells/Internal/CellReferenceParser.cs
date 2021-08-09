using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace CellsSharp.Cells.Internal
{
	static class CellReferenceParser
	{
		public static readonly TimeSpan MatchTimeout = TimeSpan.FromSeconds(1); // Should never take anywhere near this long to parse

		private static readonly Regex CellRangeRegex = new(@"^(?<First> \$?[A-Z]+ \$?\d+ )(?<Last> :\$?[A-Z]+ \$?\d+ )?$",
														   RegexOptions.Compiled | RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace,
														   MatchTimeout);

		private static readonly Regex CellAddressRegex = new(@"^(?<Column> [A-Z]+ ) (?<Row> \d+ )$",
															 RegexOptions.Compiled | RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace,
															 MatchTimeout);

		private static readonly Regex CellAddressWithTypeRegex = new(@"^(?<CAbs> \$ )? (?<Column> [A-Z]+ ) (?<RAbs> \$ )? (?<Row> \d+ )$",
																	 RegexOptions.Compiled | RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace,
																	 MatchTimeout);

		internal static ICellReference ParseCellReference(string cellReferenceText)
		{
			if (string.IsNullOrEmpty(cellReferenceText))
				return CellReference.Empty;

#if NET5_0_OR_GREATER
			var cellRanges = cellReferenceText.Split(',', StringSplitOptions.RemoveEmptyEntries);
#else
			var cellRanges = cellReferenceText.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
#endif
			var validRanges = new List<CellRange>();

			foreach (var cellRangeText in cellRanges)
			{
				if (TryParseCellRange(cellRangeText, out var cellRange))
					validRanges.Add(cellRange);
			}

			if (!validRanges.Any())
				// Don't want to make a new instance if there are no valid ranges; should just use Empty singleton instead
				return CellReference.Empty;

			if (validRanges.Count == 1)
			{
				var range = validRanges.Single();

				if (range.IsSingleCell)
					return range.TopLeft;

				return range;
			}

			return new CellReference(validRanges);
		}

		internal static bool TryParseCellRange(string cellRangeText, out CellRange cellRange)
		{
			cellRange = CellRange.Empty;

			var match = CellRangeRegex.Match(cellRangeText.Trim());

			if (!match.Success)
				return false;

			if (!TryParseCellAddress(match.Groups["First"].Value, out var firstAddress))
				return false;

			if (!match.Groups["Last"].Success)
			{
				cellRange = new CellRange(firstAddress, firstAddress);
				return true;
			}

			if (!TryParseCellAddress(match.Groups["Last"].Value.TrimStart(':'), out var lastAddress))
				return false;

			// Make sure we build a range where Start is in the top-left and End is in the bottom-right
			var (startRow, startColumn) = firstAddress;
			var (endRow, endColumn) = lastAddress;

			firstAddress = new CellAddress(Math.Min(startRow, endRow), Math.Min(startColumn, endColumn));
			lastAddress = new CellAddress(Math.Max(startRow, endRow), Math.Max(startColumn, endColumn));
			cellRange = new CellRange(firstAddress, lastAddress);

			return true;
		}

		internal static bool TryParseCellAddress(string cellAddressText, out CellAddress cellAddress)
		{
			cellAddress = CellAddress.None;

			if (string.IsNullOrEmpty(cellAddressText))
				return false;

			var match = CellAddressRegex.Match(cellAddressText.Trim());

			if (!match.Success)
				return false;

			var rowText = match.Groups["Row"].Value;
			var columnText = match.Groups["Column"].Value;

			if (!TryParseColumnName(columnText, out var columnIndex))
				return false;

			if (!uint.TryParse(rowText, out var rowIndex) || rowIndex is 0 or > Constants.ExcelMaxRows)
				return false;

			cellAddress = new CellAddress(rowIndex, columnIndex);
			return true;
		}

		internal static bool TryParseCellAddress(string cellAddressText, out CellAddress cellAddress, out AddressType addressType)
		{
			cellAddress = CellAddress.None;
			addressType = AddressType.Relative;

			if (string.IsNullOrEmpty(cellAddressText))
				return false;

			var match = CellAddressWithTypeRegex.Match(cellAddressText.Trim());

			if (!match.Success)
				return false;

			var rowText = match.Groups["Row"].Value;
			var columnText = match.Groups["Column"].Value;
			var isRowAbsolute = match.Groups["RAbs"].Success;
			var isColumnAbsolute = match.Groups["CAbs"].Success;

			if (!TryParseColumnName(columnText, out var columnIndex))
				return false;

			if (!uint.TryParse(rowText, out var rowIndex) || rowIndex is 0 or > Constants.ExcelMaxRows)
				return false;

			cellAddress = new CellAddress(rowIndex, columnIndex);
			addressType = CellAddress.GetAddressType(isRowAbsolute, isColumnAbsolute);
			return true;
		}

		internal static string ColumnIndexToName(uint columnIndex)
		{
			if (columnIndex is 0 or > 16384)
				throw new ArgumentOutOfRangeException(nameof(columnIndex), $"{nameof(columnIndex)} must be between 1 and 16384");

			var dividend = columnIndex;
			var columnName = string.Empty;

			while (dividend > 0)
			{
				var remainder = ( dividend - 1 ) % 26;

				columnName = (char) ( 'A' + remainder ) + columnName;
				dividend = ( dividend - remainder ) / 26;
			}

			return columnName;
		}

		private static bool TryParseColumnName(string columnName, out uint columnIndex)
		{
			columnIndex = default;

			var letters = columnName.ToUpper().Where(c => c is >= 'A' and <= 'Z').ToArray();

			if (letters.Length is 0 or > 3)
				return false;

			columnIndex = letters.Aggregate(0u, (sum, letter) => sum * 26 + (uint) ( letter - 'A' + 1 ));
			return columnIndex is > 0 and <= Constants.ExcelMaxColumns;
		}
	}
}