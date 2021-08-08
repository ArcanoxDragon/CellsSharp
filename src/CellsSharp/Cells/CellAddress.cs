using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using CellsSharp.Cells.Internal;
using JetBrains.Annotations;

namespace CellsSharp.Cells
{
	[PublicAPI]
	public sealed record CellAddress : ICellReference, IComparable<CellAddress>, IComparable
	{
		#region AddressType Enum

		[Flags]
		private enum AddressType : byte
		{
			Relative,
			RowAbsolute,
			ColumnAbsolute,
			Absolute,
		}

		private static AddressType GetAddressType(bool rowAbsolute, bool columnAbsolute)
			=> rowAbsolute
				   ? columnAbsolute
						 ? AddressType.Absolute
						 : AddressType.RowAbsolute
				   : columnAbsolute
					   ? AddressType.ColumnAbsolute
					   : AddressType.Relative;

		#endregion

		public static readonly CellAddress None = new();

		private CellAddress() { }

		public CellAddress(uint row, uint column, bool rowAbsolute = false, bool columnAbsolute = false)
		{
			if (row is 0 or > Constants.ExcelMaxRows)
				throw new ArgumentOutOfRangeException(nameof(row));
			if (column is 0 or > Constants.ExcelMaxColumns)
				throw new ArgumentOutOfRangeException(nameof(column));

			Row = row;
			Column = column;
			Type = GetAddressType(rowAbsolute, columnAbsolute);
		}

		private AddressType Type { get; }

		public uint Row    { get; }
		public uint Column { get; }

		public bool IsRowAbsolute    => ( Type & AddressType.RowAbsolute ) > 0;
		public bool IsColumnAbsolute => ( Type & AddressType.ColumnAbsolute ) > 0;

		public string ColumnName => IsValid ? CellReferenceParser.ColumnIndexToName(Column) : "Invalid";

		public CellAddress AsRelative()
		{
			if (Type == AddressType.Relative)
				return this;

			return new CellAddress(Row, Column);
		}

		public CellAddress AsAbsolute(bool rowAbsolute = true, bool columnAbsolute = true)
		{
			if (Type == GetAddressType(rowAbsolute, columnAbsolute))
				return this;

			return new CellAddress(Row, Column, rowAbsolute, columnAbsolute);
		}

		#region ICellReference

		/// <inheritdoc />
		public bool IsValid => Row > 0 && Column > 0;

		/// <inheritdoc />
		public bool IsSingleCell => true;

		/// <inheritdoc />
		public CellAddress TopLeft => this;

		/// <inheritdoc />
		public bool Contains(CellAddress cellAddress) => Equals(cellAddress);

		/// <inheritdoc />
		public bool IntersectsWith(ICellReference cellReference) => cellReference.Contains(this);

		/// <inheritdoc />
		public bool FullyContains(ICellReference cellReference) => Equals(cellReference);

		#endregion

		#region IComparable<CellAddress>

		/// <inheritdoc />
		public int CompareTo(CellAddress? other)
		{
			if (ReferenceEquals(this, other))
				return 0;
			if (ReferenceEquals(null, other))
				return 1;

			var rowComparison = Row.CompareTo(other.Row);

			if (rowComparison != 0)
				return rowComparison;

			return Column.CompareTo(other.Column);
		}

		/// <inheritdoc />
		public int CompareTo(object? obj)
		{
			if (ReferenceEquals(this, obj))
				return 0;
			if (ReferenceEquals(null, obj))
				return 1;
			if (GetType() != obj.GetType())
				return 1;

			return CompareTo((CellAddress) obj);
		}

		#endregion

		#region IEnumerable<CellAddress>

		/// <inheritdoc />
		public IEnumerator<CellAddress> GetEnumerator() => Enumerable.Repeat(this, 1).GetEnumerator();

		/// <inheritdoc />
		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

		#endregion

		#region IEquatable<ICellReference>

		/// <inheritdoc />
		public bool Equals(ICellReference? other)
		{
			if (ReferenceEquals(null, other))
				return false;
			if (ReferenceEquals(this, other))
				return true;

			return other.IsSingleCell && Equals(other.TopLeft);
		}

		#endregion

		/// <inheritdoc />
		public override string ToString()
			=> IsValid
				   ? Type switch {
					   AddressType.RowAbsolute    => $"{ColumnName}${Row}",
					   AddressType.ColumnAbsolute => $"${ColumnName}{Row}",
					   AddressType.Absolute       => $"${ColumnName}${Row}",
					   _                          => $"{ColumnName}{Row}",
				   }
				   : "Invalid";

		public void Deconstruct(out uint row, out uint column, out bool rowAbsolute, out bool columnAbsolute)
		{
			row = Row;
			column = Column;
			rowAbsolute = IsRowAbsolute;
			columnAbsolute = IsColumnAbsolute;
		}

		public void Deconstruct(out uint rowIndex, out uint columnIndex)
		{
			rowIndex = Row;
			columnIndex = Column;
		}

		#region Relational Operators

		/// <returns>Whether or not the left-side <see cref="CellAddress"/> is above and to the left of the right-side one</returns>
		public static bool operator <(CellAddress left, CellAddress right) => left.Row < right.Row && left.Column < right.Column;

		/// <returns>Whether or not the left-side <see cref="CellAddress"/> is above and to the left of, or equal to, the right-side one</returns>
		public static bool operator <=(CellAddress left, CellAddress right) => left.Row <= right.Row && left.Column <= right.Column;

		/// <returns>Whether or not the left-side <see cref="CellAddress"/> is below and to the right of the right-side one</returns>
		public static bool operator >(CellAddress left, CellAddress right) => left.Row > right.Row && left.Column > right.Column;

		/// <returns>Whether or not the left-side <see cref="CellAddress"/> is below and to the right of, or equal to, the right-side one</returns>
		public static bool operator >=(CellAddress left, CellAddress right) => left.Row >= right.Row && left.Column >= right.Column;

		#endregion

		#region Parsing/Conversion

		public static bool TryParse(string cellAddressText, out CellAddress cellAddress)
			=> CellReferenceParser.TryParseCellAddress(cellAddressText, out cellAddress);

		public static CellAddress Parse(string cellAddressText)
		{
			if (!TryParse(cellAddressText, out var cellAddress))
				throw new FormatException("The input string was not a valid cell reference");

			return cellAddress;
		}

		#endregion
	}
}