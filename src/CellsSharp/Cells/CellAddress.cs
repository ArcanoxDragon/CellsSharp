using System;
using CellsSharp.Cells.Internal;
using JetBrains.Annotations;

namespace CellsSharp.Cells
{
	[PublicAPI]
	public readonly struct CellAddress : ICellReference, IComparable<CellAddress>, IComparable
	{
		#region AddressType Conversion

		public static AddressType GetAddressType(bool rowAbsolute, bool columnAbsolute)
			=> rowAbsolute
				   ? columnAbsolute
						 ? AddressType.Absolute
						 : AddressType.RowAbsolute
				   : columnAbsolute
					   ? AddressType.ColumnAbsolute
					   : AddressType.Relative;

		#endregion

		public static readonly CellAddress None;

		public static readonly CellAddress WorksheetTopLeft     = new(1, 1);
		public static readonly CellAddress WorksheetTopRight    = new(1, Constants.ExcelMaxColumns);
		public static readonly CellAddress WorksheetBottomLeft  = new(Constants.ExcelMaxRows, 1);
		public static readonly CellAddress WorksheetBottomRight = new(Constants.ExcelMaxRows, Constants.ExcelMaxColumns);

		public CellAddress(uint row, uint column)
		{
			if (row is 0 or > Constants.ExcelMaxRows)
				throw new ArgumentOutOfRangeException(nameof(row));
			if (column is 0 or > Constants.ExcelMaxColumns)
				throw new ArgumentOutOfRangeException(nameof(column));

			Row = row;
			Column = column;
		}

		public uint Row    { get; }
		public uint Column { get; }

		public string ColumnName => IsValid ? CellReferenceParser.ColumnIndexToName(Column) : "Invalid";

		#region ICellReference

		public bool IsEmpty      => Row == 0 || Column == 0;
		public bool IsValid      => !IsEmpty;
		public bool IsSingleCell => true;

		public CellAddress TopLeft     => this;
		public CellAddress TopRight    => this;
		public CellAddress BottomLeft  => this;
		public CellAddress BottomRight => this;

		public CellReference ToCellReference() => new(this);
		public bool Contains(CellAddress cellAddress) => Equals(cellAddress);
		public bool IntersectsWith(ICellReference cellReference) => cellReference.Contains(this);
		public bool FullyContains(ICellReference cellReference) => Equals(cellReference);

		#endregion

		#region IComparable<CellAddress>

		public int CompareTo(CellAddress other)
		{
			var rowComparison = Row.CompareTo(other.Row);

			if (rowComparison != 0)
				return rowComparison;

			return Column.CompareTo(other.Column);
		}

		public int CompareTo(object? obj) => obj switch {
			CellAddress other => CompareTo(other),
			_                 => 1,
		};

		#endregion

		#region IEquatable<ICellReference>

		public bool Equals(ICellReference? other) => other switch {
			{ IsSingleCell: true } => Equals(other.TopLeft),
			_                      => false,
		};

		#endregion

		#region Equality

		public override bool Equals(object? obj) => obj switch {
			CellAddress other => Equals(other),
			_                 => false,
		};

		public bool Equals(CellAddress other)
			=> Row == other.Row && Column == other.Column;

		public override int GetHashCode()
		{
			unchecked
			{
				return ( (int) Row * 397 ) ^ (int) Column;
			}
		}

		public static bool operator ==(CellAddress left, CellAddress right) => left.Equals(right);
		public static bool operator !=(CellAddress left, CellAddress right) => !left.Equals(right);

		#endregion

		#region Relational Operators

		/// <returns>Whether or not the left-side <see cref="CellAddress"/> is above and to the left of the right-side one</returns>
		public static bool operator <(CellAddress left, CellAddress right) => left.Row < right.Row && left.Column < right.Column;

		/// <returns>Whether or not the left-side <see cref="CellAddress"/> is above and to the left of, or equal to, the right-side one</returns>
		public static bool operator <=(CellAddress left, CellAddress right) => left.Row <= right.Row && left.Column <= right.Column;

		/// <returns>Whether or not the left-side <see cref="CellAddress"/> is below and to the right of the right-side one</returns>
		public static bool operator >(CellAddress left, CellAddress right) => left.Row > right.Row && left.Column > right.Column;

		/// <returns>Whether or not the left-side <see cref="CellAddress"/> is below and to the right of, or equal to, the right-side one</returns>
		public static bool operator >=(CellAddress left, CellAddress right) => left.Row >= right.Row && left.Column >= right.Column;

		/// <returns>
		/// The cartesian distance between the top-left corners of the cells at
		/// <paramref name="left"/> and <paramref name="right"/>.
		/// </returns>
		public static double operator -(CellAddress left, CellAddress right)
		{
			double xDist = Math.Abs(right.Column - left.Column);
			double yDist = Math.Abs(right.Row - left.Row);

			return Math.Sqrt(xDist * xDist + yDist * yDist);
		}

		#endregion

		#region Parsing/Conversion

		public override string ToString() => ToString(AddressType.Relative);

		public string ToString(AddressType addressType)
			=> IsValid
				   ? addressType switch {
					   AddressType.RowAbsolute    => $"{ColumnName}${Row}",
					   AddressType.ColumnAbsolute => $"${ColumnName}{Row}",
					   AddressType.Absolute       => $"${ColumnName}${Row}",
					   _                          => $"{ColumnName}{Row}",
				   }
				   : "Invalid";

		public void Deconstruct(out uint rowIndex, out uint columnIndex)
		{
			rowIndex = Row;
			columnIndex = Column;
		}

		public static bool TryParse(string cellAddressText, out CellAddress cellAddress)
			=> CellReferenceParser.TryParseCellAddress(cellAddressText, out cellAddress);

		public static bool TryParse(string cellAddressText, out CellAddress cellAddress, out AddressType addressType)
			=> CellReferenceParser.TryParseCellAddress(cellAddressText, out cellAddress, out addressType);

		public static CellAddress Parse(string cellAddressText)
		{
			if (!TryParse(cellAddressText, out var cellAddress))
				throw new FormatException("The input string was not a valid cell reference");

			return cellAddress;
		}

		public static CellAddress Parse(string cellAddressText, out AddressType addressType)
		{
			if (!TryParse(cellAddressText, out var cellAddress, out addressType))
				throw new FormatException("The input string was not a valid cell reference");

			return cellAddress;
		}

		#endregion
	}
}