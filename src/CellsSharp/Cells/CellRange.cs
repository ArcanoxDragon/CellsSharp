using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using CellsSharp.Cells.Internal;
using JetBrains.Annotations;

namespace CellsSharp.Cells
{
	[PublicAPI]
	public readonly struct CellRange : ICellReference, IEnumerable<CellAddress>
	{
		public static readonly CellRange Empty = new(CellAddress.None, CellAddress.None);

		public CellRange(uint startRow, uint startColumn, uint endRow, uint endColumn)
			: this(new CellAddress(startRow, startColumn), new CellAddress(endRow, endColumn)) { }

		public CellRange(CellAddress start, CellAddress end)
		{
			if (!( start <= end ))
				throw new ArgumentOutOfRangeException(nameof(start), $"{nameof(start)} must not be below or to the right of {nameof(end)}");

			Start = start;
			End = end;
		}

		public CellRange(CellAddress address) : this(address, address) { }

		public CellAddress Start { get; }
		public CellAddress End   { get; }

		public uint Top    => Start.Row;
		public uint Bottom => End.Row;
		public uint Left   => Start.Column;
		public uint Right  => End.Column;

		public bool IntersectsWith(CellRange other)
			=> Start <= other.End && other.Start <= End;

		public bool FullyContains(CellRange other)
			=> other.Start >= Start && other.End <= End;

		public CellRange GetIntersection(CellRange other)
		{
			var intersectionStart = new CellAddress(Math.Max(Top, other.Top),
													Math.Max(Left, other.Left));
			var intersectionEnd = new CellAddress(Math.Min(Bottom, other.Bottom),
												  Math.Min(Right, other.Right));

			return new CellRange(intersectionStart, intersectionEnd);
		}

		#region ICellReference

		public bool IsEmpty      => Start.IsEmpty || End.IsEmpty;
		public bool IsValid      => Start.IsValid && End.IsValid;
		public bool IsSingleCell => Start == End;

		public CellAddress TopLeft     => Start;
		public CellAddress TopRight    => new(Start.Row, End.Column);
		public CellAddress BottomLeft  => new(End.Row, Start.Column);
		public CellAddress BottomRight => End;

		public CellReference ToCellReference() => new(this);

		public bool Contains(CellAddress cellAddress) => cellAddress >= Start && cellAddress <= End;

		public bool IntersectsWith(ICellReference cellReference) => cellReference switch {
			CellAddress address     => Contains(address),
			CellRange range         => IntersectsWith(range),
			CellReference reference => reference.IntersectsWith(this),
			_                       => false,
		};

		public bool FullyContains(ICellReference cellReference) => cellReference switch {
			CellAddress address     => Contains(address),
			CellRange range         => FullyContains(range),
			CellReference reference => reference.Ranges.All(FullyContains),
			_                       => false,
		};

		#endregion

		#region IEnumerable<CellAddress>

		private IEnumerable<CellAddress> EnumerateAddresses()
		{
			for (uint row = Start.Row; row <= End.Row; row++)
			for (uint column = Start.Column; column <= End.Column; column++)
				yield return new CellAddress(row, column);
		}

		public IEnumerator<CellAddress> GetEnumerator() => EnumerateAddresses().GetEnumerator();

		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

		#endregion

		#region IEquatable<ICellReference>

		public bool Equals(ICellReference? other) => other switch {
			CellAddress address     => address == Start && address == End,
			CellRange range         => Equals(range),
			CellReference reference => reference.Ranges.All(Equals),
			_                       => false,
		};

		#endregion

		#region Equality

		public bool Equals(CellRange other)
			=> Start.Equals(other.Start) && End.Equals(other.End);

		public override bool Equals(object? obj)
			=> obj is CellRange other && Equals(other);

		public override int GetHashCode()
		{
			unchecked
			{
				return ( Start.GetHashCode() * 397 ) ^ End.GetHashCode();
			}
		}

		public static bool operator ==(CellRange left, CellRange right) => left.Equals(right);

		public static bool operator !=(CellRange left, CellRange right) => !left.Equals(right);

		#endregion

		#region Parsing/Conversion

		public override string ToString()
			=> IsValid
				   ? IsSingleCell
						 ? Start.ToString()
						 : $"{Start}:{End}"
				   : "Invalid";

		public static implicit operator CellRange(CellAddress singleAddress) => new(singleAddress);

		public static implicit operator CellRange((CellAddress Start, CellAddress End) addressPair)
			=> new(addressPair.Start, addressPair.End);

		public static explicit operator CellRange(string cellRangeStr)
			=> TryParse(cellRangeStr, out var cellRange) ? cellRange : Empty;

		public static bool TryParse(string cellRangeText, out CellRange cellRange)
			=> CellReferenceParser.TryParseCellRange(cellRangeText, out cellRange);

		public static CellRange Parse(string cellRangeText)
		{
			if (!TryParse(cellRangeText, out var cellRange))
				throw new FormatException("The input string was not a valid cell reference");

			return cellRange;
		}

		#endregion
	}
}