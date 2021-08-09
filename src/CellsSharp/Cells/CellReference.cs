using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using CellsSharp.Cells.Internal;
using CellsSharp.Extensions;
using JetBrains.Annotations;

namespace CellsSharp.Cells
{
	[PublicAPI]
	public sealed class CellReference : ICellReference
	{
		public static readonly CellReference Empty = new(Enumerable.Empty<CellRange>());

		private readonly List<CellRange> ranges;

		public CellReference(IEnumerable<CellRange> ranges)
		{
			this.ranges = new List<CellRange>(ranges.Where(r => r.IsValid).Distinct());
		}

		public CellReference(CellRange range)
		{
			this.ranges = new List<CellRange> { range };
		}

		public IEnumerable<CellRange> Ranges => this.ranges;

		#region ICellReference

		/// <inheritdoc />
		public bool IsEmpty => this.ranges.Count == 0;

		/// <inheritdoc />
		public bool IsValid => true;

		/// <inheritdoc />
		public bool IsSingleCell => this.ranges.Count == 1 && this.ranges.First().IsSingleCell;

		/// <inheritdoc />
		/// <remarks>
		/// The top-left cell will be determined by sorting the Start address of all ranges
		/// represented by this <see cref="CellReference"/> by their distance to the top-left
		/// corner of the worksheet. In the event of a tie, the highest row wins.
		/// </remarks>
		public CellAddress TopLeft
			=> IsEmpty
				   ? CellAddress.None
				   : this.ranges
						 .OrderBy(r => Math.Sqrt(r.Start.Row * r.Start.Row + r.Start.Column * r.Start.Column))
						 .ThenBy(r => r.Start.Row)
						 .First()
						 .Start;

		/// <inheritdoc />
		public bool Contains(CellAddress cellAddress) => this.ranges.Any(r => r.Contains(cellAddress));

		/// <inheritdoc />
		public bool IntersectsWith(ICellReference cellReference) => this.ranges.Any(r => r.IntersectsWith(cellReference));

		/// <inheritdoc />
		public bool FullyContains(ICellReference cellReference) => this.ranges.Any(r => r.FullyContains(cellReference));

		#endregion

		#region IEnumerable<CellAddress>

		/// <inheritdoc />
		public IEnumerator<CellAddress> GetEnumerator() => this.ranges.Flatten().GetEnumerator();

		/// <inheritdoc />
		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

		#endregion

		#region IEquatable<ICellReference>

		/// <inheritdoc />
		public bool Equals(ICellReference? other)
		{
			if (ReferenceEquals(this, other))
				return true;
			if (ReferenceEquals(null, other))
				return false;

			return other switch {
				CellAddress address     => IsSingleCell && address == this.ranges.First().Start,
				CellRange range         => this.ranges.Count == 1 && range.Equals(this.ranges.First()),
				CellReference reference => Equals(reference),
				_                       => false,
			};
		}

		#endregion

		#region Equality

		private bool Equals(CellReference other)
		{
			return this.ranges.SequenceEqual(other.ranges);
		}

		/// <inheritdoc />
		public override bool Equals(object? obj)
		{
			return ReferenceEquals(this, obj) || obj is ICellReference other && Equals(other);
		}

		/// <inheritdoc />
		public override int GetHashCode()
		{
#if NET5_0_OR_GREATER
			var hashCode = new HashCode();

			foreach (var range in this.ranges)
				hashCode.Add(range);

			return hashCode.ToHashCode();
#else
			unchecked
			{
				return this.ranges.Aggregate(0, (hashCode, range) => ( hashCode * 397 ) ^ range.GetHashCode());
			}
#endif
		}

		public static bool operator ==(CellReference? left, CellReference? right)
		{
			return Equals(left, right);
		}

		public static bool operator !=(CellReference? left, CellReference? right)
		{
			return !Equals(left, right);
		}

		#endregion

		#region Parsing/Conversion

		public static implicit operator CellReference(CellAddress singleAddress) => new(singleAddress);

		public static implicit operator CellReference(CellRange range) => new(range);

		public static implicit operator CellReference(string value) => Parse(value) switch {
			// Have to do this weirdly because of how implicit operators are resolved
			CellAddress address     => address,
			CellRange range         => range,
			CellReference reference => reference,
			_                       => throw new InvalidCastException(),
		};

		public static bool TryParse(string cellReferenceText, out ICellReference cellReference)
		{
			cellReference = Parse(cellReferenceText);

			return !cellReference.IsEmpty;
		}

		public static ICellReference Parse(string cellReferenceText)
			=> CellReferenceParser.ParseCellReference(cellReferenceText);

		#endregion
	}
}