using System;
using System.Collections.Generic;
using System.Linq;
using CellsSharp.Cells;
using CellsSharp.Extensions;
using CellsSharp.Internal.DataHandlers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellsSharp.Worksheets.Internal
{
	sealed class MergedCellsCollection : PartElementHandler<WorksheetPart, MergeCells>
	{
		#region Fields

		private readonly List<CellRange>      mergedCellRanges = new();
		private readonly HashSet<CellAddress> mergedCells      = new();

		#endregion

		#region Properties

		public override bool PartHasData => this.mergedCellRanges.Any();

		#endregion

		#region Public Methods

		public void Merge(CellRange mergeRange)
		{
			if (mergeRange.IsSingleCell)
				throw new ArgumentException("Range must be larger tha one cell", nameof(mergeRange));

			// Unmerge any intersecting ranges, if any
			Unmerge(mergeRange);

			// Now merge this one
			MergeRange(mergeRange);
		}

		public void Unmerge(CellRange range)
		{
			var currentMergedRanges = this.mergedCellRanges.ToList();

			foreach (var mergedRange in currentMergedRanges)
			{
				if (mergedRange.IntersectsWith(range))
					UnmergeRange(mergedRange);
			}
		}

		/// <summary>
		/// Returns whether the cell at <paramref name="address"/> is part of a
		/// merged cell range.
		/// </summary>
		public bool IsMerged(CellAddress address)
			=> this.mergedCells.Contains(address);

		/// <summary>
		/// Returns whether the cell at <paramref name="address"/> is part of a
		/// merged cell range. If it is, the merged range is stored in <paramref name="mergedRange"/>.
		/// </summary>
		public bool IsMerged(CellAddress address, out CellRange mergedRange)
		{
			mergedRange = CellRange.Empty;

			if (!IsMerged(address))
				return false;

			foreach (var range in this.mergedCellRanges)
			{
				if (range.Contains(address))
				{
					mergedRange = range;
					break;
				}
			}

			return true;
		}

		/// <summary>
		/// Returns whether <paramref name="cellRange"/> represents a merged range of
		/// cells.
		/// </summary>
		/// <remarks>
		/// If <paramref name="cellRange"/> intersects with a currently merged range of
		/// cells, but is not that exact range, this method returns false.
		/// </remarks>
		public bool IsMerged(CellRange cellRange)
			=> this.mergedCellRanges.Any(r => r == cellRange);

		#endregion

		#region Non-public Methods

		protected override void WriteElementData(OpenXmlWriter writer)
		{
			var mergeCell = new MergeCell { Reference = string.Empty };

			foreach (var range in this.mergedCellRanges)
			{
				mergeCell.Reference.Value = range.ToString();

				writer.WriteElement(mergeCell);
			}
		}

		protected override void ReadElementData(OpenXmlReader reader)
		{
			this.mergedCellRanges.Clear();
			this.mergedCells.Clear();

			reader.VisitChildren<MergeCell>(() => {
				if (reader.ElementType != typeof(MergeCell))
					return;

				MergeCell mergeCell = (MergeCell) reader.LoadCurrentElement()!;

				if (mergeCell.Reference is not { Value: { } cellReference })
					return;
				if (!CellRange.TryParse(cellReference, out var range))
					return;

				MergeRange(range);
			});
		}

		#endregion

		#region Private Methods

		private void MergeRange(CellRange range)
		{
			this.mergedCellRanges.Add(range);

			foreach (var address in range)
				this.mergedCells.Add(address);
		}

		private void UnmergeRange(CellRange range)
		{
			foreach (var address in range)
				this.mergedCells.Remove(address);

			this.mergedCellRanges.Remove(range);
		}

		#endregion
	}
}