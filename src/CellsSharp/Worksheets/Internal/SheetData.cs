using System;
using System.Collections.Generic;
using System.Linq;
using CellsSharp.Cells;
using CellsSharp.Extensions;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MsSheetData = DocumentFormat.OpenXml.Spreadsheet.SheetData;

namespace CellsSharp.Worksheets.Internal
{
	sealed class SheetData : PartElementHandler<WorksheetPart, MsSheetData>
	{
		private readonly SortedList<uint, double> rowHeightMap = new();

		// It's way more memory efficient to store potential cell metadata as primitive types in separate dictionaries.
		// If we boxed stuff like "uint" or "double" into "object", a 4-byte number becomes a 24-byte reference type

		private readonly HashSet<CellAddress> cellsWithData = new();

		private readonly SortedList<CellAddress, uint>       stringIndexMap  = new();
		private readonly SortedList<CellAddress, double>     numericValueMap = new();
		private readonly SortedList<CellAddress, CellValues> dataTypeMap     = new();
		private readonly SortedList<CellAddress, uint>       styleIndexMap   = new();

		private uint minKnownRow, minKnownColumn;
		private uint maxKnownRow, maxKnownColumn;

		public SheetData(IChangeNotifier changeNotifier)
		{
			ChangeNotifier = changeNotifier;
		}

		public CellAddress MinKnownAddress => new(this.minKnownRow, this.minKnownColumn);
		public CellAddress MaxKnownAddress => new(this.maxKnownRow, this.maxKnownColumn);

		private IChangeNotifier ChangeNotifier { get; }

		public void ClearCell(CellAddress address)
		{
			this.stringIndexMap.Remove(address);
			this.numericValueMap.Remove(address);
			this.dataTypeMap.Remove(address);
			this.styleIndexMap.Remove(address);

			UpdateOccupied(address, false);
			MarkDirty();
		}

		#region Row Height

		public bool TryGetRowHeight(uint rowIndex, out double rowHeight)
			=> this.rowHeightMap.TryGetValue(rowIndex, out rowHeight);

		public void PutRowHeight(uint rowIndex, double rowHeight)
		{
			this.rowHeightMap[rowIndex] = rowHeight;

			MarkDirty();
		}

		public void ClearRowHeight(uint rowIndex)
		{
			this.rowHeightMap.Remove(rowIndex);

			MarkDirty();
		}

		#endregion

		#region Value

		// TODO: Other data types

		public bool TryGetStringIndex(CellAddress address, out uint stringIndex)
			=> this.stringIndexMap.TryGetValue(address, out stringIndex);

		public bool TryGetNumericValue(CellAddress address, out double numericValue)
			=> this.numericValueMap.TryGetValue(address, out numericValue);

		public void PutStringIndex(CellAddress address, uint stringIndex)
		{
			this.numericValueMap.Remove(address);
			this.stringIndexMap[address] = stringIndex;

			UpdateKnownBounds(address);
			UpdateOccupied(address, true);
			MarkDirty();
		}

		public void PutNumericValue(CellAddress address, double numericValue)
		{
			this.stringIndexMap.Remove(address);
			this.numericValueMap[address] = numericValue;

			UpdateKnownBounds(address);
			UpdateOccupied(address, true);
			MarkDirty();
		}

		public void ClearValue(CellAddress address)
		{
			this.stringIndexMap.Remove(address);
			this.numericValueMap.Remove(address);
			this.dataTypeMap.Remove(address);

			UpdateKnownBounds(address);
			CheckOccupied(address);
			MarkDirty();
		}

		#endregion

		#region Style Index

		public uint GetStyleIndex(CellAddress address)
		{
			if (!this.styleIndexMap.TryGetValue(address, out var styleIndex))
				styleIndex = 0u;

			return styleIndex;
		}

		public void SetStyleIndex(CellAddress address, uint styleIndex)
		{
			if (styleIndex is 0)
			{
				this.styleIndexMap.Remove(address);
				CheckOccupied(address);
			}
			else
			{
				this.styleIndexMap[address] = styleIndex;
				UpdateOccupied(address, true);
			}

			UpdateKnownBounds(address);
			MarkDirty();
		}

		public void ClearStyleIndex(CellAddress address)
			=> SetStyleIndex(address, 0);

		#endregion

		#region Private Methods

		private void MarkDirty() => ChangeNotifier.NotifyOfChange<WorksheetPart>(this);

		private void UpdateKnownBounds(CellAddress address)
		{
			if (this.minKnownRow == 0 || address.Row < this.minKnownRow)
				this.minKnownRow = address.Row;
			if (this.minKnownColumn == 0 || address.Column < this.minKnownColumn)
				this.minKnownColumn = address.Column;
			if (address.Row > this.maxKnownRow)
				this.maxKnownRow = address.Row;
			if (address.Column > this.maxKnownColumn)
				this.maxKnownColumn = address.Column;
		}

		private void UpdateOccupied(CellAddress address, bool occupied)
		{
			if (occupied)
				this.cellsWithData.Add(address);
			else
				this.cellsWithData.Remove(address);
		}

		private void CheckOccupied(CellAddress address)
		{
			var occupied = this.stringIndexMap.ContainsKey(address) ||
						   this.numericValueMap.ContainsKey(address) ||
						   this.dataTypeMap.ContainsKey(address) ||
						   this.styleIndexMap.ContainsKey(address);

			UpdateOccupied(address, occupied);
		}

		private void Clear()
		{
			this.cellsWithData.Clear();
			this.stringIndexMap.Clear();
			this.numericValueMap.Clear();
			this.dataTypeMap.Clear();
			this.styleIndexMap.Clear();

			MarkDirty();
		}

		#endregion

		#region PartElementHandler

		protected override void WriteElementData(OpenXmlWriter writer)
		{
			var sortedOccupiedAddresses = this.cellsWithData.ToArray();

			Array.Sort(sortedOccupiedAddresses); // Ridiculously fast

			var lastRowIndex = 0u;
			var startedRow = false;

			// Row element templates
			var rowIndexValue = new UInt32Value();
			var rowHeightValue = new DoubleValue();
			var customRowHeightValue = new BooleanValue(true);
			var rowTemplate = new Row { RowIndex = rowIndexValue };

			// Cell element templates
			var cellReferenceValue = new StringValue();
			var dataTypeValue = new EnumValue<CellValues>();
			var styleIndexValue = new UInt32Value();
			var cellValueTemplate = new CellValue();
			var cellTemplate = new Cell { CellReference = cellReferenceValue };
			bool cellHasValue;

			void UpdateRowTemplate(uint rowIndex)
			{
				rowIndexValue.Value = rowIndex;

				if (TryGetRowHeight(rowIndex, out var rowHeight))
				{
					rowHeightValue.Value = rowHeight;
					rowTemplate.Height = rowHeightValue;
					rowTemplate.CustomHeight = customRowHeightValue;
				}
				else
				{
					rowTemplate.Height = null;
					rowTemplate.CustomHeight = null;
				}
			}

			void WriteRowHeightsUntil(uint newRowIndex)
			{
				// Write any empty rows that have an explicit row height set

				for (var rowIndex = lastRowIndex + 1; rowIndex < newRowIndex; rowIndex++)
				{
					if (this.rowHeightMap.ContainsKey(rowIndex))
					{
						UpdateRowTemplate(rowIndex);
						writer.WriteElement(rowTemplate);
					}
				}
			}

			void StartNewRow(uint newRowIndex)
			{
				UpdateRowTemplate(newRowIndex);
				writer.WriteStartElement(rowTemplate);

				lastRowIndex = newRowIndex;
				startedRow = true;
			}

			void UpdateCellTemplate(CellAddress address)
			{
				cellReferenceValue.Value = address.ToString(AddressType.Relative);

				// Update value/data type
				if (TryGetStringIndex(address, out var stringIndex))
				{
					dataTypeValue.Value = CellValues.SharedString;
					cellValueTemplate.Text = stringIndex.ToString();
					cellTemplate.DataType = dataTypeValue;
					cellHasValue = true;
				}
				else if (TryGetNumericValue(address, out var numericValue))
				{
					dataTypeValue.Value = CellValues.Number;
					cellValueTemplate.Text = numericValue.ToString();
					cellTemplate.DataType = dataTypeValue;
					cellHasValue = true;
				}
				// TODO: Other data types
				else
				{
					cellTemplate.DataType = null;
					cellHasValue = false;
				}

				// Update style index
				var styleIndex = GetStyleIndex(address);

				if (styleIndex > 0)
				{
					styleIndexValue.Value = styleIndex;
					cellTemplate.StyleIndex = styleIndexValue;
				}
				else
				{
					cellTemplate.StyleIndex = null;
				}
			}

			// Iterate through all addresses inside the maximum bounds we know of
			// and skip any addresses that we've marked as not containing data
			foreach (var address in (MinKnownAddress, MaxKnownAddress))
			{
				if (!this.cellsWithData.Contains(address))
					continue;

				var (rowIndex, _) = address;

				if (rowIndex != lastRowIndex)
				{
					// Finish the current row (if any)
					if (startedRow)
						writer.WriteEndElement();

					if (rowIndex - lastRowIndex > 1)
						// Write any empty rows that have an explicit height set
						// (still need to maintain that even if the row has no cell data)
						WriteRowHeightsUntil(rowIndex);

					// Write the start element for this row
					StartNewRow(rowIndex);
				}

				UpdateCellTemplate(address);

				if (cellHasValue)
				{
					writer.WriteElement(cellTemplate, () => {
						writer.WriteElement(cellValueTemplate);
					});
				}
				else
				{
					writer.WriteElement(cellTemplate);
				}
			}

			// Finish the last row if one was started
			if (startedRow)
				writer.WriteEndElement();

			if (this.rowHeightMap.Any())
			{
				// See if there are any rows with explicit heights set after the bottom-most cell
				var maxRowWithHeight = this.rowHeightMap.Keys.Max();

				if (maxRowWithHeight > lastRowIndex)
				{
					// Write out the remaining rows with custom heights
					WriteRowHeightsUntil(maxRowWithHeight + 1);
				}
			}
		}

		protected override void ReadElementData(OpenXmlReader reader)
		{
			Clear();

			reader.VisitChildren<Row>(() => {
				var rowIndex = default(uint);
				var rowHeight = default(double?);

				foreach (var attribute in reader.Attributes)
				{
					if (attribute.Is<uint>("r", out var uintValue))
						rowIndex = uintValue;
					else if (attribute.Is<double>("ht", out var doubleValue))
						rowHeight = doubleValue;
				}

				if (rowIndex is 0)
					return;
				if (rowHeight is not null)
					this.rowHeightMap.Add(rowIndex, rowHeight.Value);

				reader.VisitChildren<Cell>(() => {
					var cell = (Cell) reader.LoadCurrentElement()!;

					if (cell.CellReference?.Value is null || !CellAddress.TryParse(cell.CellReference.Value, out var address))
						return;

					// Cell data type
					if (cell.DataType is not { Value: var dataType })
						dataType = CellValues.Number;

					// Cell value
					if (cell.CellValue is { Text: { } cellValue })
					{
						switch (dataType)
						{
							case CellValues.Number:
								if (double.TryParse(cellValue, out var doubleValue))
									PutNumericValue(address, doubleValue);
								break;
							case CellValues.SharedString:
								if (uint.TryParse(cellValue, out var uintValue))
									PutStringIndex(address, uintValue);
								break;
							// TODO: More DataTypes
						}
					}

					// Style index
					if (cell.StyleIndex is { Value: var styleIndex })
						SetStyleIndex(address, styleIndex);
				});
			});
		}

		#endregion
	}
}