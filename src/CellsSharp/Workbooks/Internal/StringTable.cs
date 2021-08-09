using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using CellsSharp.Extensions;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellsSharp.Workbooks.Internal
{
	sealed class StringTable : PartElementHandler<SharedStringTablePart, SharedStringTable>, IStringTable
	{
		private readonly SortedList<uint, string> entryTable  = new();
		private readonly Dictionary<string, uint> indexLookup = new();

		private uint nextIndex;

		public StringTable(IChangeNotifier changeNotifier)
		{
			ChangeNotifier = changeNotifier;
		}

		private IChangeNotifier ChangeNotifier { get; }

		/// <inheritdoc />
		public string this[uint index]
			=> TryGetValue(index, out var value)
				   ? value
				   : throw new IndexOutOfRangeException("The provided index does not exist in the string table");

		/// <inheritdoc />
		public uint EntryCount => (uint) this.entryTable.Count;

		/// <inheritdoc />
		public void Clear()
		{
			this.entryTable.Clear();
			this.indexLookup.Clear();
			this.nextIndex = 0;

			ChangeNotifier.NotifyOfChange<SharedStringTablePart>(this);
		}

		/// <inheritdoc />
		public bool TryGetValue(uint index, [MaybeNullWhen(false)] out string value)
			=> this.entryTable.TryGetValue(index, out value);

		/// <inheritdoc />
		public uint GetOrInsertValue(string value)
		{
			if (this.indexLookup.TryGetValue(value, out var index))
				return index;

			index = this.nextIndex++;
			this.entryTable.Add(index, value);
			this.indexLookup.Add(value, index);

			ChangeNotifier.NotifyOfChange<SharedStringTablePart>(this);

			return index;
		}

		#region PartElementHandler

		/// <inheritdoc />
		public override bool HandlesRootElement => true;

		/// <inheritdoc />
		protected override SharedStringTable CreateElement()
			=> new() { Count = EntryCount, UniqueCount = EntryCount };

		/// <inheritdoc />
		protected override void WriteElementData(OpenXmlWriter writer)
		{
			Text entryText = new();
			SharedStringItem entryItem = new();

			foreach (var (_, entry) in this.entryTable)
			{
				entryText.Text = entry;

				writer.WriteElement(entryItem, () => {
					writer.WriteElement(entryText);
				});
			}
		}

		/// <inheritdoc />
		protected override void ReadElementData(OpenXmlReader reader)
		{
			Clear();

			reader.VisitChildren<SharedStringItem>(() => {
				uint thisIndex = this.nextIndex++;

				// Determine if the current element can actually be read as a SharedStringItem with non-empty text
				if (reader.LoadCurrentElement() is not SharedStringItem { Text: { Text: { } entry } })
					return;

				this.entryTable.Add(thisIndex, entry);
				this.indexLookup.Add(entry, thisIndex);
			});
		}

		#endregion
	}
}