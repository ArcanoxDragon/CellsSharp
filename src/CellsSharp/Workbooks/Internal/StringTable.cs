using System.Collections.Generic;
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

		public string this[uint index] => GetValue(index);

		public uint EntryCount => (uint) this.entryTable.Count;

		public void Clear()
		{
			this.entryTable.Clear();
			this.indexLookup.Clear();
			this.nextIndex = 0;

			ChangeNotifier.NotifyOfChange<SharedStringTablePart>(this);
		}

		public string GetValue(uint index)
		{
			if (index >= EntryCount)
				return string.Empty;

			return this.entryTable[index];
		}

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

		public override bool HandlesRootElement => true;

		public override bool PartHasData => EntryCount > 0;

		protected override SharedStringTable CreateElement()
			=> new() { Count = EntryCount, UniqueCount = EntryCount };

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