using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Workbooks.Internal
{
	sealed class StringTablePartHandler : DefaultChildPartHandler<WorkbookPart, SharedStringTablePart>
	{
		/// <inheritdoc />
		public StringTablePartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<SharedStringTablePart>> childPartHandlers,
			IList<IPartElementHandler<SharedStringTablePart>> partElementHandlers,
			IStringTable stringTable
		) : base(changeNotifier, childPartHandlers, partElementHandlers)
		{
			StringTable = stringTable;
		}

		private IStringTable StringTable { get; }

		/// <inheritdoc />
		protected override bool PartHasData => StringTable.EntryCount > 0;
	}
}