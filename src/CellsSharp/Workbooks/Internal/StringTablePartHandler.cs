using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Workbooks.Internal
{
	sealed class StringTablePartHandler : ChildPartHandler<WorkbookPart, SharedStringTablePart>
	{
		/// <inheritdoc />
		public StringTablePartHandler(
			IChangeNotifier changeNotifier,
			IList<IPartElementHandler<SharedStringTablePart>> partElementHandlers,
			IStringTable stringTable
		) : base(changeNotifier, partElementHandlers)
		{
			StringTable = stringTable;
		}

		private IStringTable StringTable { get; }

		/// <inheritdoc />
		protected override bool PartHasData => StringTable.EntryCount > 0;

		/// <inheritdoc />
		protected override SharedStringTablePart CreateChildPart(WorkbookPart parentPart)
			=> parentPart.AddNewPart<SharedStringTablePart>();
	}
}