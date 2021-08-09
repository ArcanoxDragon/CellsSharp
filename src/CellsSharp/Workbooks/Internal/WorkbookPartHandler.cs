using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Workbooks.Internal
{
	sealed class WorkbookPartHandler : DefaultRootPartHandler<WorkbookPart>, IDocumentSaveLoadHandler
	{
		/// <inheritdoc />
		public WorkbookPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<WorkbookPart>> childPartHandlers,
			IList<IPartElementHandler<WorkbookPart>> partElementHandlers,
			WorkbookPart workbookPart
		) : base(changeNotifier, childPartHandlers, partElementHandlers, workbookPart) { }
	}
}