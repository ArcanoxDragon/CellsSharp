using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MsWorkbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;

namespace CellsSharp.Workbooks.Internal
{
	sealed class WorkbookPartHandler : DefaultRootPartHandler<WorkbookPart>, IDocumentSaveLoadHandler
	{
		public WorkbookPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<WorkbookPart>> childPartHandlers,
			IList<IPartElementHandler<WorkbookPart>> partElementHandlers,
			WorkbookPart workbookPart
		) : base(changeNotifier, childPartHandlers, partElementHandlers, workbookPart) { }

		protected override OpenXmlElement CreateRootElement()
			=> new MsWorkbook();
	}
}