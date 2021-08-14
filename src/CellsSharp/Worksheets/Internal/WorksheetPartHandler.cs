using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MsWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace CellsSharp.Worksheets.Internal
{
	sealed class WorksheetPartHandler : DefaultRootPartHandler<WorksheetPart>, IWorksheetSaveLoadHandler
	{
		public WorksheetPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<WorksheetPart>> childPartHandlers,
			IList<IPartElementHandler<WorksheetPart>> partElementHandlers,
			WorksheetPart worksheetPart
		) : base(changeNotifier, childPartHandlers, partElementHandlers, worksheetPart) { }

		protected override OpenXmlElement CreateRootElement() => new MsWorksheet();
	}
}