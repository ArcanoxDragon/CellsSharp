using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Workbooks.Internal
{
	sealed class WorkbookPartHandler : RootPartHandler<WorkbookPart>
	{
		/// <inheritdoc />
		public WorkbookPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<WorkbookPart>> childPartHandlers,
			WorkbookPart workbookPart
		) : base(changeNotifier, childPartHandlers)
		{
			WorkbookPart = workbookPart;
		}

		private WorkbookPart WorkbookPart { get; }

		/// <inheritdoc />
		protected override WorkbookPart HandledPart => WorkbookPart;
	}
}