namespace CellsSharp.Workbooks.Internal
{
	sealed class WorkbookManager : IWorkbook
	{
		public WorkbookManager(ISheetCollection sheetCollection)
		{
			Sheets = sheetCollection;
		}

		/// <inheritdoc />
		public ISheetCollection Sheets { get; }
	}
}