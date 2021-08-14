namespace CellsSharp.Workbooks.Internal
{
	sealed class Workbook : IWorkbook
	{
		public Workbook(ISheetCollection sheetCollection, IStringTable stringTable)
		{
			Sheets = sheetCollection;
			Strings = stringTable;
		}

		#region IWorkbook

		public ISheetCollection Sheets  { get; }
		public IStringTable     Strings { get; }

		#endregion
	}
}