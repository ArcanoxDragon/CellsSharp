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

		/// <inheritdoc />
		public ISheetCollection Sheets { get; }

		/// <inheritdoc />
		public IStringTable Strings { get; }

		#endregion
	}
}