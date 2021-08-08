using DocumentFormat.OpenXml.Spreadsheet;

namespace CellsSharp.Worksheets.Internal
{
	sealed class WorksheetInfo : IWorksheetInfo
	{
		public WorksheetInfo(Sheet sheet)
		{

		}

		/// <inheritdoc />
		public uint Index { get; }

		/// <inheritdoc />
		public string Name { get; }
	}
}