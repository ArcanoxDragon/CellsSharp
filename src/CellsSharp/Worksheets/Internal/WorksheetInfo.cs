using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellsSharp.Worksheets.Internal
{
	sealed class WorksheetInfo : IWorksheetInfo
	{
		public WorksheetInfo(Sheet sheet)
		{
			if (sheet.SheetId is null || sheet.Name?.Value is null)
				throw new ArgumentException("The provided Sheet is not valid", nameof(sheet));

			Index = sheet.SheetId.Value;
			Name = sheet.Name.Value;
		}

		/// <inheritdoc />
		public uint Index { get; }

		/// <inheritdoc />
		public string Name { get; }
	}
}