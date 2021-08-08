using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellsSharp.Extensions
{
	static class OpenXmlSpreadsheetExtensions
	{
		internal static Sheet? GetSheetForWorksheetPart(this WorkbookPart workbookPart, WorksheetPart worksheetPart)
		{
			var worksheetId = workbookPart.GetIdOfPart(worksheetPart);
			var sheets = workbookPart.Workbook.Sheets;

			return sheets?.Elements<Sheet>().SingleOrDefault(s => s.Id == worksheetId);
		}
	}
}