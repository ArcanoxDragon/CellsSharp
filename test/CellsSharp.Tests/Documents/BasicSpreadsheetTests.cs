using CellsSharp.Cells;
using NUnit.Framework;

namespace CellsSharp.Tests.Documents
{
	[TestFixture]
	public class BasicSpreadsheetTests : SpreadsheetDocumentTestFixture
	{
		[Test, Order(1)]
		public void CreateSimpleSpreadsheet()
		{
			using var document = CreateDocument();
			var testSheet = document.Workbook.Sheets.AddNew("Test Sheet");

			testSheet["A1:D4"].CellText = "Fill Test";

			document.Save();
		}

		[Test, Order(2)]
		public void ValidateSimpleSpreadsheet()
		{
			using var document = OpenDocument();
			var testSheetInfo = document.Workbook.Sheets["Test Sheet"];
			var testSheet = document.Workbook.Sheets.Open(testSheetInfo);

			foreach (var address in (CellReference) "A1:D4")
				Assert.That(testSheet[address].CellText, Is.EqualTo("Fill Test"));
		}
	}
}