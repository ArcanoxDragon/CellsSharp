using System.IO;
using System.IO.Packaging;
using CellsSharp.Cells;
using NUnit.Framework;

namespace CellsSharp.Tests.Documents
{
	[TestFixture]
	public class BasicSpreadsheetTests
	{
		private const string TestDocumentFilename = "TestDocument.xlsx";

		[Test, Order(1)]
		public void CreateSimpleSpreadsheet()
		{
			using var package = Package.Open(TestDocumentFilename, FileMode.Create, FileAccess.ReadWrite);
			using var document = SpreadsheetDocument.Create(package);
			var testSheet = document.Workbook.Sheets.AddNew("Test Sheet");

			testSheet["A1:D4"].CellText = "Fill Test";

			document.Save();
		}

		[Test, Order(2)]
		public void OpenSimpleSpreadsheet()
		{
			using var package = Package.Open(TestDocumentFilename, FileMode.Open, FileAccess.Read);
			using var document = SpreadsheetDocument.Open(package);
			var testSheetInfo = document.Workbook.Sheets["Test Sheet"];
			var testSheet = document.Workbook.Sheets.Open(testSheetInfo);

			foreach (var address in (CellReference) "A1:D4")
				Assert.That(testSheet[address].CellText, Is.EqualTo("Fill Test"));
		}
	}
}