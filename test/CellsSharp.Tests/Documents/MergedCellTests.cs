using CellsSharp.Cells;
using NUnit.Framework;

namespace CellsSharp.Tests.Documents
{
	[TestFixture]
	public class MergedCellTests : SpreadsheetDocumentTestFixture
	{
		[Test, Order(1)]
		public void CreateMergedCellTestDocument()
		{
			using var document = CreateDocument();
			var testSheet = document.Workbook.Sheets.AddNew("Test Sheet");

			testSheet.ForRange("B2:C3", r => {
				r.Merge();

				r.CellText = "Merged (set after by top-left)";
			});

			testSheet["B5:C6"].Merge();
			testSheet["C6"].CellText = "Merged (set after by bottom-right)";

			testSheet["B8"].CellText = "Merged (set before by top-left)";
			testSheet["B8:C9"].Merge();

			testSheet["C12"].CellText = "Merged (set before by bottom-right)";
			testSheet["B11:C12"].Merge();
		}

		[Test, Order(2)]
		public void ValidateMergedCellTestDocument()
		{
			using var document = OpenDocument();
			var testSheetInfo = document.Workbook.Sheets["Test Sheet"];
			var testSheet = document.Workbook.Sheets.Open(testSheetInfo);

			Assert.That(testSheet["B2:C3"].IsMerged);

			foreach (var address in (CellRange) "B2:C3")
				Assert.That(testSheet[address].CellText, Is.EqualTo("Merged (set after by top-left)"));

			Assert.That(testSheet["B5:C6"].IsMerged);

			foreach (var address in (CellRange) "B5:C6")
				Assert.That(testSheet[address].CellText, Is.EqualTo("Merged (set after by bottom-right)"));

			Assert.That(testSheet["B8:C9"].IsMerged);

			foreach (var address in (CellRange) "B8:C9")
				Assert.That(testSheet[address].CellText, Is.EqualTo("Merged (set before by top-left)"));

			Assert.That(testSheet["B11:C12"].IsMerged);

			foreach (var address in (CellRange) "B11:C12")
				// C12's text was set, and then B11:C12 were merged, so the
				// merged cell's contents were cleared. The value returned
				// should be an empty string
				Assert.That(testSheet[address].CellText, Is.Empty);
		}
	}
}