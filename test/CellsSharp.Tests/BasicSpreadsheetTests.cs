using System.IO;
using System.IO.Packaging;
using NUnit.Framework;

namespace CellsSharp.Tests
{
	[TestFixture]
	public class BasicSpreadsheetTests
	{
		private const string TestDocumentFilename = "TestDocument.xlsx";

		[Test]
		public void CreateSimpleSpreadsheet()
		{
			using var package = Package.Open(TestDocumentFilename, FileMode.Create, FileAccess.ReadWrite);
			using var document = SpreadsheetDocument.Create(package);

			var testSheet = document.Workbook.Sheets.AddNew("Test Sheet");

			document.Save();
		}
	}
}