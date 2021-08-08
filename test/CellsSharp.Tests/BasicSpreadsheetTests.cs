using System.IO;
using System.IO.Packaging;
using NUnit.Framework;

namespace CellsSharp.Tests
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

			document.Workbook.Strings.GetOrInsertValue("Test string 1");
			document.Workbook.Strings.GetOrInsertValue("Test string 1");
			document.Workbook.Strings.GetOrInsertValue("Test string 1");
			document.Workbook.Strings.GetOrInsertValue("Test string 2");
			document.Workbook.Strings.GetOrInsertValue("Test string 3");

			// var testSheet = document.Workbook.Sheets.AddNew("Test Sheet");

			document.Save();
		}

		[Test, Order(2)]
		public void OpenSimpleSpreadsheet()
		{
			using var package = Package.Open(TestDocumentFilename, FileMode.Open, FileAccess.Read);
			using var document = SpreadsheetDocument.Open(package);

			// TODO: Test assertions
		}
	}
}