using CellsSharp.Documents;

namespace CellsSharp.Tests.Documents
{
	public class SpreadsheetDocumentTestFixture
	{
		protected virtual string TestDocumentFilename => "TestDocument.xlsx";

		protected ISpreadsheetDocument CreateDocument()
			=> SpreadsheetDocument.Create(TestDocumentFilename);

		protected ISpreadsheetDocument OpenDocument(bool editable = false)
			=> SpreadsheetDocument.Open(TestDocumentFilename, editable);
	}
}