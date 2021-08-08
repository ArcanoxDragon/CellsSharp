namespace CellsSharp.Documents.Internal
{
	interface ISpreadsheetDocumentImpl : ISpreadsheetDocument
	{
		/// <summary>
		/// Loads the document from the source with which it was opened.
		/// </summary>
		void Load();
	}
}