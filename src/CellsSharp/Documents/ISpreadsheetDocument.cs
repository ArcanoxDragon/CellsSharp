using System;
using CellsSharp.Workbooks;
using JetBrains.Annotations;

namespace CellsSharp.Documents
{
	[PublicAPI]
	public interface ISpreadsheetDocument : IDisposable
	{
		/// <summary>
		/// Gets the document's workbook part
		/// </summary>
		IWorkbook Workbook { get; }

		/// <summary>
		/// Saves the document and all related parts.
		/// </summary>
		void Save();

		/// <summary>
		/// Saves the document and all related parts to the provided file path.
		/// </summary>
		void SaveAs(string path);
	}
}