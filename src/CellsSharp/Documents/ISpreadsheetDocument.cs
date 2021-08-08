using System;
using CellsSharp.Workbooks;
using JetBrains.Annotations;

namespace CellsSharp.Documents
{
	[PublicAPI]
	// TODO: Documentation
	public interface ISpreadsheetDocument : IDisposable
	{
		IWorkbook Workbook { get; }

		void Save();
		void SaveAs(string path);
	}
}