using System.Collections.Generic;
using CellsSharp.Worksheets;
using JetBrains.Annotations;

namespace CellsSharp.Workbooks
{
	[PublicAPI]
	// TODO: Documentation
	public interface ISheetCollection : IEnumerable<IWorksheetInfo>
	{
		IWorksheetInfo this[string name] { get; }
		IWorksheetInfo this[uint index] { get; }

		IWorksheet AddNew();
		IWorksheet AddNew(string name);

		/// <summary>
		/// Opens an instance of <see cref="IWorksheet"/> for the worksheet the provided
		/// <paramref name="worksheetInfo"/> represents
		/// </summary>
		IWorksheet Open(IWorksheetInfo worksheetInfo);

		void Remove(string name);
		void Remove(uint index);
		void Remove(IWorksheetInfo worksheetInfo);
	}
}