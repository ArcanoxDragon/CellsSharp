using JetBrains.Annotations;

namespace CellsSharp.Workbooks
{
	[PublicAPI]
	// TODO: Documentation
	public interface IWorkbook
	{
		public ISheetCollection Sheets { get; }
	}
}