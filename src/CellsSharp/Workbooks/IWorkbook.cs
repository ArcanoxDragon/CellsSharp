using JetBrains.Annotations;

namespace CellsSharp.Workbooks
{
	[PublicAPI]
	// TODO: Documentation
	public interface IWorkbook
	{
		ISheetCollection Sheets  { get; }
		IStringTable     Strings { get; }
	}
}