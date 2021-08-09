using JetBrains.Annotations;

namespace CellsSharp.Worksheets
{
	[PublicAPI]
	public interface ISheetView
	{
		/// <summary>
		/// Gets or sets the text of the cells in this <see cref="ISheetView"/>.
		/// </summary>
		string CellText { get; set; }

		/// <summary>
		/// Gets or sets the numeric value of the cells in this <see cref="ISheetView"/>.
		/// </summary>
		double CellValue { get; set; }

		/// <summary>
		/// Clears the values of all cells in this <see cref="ISheetView"/>.
		/// </summary>
		void ClearValue();

		/// <summary>
		/// Clears the formatting of all cells in this <see cref="ISheetView"/>.
		/// </summary>
		void ClearFormatting();

		/// <summary>
		/// Clears the values and formatting of all cells in this <see cref="ISheetView"/>.
		/// </summary>
		void ClearAll();
	}
}