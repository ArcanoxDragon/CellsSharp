using JetBrains.Annotations;

namespace CellsSharp.Styles
{
	[PublicAPI]
	public interface IAlignmentProperties
	{
		/// <summary>
		/// Gets or sets the horizontal alignment of text in a cell or row of cells.
		/// </summary>
		HorizontalAlignment Horizontal { get; set; }

		/// <summary>
		/// Gets or sets the vertical alignment of text in a cell.
		/// </summary>
		VerticalAlignment Vertical { get; set; }

		/// <summary>
		/// Gets or sets the angle of rotation, in degrees, of the text in a cell.
		/// </summary>
		/// <remarks>
		/// Valid values range from -90 to 90. Positive values represent text rotated
		/// such that the trailing edge is higher than the leading edge, and negative
		/// values represent text rotated such that the trailing edge is lower than
		/// the leading edge.
		/// </remarks>
		int TextRotation { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not to wrap text in a cell
		/// when it is too long to fit inside the cell's width.
		/// </summary>
		bool WrapText { get; set; }

		/// <summary>
		/// Gets or sets the absolute indent level of a cell.
		/// </summary>
		/// <remarks>
		/// One indent level represents a width of 3 spaces in the cell's font.
		/// </remarks>
		uint Indent { get; set; }

		/// <summary>
		/// Gets or sets the indent level of cells in a table relative to their absolute
		/// indent level set by <see cref="Indent"/>.
		/// </summary>
		/// <remarks>
		///	<para>
		/// <inheritdoc cref="Indent" path="/remarks"/>
		/// </para>
		/// <para>
		/// This property is only valid for differential table styles and has no effect
		/// when set on a cell style.
		/// </para>
		/// </remarks>
		int RelativeIndent { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not to justify text on the last
		/// line of text in a cell.
		/// </summary>
		/// <remarks>
		/// Typically only used in East Asian alignment contexts.
		/// </remarks>
		bool JustifyLastLine { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not to shrink text in a cell when
		/// it is too long to fit within the cell's width.
		/// </summary>
		/// <remarks>
		/// This property does not have an effect on cells with multiple lines of text.
		/// </remarks>
		bool ShrinkToFit { get; set; }

		/// <summary>
		/// Gets or sets the reading order of a cell.
		/// </summary>
		ReadingOrder ReadingOrder { get; set; }
	}
}