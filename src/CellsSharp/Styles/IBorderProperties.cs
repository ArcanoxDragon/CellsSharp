using System.Drawing;
using JetBrains.Annotations;

namespace CellsSharp.Styles
{
	[PublicAPI]
	public interface IBorderProperties
	{
		/// <summary>
		/// Gets or sets a value indicating whether or not this border style
		/// has diagonal borders from the bottom left to the top right of a
		/// cell.
		/// </summary>
		bool DiagonalUp { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not this border style
		/// has diagonal borders from the top left to the bottom right of a
		/// cell.
		/// </summary>
		bool DiagonalDown { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not this border style's
		/// outer borders (i.e. top, bottom, left, and right) apply only to the
		/// outline of a range of cells.
		/// </summary>
		/// <remarks>
		/// This property only has an effect when set on a differential table
		/// style.
		/// </remarks>
		bool Outline { get; set; }

		/// <summary>
		/// Gets an <see cref="IBorderSegmentProperties"/> object representing
		/// the border properties for the start (leading-edge) border.
		/// </summary>
		IBorderSegmentProperties Start { get; }

		/// <summary>
		/// Gets an <see cref="IBorderSegmentProperties"/> object representing
		/// the border properties for the end (trailing-edge) border.
		/// </summary>
		IBorderSegmentProperties End { get; }

		/// <summary>
		/// Gets an <see cref="IBorderSegmentProperties"/> object representing
		/// the border properties for the left side border.
		/// </summary>
		IBorderSegmentProperties Left { get; }

		/// <summary>
		/// Gets an <see cref="IBorderSegmentProperties"/> object representing
		/// the border properties for the right side border.
		/// </summary>
		IBorderSegmentProperties Right { get; }

		/// <summary>
		/// Gets an <see cref="IBorderSegmentProperties"/> object representing
		/// the border properties for the top side border.
		/// </summary>
		IBorderSegmentProperties Top { get; }

		/// <summary>
		/// Gets an <see cref="IBorderSegmentProperties"/> object representing
		/// the border properties for the bottom side border.
		/// </summary>
		IBorderSegmentProperties Bottom { get; }

		/// <summary>
		/// Gets an <see cref="IBorderSegmentProperties"/> object representing
		/// the border properties for the diagonal borders.
		/// </summary>
		/// <remarks>
		/// This property applies to diagonal borders in both directions. The
		/// <see cref="DiagonalUp"/> and <see cref="DiagonalDown"/> properties
		/// are used to determine which of the diagonal borders are shown.
		/// </remarks>
		IBorderSegmentProperties Diagonal { get; }

		/// <summary>
		/// Gets an <see cref="IBorderSegmentProperties"/> object representing
		/// the border properties for the vertical borders between cells.
		/// </summary>
		/// <remarks>
		/// This property only has an effect when set on a differential table
		/// style.
		/// </remarks>
		IBorderSegmentProperties Vertical { get; }

		/// <summary>
		/// Gets an <see cref="IBorderSegmentProperties"/> object representing
		/// the border properties for the horizontal borders between cells.
		/// </summary>
		/// <remarks>
		/// This property only has an effect when set on a differential table
		/// style.
		/// </remarks>
		IBorderSegmentProperties Horizontal { get; }
	}

	[PublicAPI]
	public interface IBorderSegmentProperties
	{
		/// <summary>
		/// Gets or sets the border style of this border segment.
		/// </summary>
		BorderStyle Style { get; set; }

		/// <summary>
		/// Gets or sets the line color of this border segment.
		/// </summary>
		Color Color { get; set; }
	}
}