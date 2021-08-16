using JetBrains.Annotations;

namespace CellsSharp.Styles
{
	#region Alignment

	[PublicAPI]
	public enum HorizontalAlignment
	{
		/// <summary>
		/// General alignment aligns content in a cell using the following rules:
		/// <list type="bullet">
		/// <item>Text is left-aligned</item>
		/// <item>Numbers, dates, and times are right-aligned</item>
		/// <item>Boolean values are center-aligned</item>
		/// </list>
		/// </summary>
		General,

		/// <summary>
		/// Left-aligns content in a cell.
		/// </summary>
		Left,

		/// <summary>
		/// Center-alignes content in a cell.
		/// </summary>
		Center,

		/// <summary>
		/// Right-aligns content in a cell.
		/// </summary>
		Right,

		/// <summary>
		/// Fills a cell by repeating the contents as many times as needed to
		/// fill the width of the cell. The contents will only be repeated in
		/// whole increments.
		/// </summary>
		/// <remarks>
		/// If multiple consecutive empty cells have the <see cref="Fill"/>
		/// alignment, and the left-most cell has a value, the contents will
		/// be repeated to fill the entire combined width of the blank cells.
		/// </remarks>
		Fill,

		/// <summary>
		/// Justifies multi-line content in a cell such that the starting and
		/// ending edges of each line are aligned with the starting and ending
		/// edges of the cell.
		/// </summary>
		/// <remarks>
		/// This alignment style has no effect on cells with only one line of
		/// text.
		/// </remarks>
		Justify,

		/// <summary>
		/// Center-aligns text within a consecutive range of cells with this
		/// alignment style.
		/// </summary>
		/// <remarks>
		/// When a cell with text is followed by one or more blank cells with
		/// this alignment style, the first cell's text is centered within the
		/// width of it and all following blank cells, similar to if the cells
		/// were merged, but without actually merging them.
		/// </remarks>
		CenterContinuous,

		/// <summary>
		/// Evenly distributes content within a cell. For each line of text in
		/// a cell with this alignment style, the words of the cell are evenly
		/// distributed along the width of the cell.
		/// </summary>
		Distributed,
	}

	[PublicAPI]
	public enum VerticalAlignment
	{
		/// <summary>
		/// Aligns the contents of a cell against the top edge.
		/// </summary>
		Top,

		/// <summary>
		/// Centers the contents of a cell within the height of the cell.
		/// </summary>
		Center,

		/// <summary>
		/// Aligns the contents of a cell against the bottom edge.
		/// </summary>
		Bottom,

		/// <summary>
		/// For horizontal cell text, the lines of the cell are evenly
		/// distributed along the height of the cell. For vertical cell text,
		/// this style behaves in a similar manner to <see cref="HorizontalAlignment.Justify"/>.
		/// </summary>
		Justify,

		/// <summary>
		/// For horizontal cell text, the lines of the cell are evenly
		/// distributed along the height of the cell. For vertical cell text,
		/// this style behaves in a similar manner to <see cref="HorizontalAlignment.Distributed"/>.
		/// </summary>
		Distributed,
	}

	[PublicAPI]
	public enum ReadingOrder : uint
	{
		/// <summary>
		/// The reading order of a cell is determined by examining the first
		/// character of the cell's text for a reading order mark.
		/// </summary>
		ContextDependent,

		/// <summary>
		/// The reading order of a cell is left-to-right.
		/// </summary>
		LeftToRight,

		/// <summary>
		/// The reading order of a cell is right-to-left.
		/// </summary>
		RightToLeft,
	}

	#endregion

	#region Font

	[PublicAPI]
	public enum UnderlineStyle
	{
		/// <summary>
		/// Text has a single underline that crosses through descenders of characters
		/// such as "g" and "p".
		/// </summary>
		Single,

		/// <summary>
		/// Text has a double underline that crosses through descenders of characters
		/// such as "g" and "p".
		/// </summary>
		Double,

		/// <summary>
		/// Text has a single underline that is drawn below descenders of characters
		/// such as "g" and "p".
		/// </summary>
		SingleAccounting,

		/// <summary>
		/// Text has a double underline that is drawn below descenders of characters
		/// such as "g" and "p".
		/// </summary>
		DoubleAccounting,

		/// <summary>
		/// Text has no underline.
		/// </summary>
		None,
	}

	[PublicAPI]
	public enum VerticalTextAlignment
	{
		/// <summary>
		/// Text is aligned at the baseline of each line and is sized normally.
		/// </summary>
		Baseline,

		/// <summary>
		/// Text is render as superscript, being raised up above the baseline and
		/// shrunk down in size (if possible).
		/// </summary>
		Superscript,

		/// <summary>
		/// Text is render as subscript, being lowered below the baseline and
		/// shrunk down in size (if possible).
		/// </summary>
		Subscript,
	}

	[PublicAPI]
	public enum FontScheme
	{
		/// <summary>
		/// The font does not belong to a font scheme.
		/// </summary>
		None,

		/// <summary>
		/// The font is a major font in a font scheme (such as for headings, etc.).
		/// </summary>
		Major,

		/// <summary>
		/// The font is a minor font in a font scheme (such as for body/paragraph
		/// text, etc.).
		/// </summary>
		Minor,
	}

	#endregion

	#region Fill

	[PublicAPI]
	public enum PatternStyle
	{
		/// <summary>
		/// Cells have no fill.
		/// </summary>
		None,

		/// <summary>
		/// Cells have a solid color fill, described by a pattern's
		/// foreground color.
		/// </summary>
		Solid,

		/// <summary>
		/// Cells have a fill with even dithering between the pattern's
		/// foreground and background colors.
		/// </summary>
		MediumGray,

		/// <summary>
		/// Cells have a fill with a dithering biased towards the pattern's
		/// background color.
		/// </summary>
		DarkGray,

		/// <summary>
		/// Cells have a fill with a dithering biased toward's the pattern's
		/// foreground color.
		/// </summary>
		LightGray,

		/// <summary>
		/// Cells have a fill with thick horizontal bands alternating between
		/// the pattern's foreground and background colors.
		/// </summary>
		DarkHorizontal,

		/// <summary>
		/// Cells have a fill with thick vertical bands alternating between
		/// the pattern's foreground and background colors.
		/// </summary>
		DarkVertical,

		/// <summary>
		/// Cells have a fill with thick diagonal bands, from the upper-left
		/// corner to the lower-right corner, alternating between the pattern's
		/// foreground and background colors.
		/// </summary>
		DarkDown,

		/// <summary>
		/// Cells have a fill with thick diagonal bands, from the lower-left
		/// corner to the upper-right corner, alternating between the pattern's
		/// foreground and background colors.
		/// </summary>
		DarkUp,

		/// <summary>
		/// Cells have a fill with a thick checkered pattern alternating between
		/// the pattern's foreground and background colors.
		/// </summary>
		DarkGrid,

		/// <summary>
		/// Cells have a fill with a thick trellis-like pattern alternating
		/// between the pattern's foreground and background colors.
		/// </summary>
		DarkTrellis,

		/// <summary>
		/// Cells have a fill with thin horizontal bands alternating between
		/// the pattern's foreground and background colors.
		/// </summary>
		LightHorizontal,

		/// <summary>
		/// Cells have a fill with thin vertical bands alternating between
		/// the pattern's foreground and background colors.
		/// </summary>
		LightVertical,

		/// <summary>
		/// Cells have a fill with thin diagonal bands, from the upper-left
		/// corner to the lower-right corner, alternating between the pattern's
		/// foreground and background colors.
		/// </summary>
		LightDown,

		/// <summary>
		/// Cells have a fill with thin diagonal bands, from the lower-left
		/// corner to the upper-right corner, alternating between the pattern's
		/// foreground and background colors.
		/// </summary>
		LightUp,

		/// <summary>
		/// Cells have a fill with a thin grid pattern, the lines being the
		/// pattern's background color, and the fill being the pattern's foreground
		/// color.
		/// </summary>
		LightGrid,

		/// <summary>
		/// Cells have a fill with a thin trellis-like pattern alternating
		/// between the pattern's foreground and background colors.
		/// </summary>
		LightTrellis,

		/// <summary>
		/// Cells have a fill with a dithering, 1/16th of which is the pattern's
		/// background color, and the rest is the pattern's foreground color.
		/// </summary>
		Gray125,

		/// <summary>
		/// Cells have a fill with a dithering, 1/8th of which is the pattern's
		/// background color, and the rest is the pattern's foreground color.
		/// </summary>
		Gray0625,
	}

	[PublicAPI]
	public enum GradientStyle
	{
		/// <summary>
		/// Represents a gradient that follows a straignt line between
		/// two points at a given angle.
		/// </summary>
		Linear,

		/// <summary>
		/// Represents a gradient that progresses inward from the edges
		/// of a cell towards the center.
		/// </summary>
		Path,
	}

	#endregion

	#region Border

	[PublicAPI]
	public enum BorderStyle
	{
		/// <summary>
		/// The style of the border segment is none (no visible border).
		/// </summary>
		None,

		/// <summary>
		/// The style of the border segment is a thin solid line.
		/// </summary>
		Thin,
		
		/// <summary>
		/// The style of the border is a medium-weight solid line.
		/// </summary>
		Medium,
		
		/// <summary>
		/// The style of the border is a thin dashed line.
		/// </summary>
		Dashed,
		
		/// <summary>
		/// The style of the border is a thin dotted line.
		/// </summary>
		Dotted,
		
		/// <summary>
		/// The style of the border is a thick solid line.
		/// </summary>
		Thick,
		
		/// <summary>
		/// The style of the border is a thin solid double line.
		/// </summary>
		Double,
		
		/// <summary>
		/// The style of the border is a thin hairline-dotted line.
		/// </summary>
		Hair,
		
		/// <summary>
		/// The style of the border is a medium-weight dashed line.
		/// </summary>
		MediumDashed,
		
		/// <summary>
		/// The style of the border is a thin line alternating between
		/// a dash and a dot.
		/// </summary>
		DashDot,

		/// <summary>
		/// The style of the border is a medium-weight line alternating
		/// between a dash and a dot.
		/// </summary>
		MediumDashDot,

		/// <summary>
		/// The style of the border is a thin line alternating between
		/// a dash and two dots.
		/// </summary>
		DashDotDot,

		/// <summary>
		/// The style of the border is a medium-weight line alternating
		/// between a dash and two dots.
		/// </summary>
		MediumDashDotDot,
		
		/// <summary>
		///
		/// The style of the border is a medium-weight line alternating
		/// between a dash and a dot, with the separation between the two
		/// being at an angle.
		/// </summary>
		SlantDashDot,
	}

	#endregion
}