using System.Drawing;
using JetBrains.Annotations;

namespace CellsSharp.Styles
{
    [PublicAPI]
    public interface IFontProperties
    {
        /// <summary>
        /// Gets or sets a value indicating whether or not this font is bold.
        /// </summary>
        bool Bold { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not this font is italic.
		/// </summary>
        bool Italic { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not this font has a
		/// strikethrough.
		/// </summary>
        bool Strike { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not this font has an
		/// outline around characters.
		/// </summary>
        bool Outline { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not this font has a
		/// shadow behind text.
		/// </summary>
        bool Shadow { get; set; }

        /// <summary>
        /// Gets or sets the underline style for this font.
        /// </summary>
        UnderlineStyle Underline { get; set; }

        /// <summary>
        /// Gets or sets the vertical text alignment style for this font
        /// (i.e. if the font is superscript/subscript or not).
        /// </summary>
        VerticalTextAlignment VerticalTextAlignment { get; set; }

        /// <summary>
        /// Gets or sets the size of this font, in points.
        /// </summary>
        double FontSize { get; set; }

        /// <summary>
        /// Gets or sets the text color of this font.
        /// </summary>
        Color Color { get; set; }

        /// <summary>
        /// Gets or sets the family name of this font.
        /// </summary>
        string FontName { get; set; }

        /// <summary>
        /// Gets or sets the family number of this font.
        /// </summary>
        int FontFamilyNumbering { get; set; }

        /// <summary>
        /// Gets or sets the font scheme type for this font.
        /// </summary>
        FontScheme FontScheme { get; set; }
    }
}
