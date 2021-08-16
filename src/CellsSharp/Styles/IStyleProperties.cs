using JetBrains.Annotations;

namespace CellsSharp.Styles
{
    [PublicAPI]
    public interface IStyleProperties
    {
        /// <summary>
        /// Gets an <see cref="IAlignmentProperties"/> object representing the
        /// text alignment properties for the style.
        /// </summary>
        IAlignmentProperties Alignment { get; }

        /// <summary>
        /// Gets an <see cref="IProtectionProperties"/> object representing the
        /// cell protection properties for the style.
        /// </summary>
        IProtectionProperties Protection { get; }

        /// <summary>
        /// Gets an <see cref="INumberingFormatProperties"/> object representing
        /// the numbering format properties for this style.
        /// </summary>
        INumberingFormatProperties NumberingFormat { get; }

        /// <summary>
        /// Gets an <see cref="IFontProperties"/> object representing the font
        /// properties for this style.
        /// </summary>
        IFontProperties Font { get; }

        /// <summary>
        /// Gets an <see cref="IFillProperties"/> object representing the fill
        /// properties for this style.
        /// </summary>
        IFillProperties Fill { get; }

        /// <summary>
        /// Gets an <see cref="IBorderProperties"/> object representing the border
        /// properties for this style.
        /// </summary>
        IBorderProperties Border { get; }
    }
}
