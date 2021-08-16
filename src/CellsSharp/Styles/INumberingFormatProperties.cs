using JetBrains.Annotations;

namespace CellsSharp.Styles
{
    [PublicAPI]
    public interface INumberingFormatProperties
    {
        /// <summary>
        /// Gets or sets the format code representing how this numbering
        /// format will format numeric values as text.
        /// </summary>
        string FormatCode { get; set; }
    }
}
