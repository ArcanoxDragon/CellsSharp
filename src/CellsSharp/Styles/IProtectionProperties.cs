using JetBrains.Annotations;

namespace CellsSharp.Styles
{
    [PublicAPI]
    public interface IProtectionProperties
    {
        /// <summary>
        /// Gets or sets a value indicating whether or not a cell is locked.
        /// </summary>
        /// <remarks>
        /// When a cell in a protected worksheet is locked, the protection options
        /// for the sheet apply to that cell, prohibiting certain actions from
        /// being taken on it.
        /// </remarks>
        bool Locked { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not a cell is hidden.
        /// </summary>
        /// <remarks>
        /// When a cell in a protected worksheet is hidden, the cell's contents
        /// will not be shown in the formula bar of the editing application.
        /// </remarks>
        bool Hidden { get; set; }
    }
}
