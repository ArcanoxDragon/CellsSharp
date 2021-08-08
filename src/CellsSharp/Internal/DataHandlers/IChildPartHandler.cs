using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.DataHandlers
{
	/// <summary>
	/// Represents an interface for services that can save and load child parts
	/// of a specific parent part type.
	/// </summary>
	/// <typeparam name="TParentPart">
	/// The type of the parent part of which this handler can handle one or more children
	/// </typeparam>
	interface IChildPartHandler<in TParentPart> where TParentPart : OpenXmlPart
	{
		/// <summary>
		/// Called when a parent <typeparamref name="TParentPart"/> is being saved.
		/// </summary>
		void PartSaving(TParentPart parentPart);

		/// <summary>
		/// Called when a parent <typeparamref name="TParentPart"/> is being loaded.
		/// </summary>
		void PartLoading(TParentPart parentPart);
	}
}