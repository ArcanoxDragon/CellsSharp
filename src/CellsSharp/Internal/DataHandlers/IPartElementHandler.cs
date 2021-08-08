using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.DataHandlers
{
	// ReSharper disable once UnusedTypeParameter
	interface IPartElementHandler<TPart> where TPart : OpenXmlPart
	{
		/// <summary>
		/// Gets the order the element should appear within the parent part's
		/// root element.
		/// </summary>
		uint ElementOrder { get; }

		/// <summary>
		/// Gets the type of element this <see cref="IPartElementHandler{TPart}"/>
		/// can read and write.
		/// </summary>
		Type HandledElementType { get; }

		/// <summary>
		/// Whether or not this <see cref="IPartElementHandler{TPart}"/> can handle
		/// reading and writing of the root element for a <typeparamref name="TPart"/>.
		/// </summary>
		bool HandlesRootElement { get; }

		/// <summary>
		/// Writes data to the provided <paramref name="writer"/> for an element.
		/// </summary>
		void WriteElement(OpenXmlWriter writer);

		/// <summary>
		/// Writes data to the provided <paramref name="reader"/> for an element.
		/// </summary>
		void ReadElement(OpenXmlReader reader);
	}
}