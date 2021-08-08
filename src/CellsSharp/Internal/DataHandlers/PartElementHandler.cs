using System;
using CellsSharp.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.DataHandlers
{
	abstract class PartElementHandler<TPart, TElement> : IPartElementHandler<TPart>
		where TPart : OpenXmlPart
		where TElement : OpenXmlElement, new()
	{
		#region IPartElementHandler

		/// <inheritdoc />
		public virtual uint ElementOrder => 0;

		/// <inheritdoc />
		public Type HandledElementType => typeof(TElement);

		/// <inheritdoc />
		public virtual bool HandlesRootElement => false;

		/// <inheritdoc />
		public void WriteElement(OpenXmlWriter writer)
		{
			writer.WriteElement(CreateElement(), () => {
				WriteElementData(writer);
			});
		}

		/// <inheritdoc />
		public void ReadElement(OpenXmlReader reader)
		{
			if (reader.ElementType != HandledElementType)
				throw new InvalidOperationException($"The reader's current element is not a(n) {typeof(TElement).Name} element.");

			ReadElementData(reader);
		}

		#endregion

		protected virtual TElement CreateElement() => new();

		protected abstract void WriteElementData(OpenXmlWriter writer);
		protected abstract void ReadElementData(OpenXmlReader reader);
	}
}