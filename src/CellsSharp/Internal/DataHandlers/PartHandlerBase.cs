using System;
using System.Collections.Generic;
using System.Linq;
using CellsSharp.Extensions;
using CellsSharp.Internal.ChangeTracking;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.DataHandlers
{
	abstract class PartHandlerBase<TPart> where TPart : OpenXmlPart
	{
		protected IList<IChildPartHandler<TPart>>   childPartHandlers;
		protected IList<IPartElementHandler<TPart>> partElementHandlers;

		protected PartHandlerBase(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<TPart>> childPartHandlers,
			IList<IPartElementHandler<TPart>> partElementHandlers
		)
		{
			this.childPartHandlers = childPartHandlers;
			this.partElementHandlers = partElementHandlers;

			ChangeNotifier = changeNotifier;
		}

		protected IChangeNotifier ChangeNotifier { get; }

		protected virtual OpenXmlElement CreateRootElement()
		{
			throw new InvalidOperationException($"There was no root element handler for part \"{typeof(TPart).Name}\" and its part handler " +
												$"does not implement {nameof(CreateRootElement)}.");
		}

		protected virtual void SavePart(TPart part)
		{
			foreach (var handler in this.childPartHandlers)
				handler.PartSaving(part);

			if (ChangeNotifier.IsPartChanged(part) && this.partElementHandlers.Any())
			{
				if (this.partElementHandlers.Count(h => h.HandlesRootElement) > 1)
					throw new InvalidOperationException($"More than one root element handler found for part \"{typeof(TPart).Name}\"");

				using var writer = OpenXmlWriter.Create(part);

				writer.WriteStartDocument();

				var rootElementHandler = this.partElementHandlers.SingleOrDefault(h => h.HandlesRootElement);

				if (rootElementHandler is not null)
				{
					rootElementHandler.WriteElement(writer);
				}
				else
				{
					writer.WriteElement(CreateRootElement(), () => {
						foreach (var handler in this.partElementHandlers.OrderBy(h => h.ElementOrder))
							handler.WriteElement(writer);
					});
				}
			}
		}

		protected virtual void LoadPart(TPart part)
		{
			foreach (var handler in this.childPartHandlers)
				handler.PartLoading(part);

			if (this.partElementHandlers.Any())
			{
				var elementHandlers = this.partElementHandlers.ToDictionary(h => h.HandledElementType, h => h);

				using var reader = OpenXmlReader.Create(part);

				// Move to the root element
				if (!reader.Read())
					return;

				if (elementHandlers.TryGetValue(reader.ElementType, out var rootElementHandler) && rootElementHandler.HandlesRootElement)
				{
					rootElementHandler.ReadElement(reader);
				}
				else
				{
					// Read the children if present
					foreach (var elementType in reader.VisitChildren())
					{
						// Try and find a handler to read this element
						if (elementHandlers.TryGetValue(elementType, out var handler))
							handler.ReadElement(reader);
					}
				}
			}
		}
	}
}