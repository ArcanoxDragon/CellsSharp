using System;
using System.Collections.Generic;
using System.Linq;
using CellsSharp.Extensions;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.Utilities;
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

		protected PartHandlerBase(IChangeNotifier changeNotifier, IList<IChildPartHandler<TPart>> childPartHandlers)
			: this(changeNotifier, childPartHandlers, EmptyList<IPartElementHandler<TPart>>.Instance) { }

		protected PartHandlerBase(IChangeNotifier changeNotifier, IList<IPartElementHandler<TPart>> partElementHandlers)
			: this(changeNotifier, EmptyList<IChildPartHandler<TPart>>.Instance, partElementHandlers) { }

		protected PartHandlerBase(IChangeNotifier changeNotifier) : this(
			changeNotifier,
			EmptyList<IChildPartHandler<TPart>>.Instance,
			EmptyList<IPartElementHandler<TPart>>.Instance
		) { }

		protected IChangeNotifier ChangeNotifier { get; }

		protected virtual void SavePart(TPart part)
		{
			foreach (var handler in this.childPartHandlers)
				handler.PartSaving(part);

			if (ChangeNotifier.IsPartChanged(part) && this.partElementHandlers.Any())
			{
				if (this.partElementHandlers.Count(h => h.HandlesRootElement) > 1)
					throw new InvalidOperationException($"More than one root element handler found for part \"{typeof(TPart).Name}\"");

				using OpenXmlWriter writer = OpenXmlWriter.Create(part);

				writer.WriteStartDocument();

				var rootElementHandler = this.partElementHandlers.SingleOrDefault();

				if (rootElementHandler is not null)
				{
					rootElementHandler.WriteElement(writer);
				}
				else
				{
					if (part.RootElement is null)
						throw new InvalidOperationException($"The \"{typeof(TPart).Name}\" part does not have a root element");

					writer.WriteElement(part.RootElement, () => {
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

				using OpenXmlReader reader = OpenXmlReader.Create(part);

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