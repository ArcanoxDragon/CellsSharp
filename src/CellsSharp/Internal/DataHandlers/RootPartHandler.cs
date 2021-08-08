using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.DataHandlers
{
	/// <summary>
	/// Handles saving and loading a root-level part that can be resolved without any parent dependencies
	/// </summary>
	/// <typeparam name="TPart">The type of the part this handler can save and load</typeparam>
	abstract class RootPartHandler<TPart> : PartHandlerBase<TPart>, ISaveLoadHandler
		where TPart : OpenXmlPart
	{
		protected RootPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<TPart>> childPartHandlers,
			IList<IPartElementHandler<TPart>> partElementHandlers
		) : base(changeNotifier, childPartHandlers, partElementHandlers) { }

		protected RootPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<TPart>> childPartHandlers
		) : base(changeNotifier, childPartHandlers) { }

		protected RootPartHandler(
			IChangeNotifier changeNotifier,
			IList<IPartElementHandler<TPart>> partElementHandlers
		) : base(changeNotifier, partElementHandlers) { }

		protected RootPartHandler(
			IChangeNotifier changeNotifier
		) : base(changeNotifier) { }

		/// <summary>
		/// Gets the <typeparamref name="TPart"/> instance for which this <see cref="RootPartHandler{TPart}"/>
		/// will handle saving and loading.
		/// </summary>
		protected abstract TPart HandledPart { get; }

		/// <inheritdoc />
		public virtual void Save()
		{
			SavePart(HandledPart);
		}

		/// <inheritdoc />
		public virtual void Load()
		{
			LoadPart(HandledPart);
		}
	}
}