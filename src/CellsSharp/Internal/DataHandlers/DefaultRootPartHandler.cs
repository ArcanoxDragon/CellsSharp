using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.DataHandlers
{
	class DefaultRootPartHandler<TPart> : RootPartHandler<TPart> where TPart : OpenXmlPart
	{
		/// <inheritdoc />
		public DefaultRootPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<TPart>> childPartHandlers,
			IList<IPartElementHandler<TPart>> partElementHandlers,
			TPart handledPart
		) : base(changeNotifier, childPartHandlers, partElementHandlers)
		{
			HandledPart = handledPart;
		}

		/// <inheritdoc />
		protected override TPart HandledPart { get; }
	}
}