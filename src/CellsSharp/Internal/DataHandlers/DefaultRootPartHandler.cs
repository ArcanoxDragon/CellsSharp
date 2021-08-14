using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.DataHandlers
{
	class DefaultRootPartHandler<TPart> : RootPartHandler<TPart> where TPart : OpenXmlPart
	{
		public DefaultRootPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<TPart>> childPartHandlers,
			IList<IPartElementHandler<TPart>> partElementHandlers,
			TPart handledPart
		) : base(changeNotifier, childPartHandlers, partElementHandlers)
		{
			HandledPart = handledPart;
		}

		protected override TPart HandledPart { get; }
	}
}