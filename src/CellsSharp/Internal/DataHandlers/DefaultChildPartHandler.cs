using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.DataHandlers
{
	class DefaultChildPartHandler<TParentPart, TChildPart> : ChildPartHandler<TParentPart, TChildPart>
		where TParentPart : OpenXmlPart
		where TChildPart : OpenXmlPart, IFixedContentTypePart
	{
		public DefaultChildPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<TChildPart>> childPartHandlers,
			IList<IPartElementHandler<TChildPart>> partElementHandlers
		) : base(changeNotifier, childPartHandlers, partElementHandlers) { }

		protected override bool PartHasData => true;

		protected override TChildPart CreateChildPart(TParentPart parentPart)
			=> parentPart.AddNewPart<TChildPart>();
	}
}