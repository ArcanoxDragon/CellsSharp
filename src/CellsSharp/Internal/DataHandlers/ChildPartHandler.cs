using System.Collections.Generic;
using System.Linq;
using CellsSharp.Internal.ChangeTracking;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.DataHandlers
{
	/// <summary>
	/// Handles saving and loading a part that is a child of another part.
	/// </summary>
	/// <typeparam name="TParentPart">The type of the part that is a direct parent of the child part</typeparam>
	/// <typeparam name="TChildPart">The type of the chlid part that this handler can save and load</typeparam>
	abstract class ChildPartHandler<TParentPart, TChildPart> : PartHandlerBase<TChildPart>, IChildPartHandler<TParentPart>
		where TParentPart : OpenXmlPart
		where TChildPart : OpenXmlPart
	{
		#region IChildPartHandler

		protected ChildPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<TChildPart>> childPartHandlers,
			IList<IPartElementHandler<TChildPart>> partElementHandlers
		) : base(changeNotifier, childPartHandlers, partElementHandlers) { }

		protected ChildPartHandler(
			IChangeNotifier changeNotifier,
			IList<IChildPartHandler<TChildPart>> childPartHandlers
		) : base(changeNotifier, childPartHandlers) { }

		protected ChildPartHandler(
			IChangeNotifier changeNotifier,
			IList<IPartElementHandler<TChildPart>> partElementHandlers
		) : base(changeNotifier, partElementHandlers) { }

		protected ChildPartHandler(
			IChangeNotifier changeNotifier
		) : base(changeNotifier) { }

		/// <inheritdoc />
		public void PartSaving(TParentPart parentPart)
		{
			var childPart = parentPart.GetPartsOfType<TChildPart>().SingleOrDefault();

			if (PartHasData)
			{
				childPart ??= CreateChildPart(parentPart);

				SavePart(childPart);
			}
			else
			{
				if (childPart is not null)
					parentPart.DeletePart(childPart);
			}
		}

		/// <inheritdoc />
		public void PartLoading(TParentPart parentPart)
		{
			var childPart = parentPart.GetPartsOfType<TChildPart>().SingleOrDefault();

			if (childPart is not null)
				LoadPart(childPart);
		}

		#endregion

		protected abstract bool PartHasData { get; }

		protected abstract TChildPart CreateChildPart(TParentPart parentPart);
	}
}