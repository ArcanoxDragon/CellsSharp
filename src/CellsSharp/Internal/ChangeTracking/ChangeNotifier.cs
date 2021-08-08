using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.ChangeTracking
{
	sealed class ChangeNotifier : IChangeNotifier
	{
		private readonly HashSet<OpenXmlPart> changedParts     = new();
		private readonly HashSet<Type>        changedPartTypes = new();

		/// <inheritdoc />
		public event EventHandler<ChangeOccurredEventArgs>? ChangeOccurred;

		/// <inheritdoc />
		public bool HasChanges => this.changedParts.Any() ||
								  this.changedPartTypes.Any();

		/// <inheritdoc />
		public IEnumerable<OpenXmlPart> ChangedParts => this.changedParts;

		/// <inheritdoc />
		public IEnumerable<Type> ChangedPartTypes => this.changedPartTypes;

		/// <inheritdoc />
		public bool IsPartChanged(Type partType) => this.changedPartTypes.Contains(partType);

		/// <inheritdoc />
		public bool IsPartChanged(OpenXmlPart part) => this.changedParts.Contains(part) ||
													   this.changedPartTypes.Contains(part.GetType());

		/// <inheritdoc />
		public void MarkClean()
		{
			this.changedParts.Clear();
			this.changedPartTypes.Clear();
		}

		/// <inheritdoc />
		public void NotifyOfChange(object sender, OpenXmlPart changedPart)
		{
			this.changedParts.Add(changedPart);
			ChangeOccurred?.Invoke(sender, new ChangeOccurredEventArgs(changedPart.GetType(), changedPart));
		}

		/// <inheritdoc />
		public void NotifyOfChange(object sender, Type changedPartType)
		{
			this.changedPartTypes.Add(changedPartType);
			ChangeOccurred?.Invoke(sender, new ChangeOccurredEventArgs(changedPartType));
		}
	}
}