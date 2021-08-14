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

		public event EventHandler<ChangeOccurredEventArgs>? ChangeOccurred;

		public bool HasChanges => this.changedParts.Any() ||
								  this.changedPartTypes.Any();

		public IEnumerable<OpenXmlPart> ChangedParts     => this.changedParts;
		public IEnumerable<Type>        ChangedPartTypes => this.changedPartTypes;

		public bool IsPartChanged(Type partType) => this.changedPartTypes.Contains(partType);

		public bool IsPartChanged(OpenXmlPart part) => this.changedParts.Contains(part) ||
													   this.changedPartTypes.Contains(part.GetType());

		public void MarkClean()
		{
			this.changedParts.Clear();
			this.changedPartTypes.Clear();
		}

		public void NotifyOfChange(object sender, OpenXmlPart changedPart)
		{
			this.changedParts.Add(changedPart);
			ChangeOccurred?.Invoke(sender, new ChangeOccurredEventArgs(changedPart.GetType(), changedPart));
		}

		public void NotifyOfChange(object sender, Type changedPartType)
		{
			this.changedPartTypes.Add(changedPartType);
			ChangeOccurred?.Invoke(sender, new ChangeOccurredEventArgs(changedPartType));
		}
	}
}