using System;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.ChangeTracking
{
	sealed class ChangeOccurredEventArgs : EventArgs
	{
		public ChangeOccurredEventArgs(Type changedPartType, OpenXmlPart changedPart)
		{
			ChangedPartType = changedPartType;
			ChangedPart = changedPart;
		}

		public ChangeOccurredEventArgs(Type changedPartType)
		{
			ChangedPartType = changedPartType;
		}

		public Type         ChangedPartType { get; }
		public OpenXmlPart? ChangedPart     { get; }
	}
}