using System;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.ChangeTracking
{
	sealed class ChangeOccurredEventArgs : EventArgs
	{
		/// <inheritdoc />
		public ChangeOccurredEventArgs(Type changedPartType, OpenXmlPart changedPart)
		{
			ChangedPartType = changedPartType;
			ChangedPart = changedPart;
		}

		/// <inheritdoc />
		public ChangeOccurredEventArgs(Type changedPartType)
		{
			ChangedPartType = changedPartType;
		}

		public Type         ChangedPartType { get; }
		public OpenXmlPart? ChangedPart     { get; }
	}
}