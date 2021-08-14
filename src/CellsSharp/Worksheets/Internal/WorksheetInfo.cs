using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellsSharp.Worksheets.Internal
{
	sealed class WorksheetInfo : IWorksheetInfo
	{
		public WorksheetInfo(Sheet sheet)
		{
			if (sheet.SheetId is null || sheet.Name?.Value is null || sheet.Id?.Value is null)
				throw new ArgumentException("The provided Sheet is not valid", nameof(sheet));

			Index = sheet.SheetId.Value;
			Name = sheet.Name.Value;
			RelationshipId = sheet.Id.Value;
		}

		public WorksheetInfo(uint index, string name, string relationshipId)
		{
			Index = index;
			Name = name;
			RelationshipId = relationshipId;
		}

		public uint Index { get; }

		public string Name { get; }

		public string RelationshipId { get; }

		#region Equality

		private bool Equals(WorksheetInfo other)
		{
			return RelationshipId == other.RelationshipId;
		}

		public override bool Equals(object? obj)
		{
			return ReferenceEquals(this, obj) || obj is WorksheetInfo other && Equals(other);
		}

		public override int GetHashCode()
		{
			return RelationshipId.GetHashCode();
		}

		public static bool operator ==(WorksheetInfo? left, WorksheetInfo? right)
		{
			return Equals(left, right);
		}

		public static bool operator !=(WorksheetInfo? left, WorksheetInfo? right)
		{
			return !Equals(left, right);
		}

		#endregion
	}
}