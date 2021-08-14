using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Autofac;
using CellsSharp.Extensions;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using CellsSharp.IoC;
using CellsSharp.Worksheets;
using CellsSharp.Worksheets.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellsSharp.Workbooks.Internal
{
	sealed class SheetCollection : PartElementHandler<WorkbookPart, Sheets>, ISheetCollection, IDocumentSaveLoadHandler, IDisposable
	{
		private readonly List<IWorksheetInfo>                       workbookSheets      = new();
		private readonly Dictionary<IWorksheetInfo, ILifetimeScope> openWorksheetScopes = new();

		public SheetCollection(ILifetimeScope documentScope, IChangeNotifier changeNotifier, WorkbookPart workbookPart)
		{
			DocumentScope = documentScope;
			ChangeNotifier = changeNotifier;
			WorkbookPart = workbookPart;
		}

		private ILifetimeScope  DocumentScope  { get; }
		private IChangeNotifier ChangeNotifier { get; }
		private WorkbookPart    WorkbookPart   { get; }

		#region IEnumerable<IWorksheetInfo>

		public IEnumerator<IWorksheetInfo> GetEnumerator() => this.workbookSheets.GetEnumerator();

		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

		#endregion

		#region ISheetCollection

		public IWorksheetInfo this[string name] => FindSheetBy(s => s.Name == name);

		public IWorksheetInfo this[uint index] => FindSheetBy(s => s.Index == index);

		public IWorksheet AddNew()
		{
			CheckDisposed();

			var sheetId = GetNextSheetIndex();

			return AddNew($"Sheet{sheetId}");
		}

		public IWorksheet AddNew(string name)
		{
			CheckDisposed();

			var worksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
			var sheetIndex = GetNextSheetIndex();
			var relationshipId = WorkbookPart.GetIdOfPart(worksheetPart);
			var worksheetInfo = new WorksheetInfo(sheetIndex, name, relationshipId);

			this.workbookSheets.Add(worksheetInfo);
			ChangeNotifier.NotifyOfChange(this, WorkbookPart);

			return OpenWorksheet(worksheetInfo, worksheetPart);
		}

		public IWorksheet Open(IWorksheetInfo worksheetInfo)
		{
			CheckDisposed();

			if (!WorkbookPart.TryGetPartById(worksheetInfo.RelationshipId, out var part) || part is not WorksheetPart worksheetPart)
				throw new InvalidOperationException("A matching worksheet was not found in the workbook");

			var worksheet = OpenWorksheet(worksheetInfo, worksheetPart);

			worksheet.Load();

			return worksheet;
		}

		public void Remove(IWorksheetInfo worksheetInfo)
		{
			CheckDisposed();

			if (!this.workbookSheets.Contains(worksheetInfo))
				return;

			if (!WorkbookPart.TryGetPartById(worksheetInfo.RelationshipId, out var part) || part is not WorksheetPart worksheetPart)
				throw new InvalidOperationException("A matching worksheet was not found in the workbook");

			if (this.openWorksheetScopes.TryGetValue(worksheetInfo, out var worksheetScope))
			{
				worksheetScope.Dispose();
				this.openWorksheetScopes.Remove(worksheetInfo);
			}

			this.workbookSheets.Remove(worksheetInfo);
			WorkbookPart.DeletePart(worksheetPart);
			ChangeNotifier.NotifyOfChange(this, WorkbookPart);
		}

		#endregion

		#region ISaveLoadHandler

		public void Save()
		{
			foreach (var worksheetScope in this.openWorksheetScopes.Values)
			{
				var worksheet = worksheetScope.Resolve<IWorksheet>();

				worksheet.Save();
			}
		}

		public void Load()
		{
			// Do nothing; worksheets are loaded when they are opened
		}

		#endregion

		#region IPartElementHandler

		protected override void ReadElementData(OpenXmlReader reader)
		{
			this.workbookSheets.Clear();

			reader.VisitChildren<Sheet>(() => {
				var sheet = (Sheet) reader.LoadCurrentElement()!;

				this.workbookSheets.Add(new WorksheetInfo(sheet));
			});
		}

		protected override void WriteElementData(OpenXmlWriter writer)
		{
			var nameValue = new StringValue();
			var sheetIdValue = new UInt32Value();
			var relationshipIdValue = new StringValue();
			var sheet = new Sheet {
				Name = nameValue,
				SheetId = sheetIdValue,
				Id = relationshipIdValue,
			};

			foreach (var worksheetInfo in this.workbookSheets)
			{
				nameValue.Value = worksheetInfo.Name;
				sheetIdValue.Value = worksheetInfo.Index;
				relationshipIdValue.Value = worksheetInfo.RelationshipId;

				writer.WriteElement(sheet);
			}
		}

		#endregion

		#region Private Methods

		private IWorksheetInfo FindSheetBy(Func<IWorksheetInfo, bool> predicate)
		{
			CheckDisposed();

			var sheet = this.workbookSheets.SingleOrDefault(predicate);

			if (sheet is null)
				throw new InvalidOperationException("A matching sheet was not found in the workbook");

			return sheet;
		}

		private IWorksheetImpl OpenWorksheet(IWorksheetInfo worksheetInfo, WorksheetPart worksheetPart)
		{
			var worksheetScope = Services.CreateWorksheetScope(DocumentScope, worksheetPart);
			var worksheet = worksheetScope.Resolve<IWorksheetImpl>();

			this.openWorksheetScopes.Add(worksheetInfo, worksheetScope);

			return worksheet;
		}

		private uint GetNextSheetIndex()
		{
			if (!this.workbookSheets.Any())
				return 1;

			return this.workbookSheets.Max(s => s.Index) + 1;
		}

		#endregion

		#region IDisposable

		private bool disposed;

		private void CheckDisposed()
		{
			if (this.disposed)
				throw new ObjectDisposedException(nameof(SheetCollection));
		}

		public void Dispose()
		{
			if (this.disposed)
				return;

			foreach (var worksheetScope in this.openWorksheetScopes.Values)
				worksheetScope.Dispose();

			this.openWorksheetScopes.Clear();

			this.disposed = true;
		}

		#endregion
	}
}