using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Autofac;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.IoC;
using CellsSharp.Worksheets;
using CellsSharp.Worksheets.Internal;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellsSharp.Workbooks.Internal
{
	sealed class SheetCollection : ISheetCollection, IDisposable
	{
		private readonly List<ILifetimeScope> openWorksheetScopes = new();

		public SheetCollection(ILifetimeScope documentScope, IChangeNotifier changeNotifier, WorkbookPart workbookPart)
		{
			DocumentScope = documentScope;
			ChangeNotifier = changeNotifier;
			WorkbookPart = workbookPart;
			Sheets = workbookPart.Workbook.Sheets ??= new Sheets();
		}

		private ILifetimeScope  DocumentScope  { get; }
		private IChangeNotifier ChangeNotifier { get; }
		private WorkbookPart    WorkbookPart   { get; }
		private Sheets          Sheets         { get; }

		#region IEnumerable<IWorksheetInfo>

		private IEnumerable<IWorksheetInfo> EnumerateWorksheets()
		{
			CheckDisposed();

			foreach (var sheet in Sheets.Elements<Sheet>())
				yield return new WorksheetInfo(sheet);
		}

		/// <inheritdoc />
		public IEnumerator<IWorksheetInfo> GetEnumerator() => EnumerateWorksheets().GetEnumerator();

		/// <inheritdoc />
		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

		#endregion

		#region ISheetCollection

		/// <inheritdoc />
		public IWorksheetInfo this[string name] => new WorksheetInfo(FindSheetBy(s => s.Name?.Value == name));

		/// <inheritdoc />
		public IWorksheetInfo this[uint index] => new WorksheetInfo(FindSheetBy(s => s.SheetId?.Value == index));

		/// <inheritdoc />
		public IWorksheet AddNew()
		{
			CheckDisposed();

			var sheetId = GetNextSheetId();

			return AddNew($"Sheet{sheetId}");
		}

		/// <inheritdoc />
		public IWorksheet AddNew(string name)
		{
			CheckDisposed();

			var worksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
			var sheetId = GetNextSheetId();
			var sheet = new Sheet {
				SheetId = sheetId,
				Name = name,
				Id = WorkbookPart.GetIdOfPart(worksheetPart),
			};

			Sheets.AppendChild(sheet);
			ChangeNotifier.NotifyOfChange(this, WorkbookPart);

			return OpenWorksheet(worksheetPart);
		}

		/// <inheritdoc />
		public IWorksheet Open(IWorksheetInfo worksheetInfo)
		{
			CheckDisposed();

			var sheet = FindSheetBy(s => s.SheetId?.Value == worksheetInfo.Index);

			if (sheet.Id?.Value is null || !WorkbookPart.TryGetPartById(sheet.Id.Value, out var part) || part is not WorksheetPart worksheetPart)
				throw new InvalidOperationException("A matching worksheet was not found in the workbook");

			return OpenWorksheet(worksheetPart);
		}

		/// <inheritdoc />
		public void Remove(string name) => RemoveSheetBy(s => s.Name?.Value == name);

		/// <inheritdoc />
		public void Remove(uint index) => RemoveSheetBy(s => s.SheetId?.Value == index);

		/// <inheritdoc />
		public void Remove(IWorksheetInfo worksheetInfo) => Remove(worksheetInfo.Index);

		#endregion

		#region Private Methods

		private Sheet FindSheetBy(Func<Sheet, bool> predicate)
		{
			CheckDisposed();

			var sheet = Sheets.Elements<Sheet>().SingleOrDefault(predicate);

			if (sheet is null)
				throw new InvalidOperationException("A matching sheet was not found in the workbook");

			return sheet;
		}

		private void RemoveSheetBy(Func<Sheet, bool> predicate)
		{
			CheckDisposed();

			var sheet = FindSheetBy(predicate);

			Sheets.RemoveChild(sheet);
			ChangeNotifier.NotifyOfChange(this, WorkbookPart);
		}

		private IWorksheet OpenWorksheet(WorksheetPart worksheetPart)
		{
			var worksheetScope = Services.CreateWorksheetScope(DocumentScope, worksheetPart);

			this.openWorksheetScopes.Add(worksheetScope);

			return worksheetScope.Resolve<IWorksheet>();
		}

		private uint GetNextSheetId()
		{
			if (!Sheets.Elements<Sheet>().Any())
				return 1;

			return Sheets.Elements<Sheet>().Max(s => s.SheetId?.Value ?? 0) + 1;
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

			foreach (var worksheetScope in this.openWorksheetScopes)
				worksheetScope.Dispose();

			this.openWorksheetScopes.Clear();

			this.disposed = true;
		}

		#endregion
	}
}