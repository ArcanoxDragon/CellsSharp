using System;
using System.Collections.Generic;
using Autofac;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using CellsSharp.Workbooks;
using MsSpreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument;

namespace CellsSharp.Documents.Internal
{
	sealed class SpreadsheetDocumentImpl : ISpreadsheetDocumentImpl
	{
		private readonly ILifetimeScope          documentScope;
		private readonly IList<ISaveLoadHandler> saveLoadHandlers;
		private readonly IChangeNotifier         changeNotifier;

		public SpreadsheetDocumentImpl(
			ILifetimeScope documentScope, MsSpreadsheetDocument msDocument,
			IWorkbook workbook, IList<ISaveLoadHandler> saveLoadHandlers,
			IChangeNotifier changeNotifier)
		{
			this.documentScope = documentScope;
			this.saveLoadHandlers = saveLoadHandlers;
			this.changeNotifier = changeNotifier;

			MsDocument = msDocument;
			Workbook = workbook;
		}

		private MsSpreadsheetDocument MsDocument { get; }

		#region ISpreadsheetDocument

		public IWorkbook Workbook { get; }

		public void Save()
		{
			CheckDisposed();
			SaveParts();
			MsDocument.Save();
		}

		public void SaveAs(string path)
		{
			CheckDisposed();
			SaveParts();
			MsDocument.SaveAs(path);
		}

		public void Load()
		{
			CheckDisposed();
			LoadParts();
		}

		#endregion

		private void SaveParts()
		{
			if (!this.changeNotifier.HasChanges)
				return;

			foreach (var handler in this.saveLoadHandlers)
				handler.Save();

			this.changeNotifier.MarkClean();
		}

		private void LoadParts()
		{
			foreach (var handler in this.saveLoadHandlers)
				handler.Load();

			this.changeNotifier.MarkClean();
		}

		#region IDisposable

		private bool disposed;

		private void CheckDisposed()
		{
			if (this.disposed)
				throw new ObjectDisposedException(nameof(SpreadsheetDocumentImpl));
		}

		public void Dispose()
		{
			if (this.disposed)
				return;

			SaveParts();
			this.documentScope.Dispose();

			this.disposed = true;
		}

		#endregion
	}
}