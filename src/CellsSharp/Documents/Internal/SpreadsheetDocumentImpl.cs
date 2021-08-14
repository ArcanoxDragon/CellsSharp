using System;
using System.Collections.Generic;
using Autofac;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using CellsSharp.Workbooks;
using DocumentFormat.OpenXml.Packaging;
using MsSpreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument;

namespace CellsSharp.Documents.Internal
{
	sealed class SpreadsheetDocumentImpl : PartRootBase, ISpreadsheetDocumentImpl
	{
		private readonly ILifetimeScope documentScope;

		public SpreadsheetDocumentImpl(
			IEnumerable<IDocumentSaveLoadHandler> saveLoadHandlers, IChangeNotifier changeNotifier,
			ILifetimeScope documentScope, MsSpreadsheetDocument msDocument, IWorkbook workbook
		) : base(saveLoadHandlers, changeNotifier)
		{
			this.documentScope = documentScope;

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

			if (OpenXmlPackage.CanSave)
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

		#region IDisposable

		private bool disposed;

		protected override void CheckDisposed()
		{
			if (this.disposed)
				throw new ObjectDisposedException(nameof(SpreadsheetDocumentImpl));

			base.CheckDisposed();
		}

		protected override void Dispose(bool disposing)
		{
			if (this.disposed)
				return;

			base.Dispose(disposing);

			if (disposing)
			{
				this.documentScope.Dispose();
			}

			this.disposed = true;
		}

		#endregion
	}
}