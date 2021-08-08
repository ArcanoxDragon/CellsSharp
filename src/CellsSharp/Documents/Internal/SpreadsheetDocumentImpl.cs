using System;
using Autofac;
using CellsSharp.Workbooks;
using MsSpreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument;

namespace CellsSharp.Documents.Internal
{
	sealed class SpreadsheetDocumentImpl : ISpreadsheetDocument
	{
		private readonly ILifetimeScope documentScope;

		public SpreadsheetDocumentImpl(ILifetimeScope documentScope, MsSpreadsheetDocument msDocument, IWorkbook workbook)
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

			// TODO: Save internal parts

			MsDocument.Save();
		}

		public void SaveAs(string path)
		{
			CheckDisposed();

			// TODO: Save internal parts

			MsDocument.SaveAs(path);
		}

		#endregion

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

			this.documentScope.Dispose();

			this.disposed = true;
		}

		#endregion
	}
}