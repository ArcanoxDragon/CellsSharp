using System.IO;
using System.IO.Packaging;
using Autofac;
using CellsSharp.Documents;
using CellsSharp.Documents.Internal;
using CellsSharp.IoC;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using JetBrains.Annotations;
using MsSpreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument;

namespace CellsSharp
{
	[PublicAPI]
	public static class SpreadsheetDocument
	{
		#region Creating

		public static ISpreadsheetDocument Create(
			Package package,
			SpreadsheetDocumentType documentType = SpreadsheetDocumentType.Workbook,
			bool autoSave = true,
			CompressionOption compressionOption = CompressionOption.Normal)
		{
			var msDocument = MsSpreadsheetDocument.Create(package, documentType, autoSave);

			msDocument.CompressionOption = compressionOption;

			return ActivateDocumentInstance(msDocument);
		}

		public static ISpreadsheetDocument Create(
			Stream stream,
			SpreadsheetDocumentType documentType = SpreadsheetDocumentType.Workbook,
			bool autoSave = true,
			CompressionOption compressionOption = CompressionOption.Normal)
		{
			var msDocument = MsSpreadsheetDocument.Create(stream, documentType, autoSave);

			msDocument.CompressionOption = compressionOption;

			return ActivateDocumentInstance(msDocument);
		}

		public static ISpreadsheetDocument Create(
			string path,
			SpreadsheetDocumentType documentType = SpreadsheetDocumentType.Workbook,
			bool autoSave = true,
			CompressionOption compressionOption = CompressionOption.Normal)
		{
			var msDocument = MsSpreadsheetDocument.Create(path, documentType, autoSave);

			msDocument.CompressionOption = compressionOption;

			return ActivateDocumentInstance(msDocument);
		}

		#endregion

		#region Opening

		public static ISpreadsheetDocument Open(Package package, OpenSettings? openSettings = null)
		{
			var document = ActivateDocumentInstance(MsSpreadsheetDocument.Open(package, openSettings ?? new OpenSettings()));

			document.Load();

			return document;
		}

		public static ISpreadsheetDocument Open(Stream stream, bool isEditable = false, OpenSettings? openSettings = null)
		{
			var document = ActivateDocumentInstance(MsSpreadsheetDocument.Open(stream, isEditable, openSettings ?? new OpenSettings()));

			document.Load();

			return document;
		}

		public static ISpreadsheetDocument Open(string path, bool isEditable = false, OpenSettings? openSettings = null)
		{
			var document = ActivateDocumentInstance(MsSpreadsheetDocument.Open(path, isEditable, openSettings ?? new OpenSettings()));

			document.Load();

			return document;
		}

		#endregion

		private static ISpreadsheetDocumentImpl ActivateDocumentInstance(MsSpreadsheetDocument msDocument)
		{
			var documentScope = Services.CreateDocumentScope(msDocument);

			return documentScope.Resolve<ISpreadsheetDocumentImpl>();
		}
	}
}