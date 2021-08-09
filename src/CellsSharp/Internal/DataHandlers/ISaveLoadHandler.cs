using System;

namespace CellsSharp.Internal.DataHandlers
{
	[Obsolete("Do not directly inherit from this interface; use IDocumentSaveLoadHandler or IWorksheetSaveLoadHandler instead")]
	interface ISaveLoadHandler
	{
		/// <summary>
		/// Called when the document is being saved.
		/// </summary>
		void Save();

		/// <summary>
		/// Called when the document is being loaded.
		/// </summary>
		void Load();
	}

#pragma warning disable 618
	interface IDocumentSaveLoadHandler : ISaveLoadHandler { }

	interface IWorksheetSaveLoadHandler : ISaveLoadHandler { }
#pragma warning restore 618
}