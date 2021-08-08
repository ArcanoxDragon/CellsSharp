namespace CellsSharp.Internal.DataHandlers
{
	public interface ISaveLoadHandler
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
}