using System.Diagnostics.CodeAnalysis;
using JetBrains.Annotations;

namespace CellsSharp.Workbooks
{
	[PublicAPI]
	public interface IStringTable
	{
		/// <summary>
		/// Gets the value of the string at the provided <paramref name="index"/>.
		/// </summary>
		string this[uint index] { get; }

		/// <summary>
		/// Gets the number of entries currently present in the string table.
		/// </summary>
		uint EntryCount { get; }

		/// <summary>
		/// Removes all values from the string table.
		/// </summary>
		void Clear();

		/// <summary>
		/// Gets the value of the string at the provided <paramref name="index"/>
		/// and stores it in <paramref name="value"/>. Returns true if the string
		/// exists or false if it does not.
		/// </summary>
		bool TryGetValue(uint index, [MaybeNullWhen(false)] out string value);

		/// <summary>
		/// Inserts the provided <paramref name="value"/> if it does not already exist.
		/// Returns the index of the string in the string table.
		/// </summary>
		uint GetOrInsertValue(string value);
	}
}