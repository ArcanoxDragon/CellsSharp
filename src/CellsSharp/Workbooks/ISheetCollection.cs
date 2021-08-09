using System.Collections.Generic;
using CellsSharp.Worksheets;
using JetBrains.Annotations;

namespace CellsSharp.Workbooks
{
	[PublicAPI]
	// TODO: Documentation
	public interface ISheetCollection : IEnumerable<IWorksheetInfo>
	{
		/// <summary>
		/// Returns a <see cref="IWorksheetInfo"/> representing the worksheet with
		/// the provided <paramref name="name"/>.
		/// </summary>
		IWorksheetInfo this[string name] { get; }

		/// <summary>
		/// Returns a <see cref="IWorksheetInfo"/> representing the worksheet with
		/// the provided <paramref name="index"/>.
		/// </summary>
		IWorksheetInfo this[uint index] { get; }

		/// <summary>
		/// Adds a new worksheet. The name will be "SheetN" where "N" is the next
		/// available sheet number.
		/// </summary>
		IWorksheet AddNew();

		/// <summary>
		/// Adds a new worksheet with the provided <paramref name="name"/>.
		/// </summary>
		IWorksheet AddNew(string name);

		/// <summary>
		/// Opens an instance of <see cref="IWorksheet"/> for the worksheet the provided
		/// <paramref name="worksheetInfo"/> represents
		/// </summary>
		IWorksheet Open(IWorksheetInfo worksheetInfo);

		/// <summary>
		/// Removes the worksheet represented by the provided <paramref name="worksheetInfo"/>.
		/// </summary>
		void Remove(IWorksheetInfo worksheetInfo);

#if NET5_0_OR_GREATER
		/// <summary>
		/// Opens the worksheet with the provided <paramref name="name"/>.
		/// </summary>
		IWorksheet Open(string name) => Open(this[name]);

		/// <summary>
		/// Opens the worksheet with the provided <paramref name="index"/>.
		/// </summary>
		IWorksheet Open(uint index) => Open(this[index]);

		/// <summary>
		/// Removes the worksheet with the provided <paramref name="name"/>.
		/// </summary>
		void Remove(string name) => Remove(this[name]);

		/// <summary>
		/// Removes the worksheet with the provided <paramref name="index"/>.
		/// </summary>
		void Remove(uint index) => Remove(this[index]);
#endif
	}

#if !NET5_0_OR_GREATER
	[PublicAPI]
	public static class SheetCollectionExtensions
	{
		/// <summary>
		/// Opens the worksheet with the provided <paramref name="name"/>.
		/// </summary>
		public static IWorksheet Open(this ISheetCollection sheetCollection, string name) => sheetCollection.Open(sheetCollection[name]);

		/// <summary>
		/// Opens the worksheet with the provided <paramref name="index"/>.
		/// </summary>
		public static IWorksheet Open(this ISheetCollection sheetCollection, uint index) => sheetCollection.Open(sheetCollection[index]);

		/// <summary>
		/// Removes the worksheet with the provided <paramref name="name"/>.
		/// </summary>
		public static void Remove(this ISheetCollection sheetCollection, string name) => sheetCollection.Remove(sheetCollection[name]);

		/// <summary>
		/// Removes the worksheet with the provided <paramref name="index"/>.
		/// </summary>
		public static void Remove(this ISheetCollection sheetCollection, uint index) => sheetCollection.Remove(sheetCollection[index]);
	}
#endif
}