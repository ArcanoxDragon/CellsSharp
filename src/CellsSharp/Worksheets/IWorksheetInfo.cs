﻿using JetBrains.Annotations;

namespace CellsSharp.Worksheets
{
	[PublicAPI]
	public interface IWorksheetInfo
	{
		/// <summary>
		/// Gets the index of the worksheet in the workbook
		/// </summary>
		uint Index { get; }

		/// <summary>
		/// Gets the name of the worksheet
		/// </summary>
		string Name { get; }
	}
}