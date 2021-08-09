using System;

namespace CellsSharp.Cells
{
	[Flags]
	public enum AddressType : byte
	{
		Relative,
		RowAbsolute,
		ColumnAbsolute,
		Absolute,
	}
}