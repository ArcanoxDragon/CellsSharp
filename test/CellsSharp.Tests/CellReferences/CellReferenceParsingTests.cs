using CellsSharp.Cells;
using NUnit.Framework;

namespace CellsSharp.Tests.CellReferences
{
	[TestFixture]
	public class CellReferenceParsingTests
	{
		[TestCase(null)]
		[TestCase("")]
		[TestCase("!")]
		[TestCase("1A")]
		[TestCase("$$")]
		[TestCase("A$1$")]
		[TestCase("A1048577")]
		[TestCase("XFE1")]
		[TestCase("ZZZZ1")]
		public void TestSingleCellInvalid(string invalidAddress)
		{
			Assert.False(CellAddress.TryParse(invalidAddress, out var cellAddress));
			Assert.AreEqual(cellAddress, CellAddress.None);
		}

		[TestCase("A1", 1u, 1u)]
		[TestCase("$A1", 1u, 1u, true, false)]
		[TestCase("A$1", 1u, 1u, false, true)]
		[TestCase("$A$1", 1u, 1u, true, true)]
		[TestCase("B2", 2u, 2u)]
		[TestCase("A1048576", 1u, 1048576u)]
		[TestCase("AA1", 27u, 1u)]
		[TestCase("AAA1", 703u, 1u)]
		[TestCase("XFD1", 16384u, 1u)]
		public void TestSingleCell(string address, uint expectedColumn, uint expectedRow, bool expectColumnAbsolute = false, bool expectRowAbsolute = false)
		{
			Assert.True(CellAddress.TryParse(address, out var cellAddress, out var addressType));

			var isColumnAbsolute = ( addressType & AddressType.ColumnAbsolute ) > 0;
			var isRowAbsolute = ( addressType & AddressType.RowAbsolute ) > 0;

			Assert.AreEqual(expectedColumn, cellAddress.Column);
			Assert.AreEqual(expectedRow, cellAddress.Row);
			Assert.AreEqual(expectColumnAbsolute, isColumnAbsolute);
			Assert.AreEqual(expectRowAbsolute, isRowAbsolute);
		}

		[TestCase("")]
		[TestCase("1A")]
		[TestCase("A1:")]
		[TestCase("A1::B2")]
		public void TestCellRangeInvalid(string invalidRange)
		{
			Assert.False(CellRange.TryParse(invalidRange, out var cellRange));
			Assert.AreSame(cellRange, CellRange.Empty);
		}

		[TestCase("A1", 1u, 1u, 1u, 1u)]
		[TestCase("A1:A1", 1u, 1u, 1u, 1u)]
		[TestCase("A1:B2", 1u, 1u, 2u, 2u)]
		[TestCase("B2:A1", 1u, 1u, 2u, 2u)]
		[TestCase("B1:A2", 1u, 1u, 2u, 2u)]
		[TestCase("A2:B1", 1u, 1u, 2u, 2u)]
		[TestCase("A1:XFD1048576", 1u, 1u, 16384u, 1048576u)]
		public void TestCellRange(string range, uint expectedStartColumn, uint expectedStartRow, uint expectedEndColumn, uint expectedEndRow)
		{
			Assert.True(CellRange.TryParse(range, out var cellRange));

			Assert.AreEqual(expectedStartColumn, cellRange.Start.Column);
			Assert.AreEqual(expectedStartRow, cellRange.Start.Row);
			Assert.AreEqual(expectedEndColumn, cellRange.End.Column);
			Assert.AreEqual(expectedEndRow, cellRange.End.Row);
		}

		// TODO: CellReference tests
	}
}