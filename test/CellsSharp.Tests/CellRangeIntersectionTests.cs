using CellsSharp.Cells;
using NUnit.Framework;

namespace CellsSharp.Tests
{
	[TestFixture]
	public class CellRangeIntersectionTests
	{
		[TestCase("A1:H8", "E5:L12", "E5:H8")]
		[TestCase("O5:V12", "S1:Z8", "S5:V8")]
		[TestCase("A15:H22", "D18:E19", "D18:E19")]
		public void TestIntersection(string range1, string range2, string expectedIntersection)
		{
			Assert.True(CellRange.TryParse(range1, out var cellRange1));
			Assert.True(CellRange.TryParse(range2, out var cellRange2));
			Assert.True(CellRange.TryParse(expectedIntersection, out var expectedIntersectionRange));

			var intersection = cellRange1.GetIntersection(cellRange2);

			Assert.AreEqual(expectedIntersectionRange.Start, intersection.Start);
			Assert.AreEqual(expectedIntersectionRange.End, intersection.End);
		}
	}
}