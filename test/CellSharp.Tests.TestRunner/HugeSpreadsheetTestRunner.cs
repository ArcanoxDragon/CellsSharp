using CellsSharp.Tests;

namespace CellSharp.Tests.TestRunner
{
	public class HugeSpreadsheetTestRunner
	{
		public static void Main()
		{
			TestContextUtility.HookTestContextOutput();

			var testFixture = new HugeSpreadsheetTests();

			testFixture.CreateHugeDocument_AllDifferentText();
		}
	}
}