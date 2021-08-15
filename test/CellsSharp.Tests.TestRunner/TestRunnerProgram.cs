using CellsSharp.Tests.Documents;
using CellsSharp.Tests.TestRunner.Utilities;

namespace CellsSharp.Tests.TestRunner
{
	class TestRunnerProgram
	{
		public static void Main()
		{
			TestContextUtility.HookTestContextOutput();

			RunHugeSpreadsheetTests();
		}

		private static void RunHugeSpreadsheetTests()
		{
			var testFixture = new HugeSpreadsheetTests();

			testFixture.CreateHugeDocument_AllDifferentText();
		}
	}
}