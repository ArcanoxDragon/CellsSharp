using CellSharp.Tests.TestRunner.Utilities;
using CellsSharp.Tests.Documents;

namespace CellSharp.Tests.TestRunner
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