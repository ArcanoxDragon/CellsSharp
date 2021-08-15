using System.Reflection;
using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;

namespace CellsSharp.Tests.TestRunner.Utilities
{
	static class TestContextUtility
	{
		internal static void HookTestContextOutput()
		{
			var context = TestExecutionContext.CurrentContext;
			var listenerProperty = typeof(TestExecutionContext).GetProperty("Listener", BindingFlags.Instance | BindingFlags.NonPublic);
			var currentListener = (ITestListener?) listenerProperty?.GetValue(context);
			var consoleListener = new ConsoleOutputTestListener(currentListener);

			listenerProperty?.SetValue(context, consoleListener);
		}
	}
}