using System;
using NUnit.Framework.Interfaces;

namespace CellSharp.Tests.TestRunner
{
	public class ConsoleOutputTestListener : ITestListener
	{
		private readonly ITestListener? innerListener;

		public ConsoleOutputTestListener(ITestListener? innerListener)
		{
			this.innerListener = innerListener;
		}

		/// <inheritdoc />
		public void TestStarted(ITest test)
		{
			this.innerListener?.TestStarted(test);
		}

		/// <inheritdoc />
		public void TestFinished(ITestResult result)
		{
			this.innerListener?.TestFinished(result);
		}

		/// <inheritdoc />
		public void TestOutput(TestOutput output)
		{
			Console.Write(output.Text);
			this.innerListener?.TestOutput(output);
		}

		/// <inheritdoc />
		public void SendMessage(TestMessage message)
		{
			Console.Write(message.Message);
			this.innerListener?.SendMessage(message);
		}
	}
}