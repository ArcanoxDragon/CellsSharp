using System;
using System.Diagnostics;
using Humanizer;
using JetBrains.Annotations;
using NUnit.Framework;

namespace CellsSharp.Tests.Utilities
{
	sealed class MeasurementUnit
	{
		private readonly Stopwatch stopwatch = new();
		private readonly string    taskName;

		public MeasurementUnit(string taskName)
		{
			this.taskName = taskName;
		}

		public void Start()
		{
			TestContext.Progress.Write($"{this.taskName}...");
			this.stopwatch.Start();
		}

		public void Finish()
		{
			this.stopwatch.Stop();
			TestContext.Progress.WriteLine($"Done. Took: {this.stopwatch.Elapsed.Humanize(precision: 2)}");
		}
	}

	static class TestUtility
	{
		internal static MeasurementUnit Measure(string taskName)
		{
			var unit = new MeasurementUnit(taskName);

			unit.Start();

			return unit;
		}

		internal static void Measure(string taskName, [InstantHandle] Action task)
		{
			var unit = new MeasurementUnit(taskName);

			try
			{
				unit.Start();
				task();
			}
			finally
			{
				unit.Finish();
			}
		}
	}
}