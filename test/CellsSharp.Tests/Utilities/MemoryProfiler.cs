using System;
using System.Threading;
using Humanizer;

namespace CellsSharp.Tests.Utilities
{
	sealed class MemoryProfiler
	{
		private static readonly TimeSpan DefaultMeasurementInterval = 250.Milliseconds();

		private readonly TimeSpan measurementInterval;

		private bool   running;
		private Thread workerThread;

		public MemoryProfiler() : this(DefaultMeasurementInterval) { }

		public MemoryProfiler(TimeSpan measurementInterval)
		{
			this.measurementInterval = measurementInterval;

			InitialMemory = GC.GetTotalMemory(true);
		}

		public long InitialMemory { get; }
		public long PeakMemory    { get; private set; }

		public void StartMeasurement()
		{
			if (this.running)
				return;

			this.running = true;
			this.workerThread = new Thread(WorkerThread);
			this.workerThread.Start();
		}

		public void RestartMeasurement()
		{
			if (this.running)
				return;

			ResetPeakMemory();
			StartMeasurement();
		}

		public void StopMeasurement()
		{
			this.running = false;
			this.workerThread.Join();
		}

		public void ResetPeakMemory()
		{
			PeakMemory = this.InitialMemory;
		}

		private void WorkerThread()
		{
			while (this.running)
			{
				var currentMemory = GC.GetTotalMemory(false) - this.InitialMemory;

				PeakMemory = Math.Max(PeakMemory, currentMemory);

				Thread.Sleep(this.measurementInterval);
			}
		}
	}
}