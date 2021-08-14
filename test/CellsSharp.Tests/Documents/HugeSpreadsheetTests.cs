using System;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using CellsSharp.Cells;
using CellsSharp.Tests.Utilities;
using CellsSharp.Worksheets;
using Humanizer;
using NUnit.Framework;

namespace CellsSharp.Tests.Documents
{
	[TestFixture]
	public class HugeSpreadsheetTests
	{
		private const string HugeDocumentFilename = "HugeDocument.xlsx";

		private const int HugeDocumentWidth  = 20;
		private const int HugeDocumentHeight = 100000;

		private void CreateDocument(Action<IWorksheet> createWorksheet)
		{
			var memoryProfiler = new MemoryProfiler();
			var overallTimer = Stopwatch.StartNew();

			void PauseTimerWhile(Action action)
			{
				var wasTimerRunning = overallTimer.IsRunning;

				if (wasTimerRunning)
					overallTimer.Stop();

				action();

				if (wasTimerRunning)
					overallTimer.Start();
			}

			memoryProfiler.StartMeasurement();

			var measureCreateDocument = TestUtility.Measure("Creating document");

			using var package = Package.Open(HugeDocumentFilename, FileMode.Create, FileAccess.ReadWrite);
			using var document = SpreadsheetDocument.Create(package);

			PauseTimerWhile(() => {
				measureCreateDocument.Finish();
				memoryProfiler.StopMeasurement();
			});

			var createDocumentPeakMemory = memoryProfiler.PeakMemory;

			memoryProfiler.RestartMeasurement();

			var measureAddSheet = TestUtility.Measure("Adding test worksheet");
			var testSheet = document.Workbook.Sheets.AddNew("Test Sheet");

			PauseTimerWhile(() => {
				measureAddSheet.Finish();
				memoryProfiler.StopMeasurement();
			});

			var addTestSheetPeakMemory = memoryProfiler.PeakMemory;

			memoryProfiler.RestartMeasurement();
			TestUtility.Measure("Filling worksheet", () => createWorksheet(testSheet));
			PauseTimerWhile(memoryProfiler.StopMeasurement);

			var fillWorksheetPeakMemory = memoryProfiler.PeakMemory;

			memoryProfiler.RestartMeasurement();
			TestUtility.Measure("Saving document", () => document.Save());
			memoryProfiler.StopMeasurement();

			var saveDocumentPeakMemory = memoryProfiler.PeakMemory;

			memoryProfiler.RestartMeasurement();
			TestUtility.Measure("Disposing document", document.Dispose);
			overallTimer.Stop();
			memoryProfiler.StopMeasurement();

			var disposeDocumentPeakMemory = memoryProfiler.PeakMemory;

			TestContext.Progress.WriteLine($"Total elapsed time: {overallTimer.Elapsed.Humanize(precision: 2)}");
			TestContext.Progress.WriteLine($"End memory consumption: {( memoryProfiler.PeakMemory - memoryProfiler.InitialMemory ).Bytes().Humanize("0.00")}");
			TestContext.Progress.WriteLine($"Peak memory while creating document: {createDocumentPeakMemory.Bytes().Humanize("0.00")}");
			TestContext.Progress.WriteLine($"Peak memory while adding test sheet: {addTestSheetPeakMemory.Bytes().Humanize("0.00")}");
			TestContext.Progress.WriteLine($"Peak memory while filling worksheet: {fillWorksheetPeakMemory.Bytes().Humanize("0.00")}");
			TestContext.Progress.WriteLine($"Peak memory while saving document: {saveDocumentPeakMemory.Bytes().Humanize("0.00")}");
			TestContext.Progress.WriteLine($"Peak memory while disposing document: {disposeDocumentPeakMemory.Bytes().Humanize("0.00")}");
		}

		private void FillSheet(Action<CellAddress> fillCell)
		{
			const int TotalCells = HugeDocumentWidth * HugeDocumentHeight;
			int currentCell = 0;
			int lastProgressChunk = 0;

			void CheckProgress()
			{
				double progress = currentCell / (double) TotalCells;
				int progressChunk = (int) ( progress * 10 );

				if (progressChunk != lastProgressChunk)
				{
					TestContext.Progress.WriteLine($"{progressChunk * 10}% done (cell {currentCell} / {TotalCells})");
					lastProgressChunk = progressChunk;
				}
			}

			for (uint row = 1; row <= HugeDocumentHeight; row++)
			{
				for (uint col = 1; col <= HugeDocumentWidth; col++)
				{
					var address = new CellAddress(row, col);

					fillCell(address);
					currentCell++;
				}

				if (row % 1000 == 0 || currentCell == TotalCells)
					CheckProgress();
			}
		}

		[Test, Explicit]
		public void CreateHugeDocument_AllDifferentText()
			=> CreateDocument(testSheet => {
				TestContext.Progress.WriteLine();

				FillSheet(address => {
					testSheet[address].CellText = $"Cell {address}";
				});
			});
	}
}