using System;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using CellsSharp.Cells;
using CellsSharp.Worksheets;
using Humanizer;
using NUnit.Framework;

namespace CellsSharp.Tests
{
	[TestFixture]
	public class HugeSpreadsheetTests
	{
		private const string HugeDocumentFilename = "HugeDocument.xlsx";

		private const int HugeDocumentWidth  = 20;
		private const int HugeDocumentHeight = 100000;

		private void CreateDocument(Action<IWorksheet> createWorksheet)
		{
			var memoryUsageStart = GC.GetTotalMemory(true);
			var overallTimer = Stopwatch.StartNew();

			var measureCreateDocument = TestUtility.Measure("Creating document");
			using var package = Package.Open(HugeDocumentFilename, FileMode.Create, FileAccess.ReadWrite);
			using var document = SpreadsheetDocument.Create(package);

			measureCreateDocument.Finish();

			var measureAddSheet = TestUtility.Measure("Adding test worksheet");
			var testSheet = document.Workbook.Sheets.AddNew("Test Sheet");

			measureAddSheet.Finish();

			TestUtility.Measure("Filling worksheet", () => createWorksheet(testSheet));
			TestUtility.Measure("Saving document", document.Save);

			overallTimer.Stop();
			var memoryUsageEnd = GC.GetTotalMemory(true);
			overallTimer.Start();

			TestUtility.Measure("Disposing document", document.Dispose);

			overallTimer.Stop();
			TestContext.Progress.WriteLine($"Total elapsed time: {overallTimer.Elapsed.Humanize(precision: 2)}");
			TestContext.Progress.WriteLine($"End memory consumption: {( memoryUsageEnd - memoryUsageStart ).Bytes().Humanize("0.00")}");
		}

		private void FillSheet(Action<CellAddress> fillCell)
		{
			const int TotalCells = HugeDocumentWidth * HugeDocumentHeight;
			int currentCell = 0;
			int lastProgressChunk = 0;

			void CheckProgress()
			{
				double progress = ++currentCell / (double) TotalCells;
				int progressChunk = (int) ( progress * 10 );

				if (progressChunk != lastProgressChunk)
				{
					TestContext.Progress.WriteLine($"{progressChunk * 10}% done (cell {currentCell} / {TotalCells})");
					lastProgressChunk = progressChunk;
				}
			}

			for (uint row = 1; row <= HugeDocumentHeight; row++)
			for (uint col = 1; col <= HugeDocumentWidth; col++)
			{
				var address = new CellAddress(row, col);

				fillCell(address);

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