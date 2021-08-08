using Autofac;
using Autofac.Builder;
using Autofac.Core;
using CellsSharp.Documents;
using CellsSharp.Documents.Internal;
using CellsSharp.Extensions;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using CellsSharp.Workbooks;
using CellsSharp.Workbooks.Internal;
using CellsSharp.Worksheets;
using CellsSharp.Worksheets.Internal;
using DocumentFormat.OpenXml.Packaging;
using MsSpreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument;
using MsWorkbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;

namespace CellsSharp.IoC
{
	static class Services
	{
		internal const string DocumentLifetimeTag  = "Document";
		internal const string WorksheetLifetimeTag = "Worksheet";

		#region Registration Helpers

		private static void InstancePerDocument<TLimit, TActivatorData, TRegistrationStyle>(
			this IRegistrationBuilder<TLimit, TActivatorData, TRegistrationStyle> builder
		) => builder.InstancePerMatchingLifetimeScope(DocumentLifetimeTag);

		private static void InstancePerWorksheet<TLimit, TActivatorData, TRegistrationStyle>(
			this IRegistrationBuilder<TLimit, TActivatorData, TRegistrationStyle> builder
		) => builder.InstancePerMatchingLifetimeScope(WorksheetLifetimeTag);

		private static void RegisterSaveLoadHandler<THandler>(this ContainerBuilder builder)
			where THandler : ISaveLoadHandler
			=> builder.RegisterType<THandler>().As<ISaveLoadHandler>().InstancePerDocument();

		private static void RegisterChildPartHandler<TParentPart, THandler>(this ContainerBuilder builder)
			where TParentPart : OpenXmlPart
			where THandler : IChildPartHandler<TParentPart>
			=> builder.RegisterType<THandler>().As<IChildPartHandler<TParentPart>>().InstancePerDocument();

		#endregion

		private static readonly IContainer Container;

		static Services()
		{
			var builder = new ContainerBuilder();

			RegisterDocumentServices(builder);
			RegisterWorksheetServices(builder);

			Container = builder.Build();
		}

		private static void RegisterDocumentServices(ContainerBuilder builder)
		{
			builder.Register(c => {
				var document = c.Resolve<MsSpreadsheetDocument>();
				var workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();

				if (workbookPart.Workbook is null!)
					workbookPart.Workbook = new MsWorkbook();

				return workbookPart;
			}).InstancePerDocument();
			builder.RegisterType<ChangeNotifier>().As<IChangeNotifier>().InstancePerDocument();
			builder.RegisterType<SpreadsheetDocumentImpl>().As<ISpreadsheetDocument>().As<ISpreadsheetDocumentImpl>().InstancePerDocument();
			builder.RegisterType<Workbook>().As<IWorkbook>().InstancePerDocument();
			builder.RegisterType<SheetCollection>().As<ISheetCollection>().InstancePerDocument();

			// Part handlers
			builder.RegisterSaveLoadHandler<WorkbookPartHandler>();
			builder.RegisterChildPartHandler<WorkbookPart, StringTablePartHandler>();

			// Element handlers
			builder.RegisterType<StringTable>().As<IStringTable>().As<IPartElementHandler<SharedStringTablePart>>().InstancePerDocument();
		}

		private static void RegisterWorksheetServices(ContainerBuilder builder)
		{
			builder.Register(c => {
				var workbookPart = c.Resolve<WorkbookPart>();
				var worksheetPart = c.Resolve<WorksheetPart>();
				var sheet = workbookPart.GetSheetForWorksheetPart(worksheetPart)
							?? throw new DependencyResolutionException("Sheet element not found for the worksheet");

				return new WorksheetInfo(sheet);
			}).As<IWorksheetInfo>().InstancePerWorksheet();
			builder.RegisterType<WorksheetManager>().As<IWorksheet>().InstancePerWorksheet();
		}

		internal static ILifetimeScope CreateDocumentScope(MsSpreadsheetDocument document)
			=> Container.BeginLifetimeScope(DocumentLifetimeTag, builder => {
				builder.RegisterInstance(document);
			});

		internal static ILifetimeScope CreateWorksheetScope(ILifetimeScope documentScope, WorksheetPart worksheetPart)
			=> documentScope.BeginLifetimeScope(WorksheetLifetimeTag, builder => {
				builder.RegisterInstance(worksheetPart);
			});
	}
}