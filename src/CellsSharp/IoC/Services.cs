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

		private static IRegistrationBuilder<TLimit, TActivatorData, TRegistrationStyle>
			InstancePerDocument<TLimit, TActivatorData, TRegistrationStyle>(
				this IRegistrationBuilder<TLimit, TActivatorData, TRegistrationStyle> builder
			) => builder.InstancePerMatchingLifetimeScope(DocumentLifetimeTag);

		private static IRegistrationBuilder<TLimit, TActivatorData, TRegistrationStyle>
			InstancePerWorksheet<TLimit, TActivatorData, TRegistrationStyle>(
				this IRegistrationBuilder<TLimit, TActivatorData, TRegistrationStyle> builder
			) => builder.InstancePerMatchingLifetimeScope(WorksheetLifetimeTag);

		private static IRegistrationBuilder<THandler, ConcreteReflectionActivatorData, SingleRegistrationStyle>
			RegisterChildPartHandler<TParentPart, THandler>(this ContainerBuilder builder)
			where TParentPart : OpenXmlPart
			where THandler : IChildPartHandler<TParentPart>
			=> builder.RegisterType<THandler>().As<IChildPartHandler<TParentPart>>();

		private static IRegistrationBuilder<THandler, ConcreteReflectionActivatorData, SingleRegistrationStyle>
			RegisterPartElementHandler<TParentPart, THandler>(this ContainerBuilder builder)
			where TParentPart : OpenXmlPart
			where THandler : IPartElementHandler<TParentPart>
			=> builder.RegisterType<THandler>().As<IPartElementHandler<TParentPart>>();

		#endregion

		private static readonly IContainer Container;

		static Services()
		{
			var builder = new ContainerBuilder();

			Container = builder.Build();
		}

		internal static ILifetimeScope CreateDocumentScope(MsSpreadsheetDocument document)
			=> Container.BeginLifetimeScope(DocumentLifetimeTag, builder => {
				builder.RegisterInstance(document);
				RegisterDocumentServices(builder);
			});

		private static void RegisterDocumentServices(ContainerBuilder builder)
		{
			builder.Register(c => {
				var document = c.Resolve<MsSpreadsheetDocument>();
				var workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();

				if (workbookPart.Workbook is null!)
					workbookPart.Workbook = new MsWorkbook();

				return workbookPart;
			}).InstancePerDocument();
			builder.RegisterType<SpreadsheetDocumentImpl>().As<ISpreadsheetDocument>().As<ISpreadsheetDocumentImpl>().InstancePerDocument();
			builder.RegisterType<ChangeNotifier>().As<IChangeNotifier>().InstancePerDocument();
			builder.RegisterType<Workbook>().As<IWorkbook>().InstancePerDocument();
			builder.RegisterType<SheetCollection>().As<ISheetCollection>().As<IDocumentSaveLoadHandler>().InstancePerDocument();

			// Workbook parts
			builder.RegisterType<WorkbookPartHandler>().As<IDocumentSaveLoadHandler>().InstancePerDocument();
			builder.RegisterChildPartHandler<WorkbookPart, StringTablePartHandler>().InstancePerDocument();
			builder.RegisterPartElementHandler<SharedStringTablePart, StringTable>().As<IStringTable>().InstancePerDocument();
		}

		internal static ILifetimeScope CreateWorksheetScope(ILifetimeScope documentScope, WorksheetPart worksheetPart)
			=> documentScope.BeginLifetimeScope(WorksheetLifetimeTag, builder => {
				builder.RegisterInstance(worksheetPart);
				RegisterWorksheetServices(builder);
			});

		private static void RegisterWorksheetServices(ContainerBuilder builder)
		{
			builder.Register(c => {
				var workbookPart = c.Resolve<WorkbookPart>();
				var worksheetPart = c.Resolve<WorksheetPart>();
				var sheet = workbookPart.GetSheetForWorksheetPart(worksheetPart)
							?? throw new DependencyResolutionException("Sheet element not found for the worksheet");

				return new WorksheetInfo(sheet);
			}).As<IWorksheetInfo>().InstancePerWorksheet();
			builder.RegisterType<Worksheet>().As<IWorksheet>().As<IWorksheetImpl>().InstancePerWorksheet();
			builder.RegisterType<ChangeNotifier>().As<IChangeNotifier>().InstancePerWorksheet();

			// Worksheet parts
			builder.RegisterType<WorksheetPartHandler>().As<IWorksheetSaveLoadHandler>().InstancePerWorksheet();
			builder.RegisterPartElementHandler<WorksheetPart, SheetData>().AsSelf().InstancePerWorksheet();
		}
	}
}