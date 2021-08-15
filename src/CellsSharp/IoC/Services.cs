using Autofac;
using Autofac.Builder;
using CellsSharp.Documents;
using CellsSharp.Documents.Internal;
using CellsSharp.Internal.ChangeTracking;
using CellsSharp.Internal.DataHandlers;
using CellsSharp.Workbooks;
using CellsSharp.Workbooks.Internal;
using CellsSharp.Worksheets;
using CellsSharp.Worksheets.Internal;
using DocumentFormat.OpenXml.Packaging;
using MsSpreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument;
using SheetData = CellsSharp.Worksheets.Internal.SheetData;
using Workbook = CellsSharp.Workbooks.Internal.Workbook;
using Worksheet = CellsSharp.Worksheets.Internal.Worksheet;

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

		/// <summary>
		/// Registers all services that exist within the scope of a single document
		/// </summary>
		private static void RegisterDocumentServices(ContainerBuilder builder)
		{
			builder.Register(c => {
				var document = c.Resolve<MsSpreadsheetDocument>();
				var workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();

				return workbookPart;
			}).InstancePerDocument();
			builder.RegisterType<SpreadsheetDocumentImpl>().As<ISpreadsheetDocument>().As<ISpreadsheetDocumentImpl>().InstancePerDocument();
			builder.RegisterType<ChangeNotifier>().As<IChangeNotifier>().InstancePerDocument();
			builder.RegisterType<Workbook>().As<IWorkbook>().InstancePerDocument();

			// Workbook parts
			builder.RegisterType<WorkbookPartHandler>().As<IDocumentSaveLoadHandler>().InstancePerDocument();
			builder.RegisterPartElementHandler<WorkbookPart, SheetCollection>()
				   .As<ISheetCollection>()
				   .As<IDocumentSaveLoadHandler>()
				   .InstancePerDocument();
			builder.RegisterChildPartHandler<WorkbookPart, StringTablePartHandler>().InstancePerDocument();
			builder.RegisterPartElementHandler<SharedStringTablePart, StringTable>().As<IStringTable>().InstancePerDocument();
		}

		internal static ILifetimeScope CreateWorksheetScope(ILifetimeScope documentScope, WorksheetPart worksheetPart)
			=> documentScope.BeginLifetimeScope(WorksheetLifetimeTag, builder => {
				builder.RegisterInstance(worksheetPart);
				RegisterWorksheetServices(builder);
			});

		/// <summary>
		/// Registers all services that exist within the scope of a single worksheet
		/// </summary>
		private static void RegisterWorksheetServices(ContainerBuilder builder)
		{
			builder.RegisterType<Worksheet>().As<IWorksheet>().As<IWorksheetImpl>().InstancePerWorksheet();
			builder.RegisterType<ChangeNotifier>().As<IChangeNotifier>().InstancePerWorksheet();

			// Worksheet parts
			builder.RegisterType<WorksheetPartHandler>().As<IWorksheetSaveLoadHandler>().InstancePerWorksheet();
			builder.RegisterPartElementHandler<WorksheetPart, SheetData>().AsSelf().InstancePerWorksheet();
			builder.RegisterPartElementHandler<WorksheetPart, MergedCellsCollection>().AsSelf().InstancePerWorksheet();
		}
	}
}