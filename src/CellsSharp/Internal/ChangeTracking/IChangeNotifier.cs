using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace CellsSharp.Internal.ChangeTracking
{
	interface IChangeNotifier
	{
		/// <summary>
		/// Fired when a change occurs to a part or a part type.
		/// </summary>
		event EventHandler<ChangeOccurredEventArgs> ChangeOccurred;

		/// <summary>
		/// Gets whether or not any changes have occurred to a part since the
		/// change notifier was last marked as clean.
		/// </summary>
		bool HasChanges { get; }

		/// <summary>
		/// Gets an enumerable of parts that have been changed since the change
		/// notifier was last marked as clean.
		/// </summary>
		IEnumerable<OpenXmlPart> ChangedParts { get; }

		/// <summary>
		/// Gets an enumerable of the part types that have been changed since the
		/// change notifier was last marked as clean.
		/// </summary>
		IEnumerable<Type> ChangedPartTypes { get; }

		/// <summary>
		/// Returns whether or not there has been a change to the provided <paramref name="partType"/>
		/// since the change notifier was last marked as clean.
		/// </summary>
		bool IsPartChanged(Type partType);

		/// <summary>
		/// Returns whether or not there has been a change to the provided <paramref name="part"/>,
		/// or to the part's type, since the change notifier was last marked as clean.
		/// </summary>
		/// <param name="part"></param>
		/// <returns></returns>
		bool IsPartChanged(OpenXmlPart part);

		/// <summary>
		/// Marks the change notifier as clean, indicating that changes to all parts have been processed.
		/// </summary>
		void MarkClean();

		/// <summary>
		/// Marks the provided <paramref name="changedPart"/> as changed and notifies any listeners of
		/// the change.
		/// </summary>
		/// <param name="sender">The object that caused the change to occur.</param>
		/// <param name="changedPart">The part that was changed.</param>
		void NotifyOfChange(object sender, OpenXmlPart changedPart);

		/// <summary>
		/// Marks the provided <paramref name="changedPartType"/> as changed and notifies any listeners of
		/// the change.
		/// </summary>
		/// <param name="sender">The object that caused the change to occur.</param>
		/// <param name="changedPartType">The part type that was changed.</param>
		void NotifyOfChange(object sender, Type changedPartType);

#if NET5_0_OR_GREATER
		/// <summary>
		/// Returns whether or not there has been a change to the provided <typeparamref name="TPart"/> type
		/// since the change notifier was last marked as clean.
		/// </summary>
		bool IsPartChanged<TPart>()
			where TPart : OpenXmlPart
			=> IsPartChanged(typeof(TPart));

		/// <summary>
		/// Marks the provided <typeparamref name="TPart"/> type as changed and notifies any listeners of
		/// the change.
		/// </summary>
		/// <param name="sender">The object that caused the change to occur.</param>
		void NotifyOfChange<TPart>(object sender)
			where TPart : OpenXmlPart
			=> NotifyOfChange(sender, typeof(TPart));
#endif
	}

#if !NET5_0_OR_GREATER
	static class ChangeNotifierExtensions
	{
		/// <summary>
		/// Returns whether or not there has been a change to the provided <typeparamref name="TPart"/> type
		/// since the change notifier was last marked as clean.
		/// </summary>
		internal static bool IsPartChanged<TPart>(this IChangeNotifier changeNotifier)
			where TPart : OpenXmlPart
			=> changeNotifier.IsPartChanged(typeof(TPart));

		/// <summary>
		/// Marks the provided <typeparamref name="TPart"/> type as changed and notifies any listeners of
		/// the change.
		/// </summary>
		/// <param name="changeNotifier"></param>
		/// <param name="sender">The object that caused the change to occur.</param>
		internal static void NotifyOfChange<TPart>(this IChangeNotifier changeNotifier, object sender)
			where TPart : OpenXmlPart
			=> changeNotifier.NotifyOfChange(sender, typeof(TPart));
	}
#endif
}