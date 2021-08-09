using System;
using System.Collections.Generic;
using CellsSharp.Internal.ChangeTracking;

namespace CellsSharp.Internal.DataHandlers
{
	class PartRootBase : IDisposable
	{
#pragma warning disable 618 // Need non-derived type to accept all derived types
		private readonly IEnumerable<ISaveLoadHandler> saveLoadHandlers;
#pragma warning restore 618

		private readonly IChangeNotifier changeNotifier;

#pragma warning disable 618 // Need non-derived type to accept all derived types
		public PartRootBase(IEnumerable<ISaveLoadHandler> saveLoadHandlers, IChangeNotifier changeNotifier)
#pragma warning restore 618
		{
			this.saveLoadHandlers = saveLoadHandlers;
			this.changeNotifier = changeNotifier;
		}

		protected void SaveParts()
		{
			if (!this.changeNotifier.HasChanges)
				return;

			foreach (var handler in this.saveLoadHandlers)
				handler.Save();

			this.changeNotifier.MarkClean();
		}

		protected void LoadParts()
		{
			foreach (var handler in this.saveLoadHandlers)
				handler.Load();

			this.changeNotifier.MarkClean();
		}

		#region IDisposable

		private bool disposed;

		~PartRootBase() => Dispose(false);

		protected virtual void CheckDisposed()
		{
			if (this.disposed)
				throw new ObjectDisposedException(nameof(PartRootBase));
		}

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (this.disposed)
				return;

			if (disposing)
			{
				SaveParts();
			}

			this.disposed = true;
		}

		#endregion
	}
}