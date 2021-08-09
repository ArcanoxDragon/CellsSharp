using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using JetBrains.Annotations;

namespace CellsSharp.Extensions
{
	static class OpenXmlReaderWriterExtensions
	{
		#region Reader.VisitChildren

		/// <summary>
		/// Enumerates through the children of the <paramref name="reader"/>'s current element.
		/// Each item returned will be the element type of one of the children, and the reader
		/// will be positioned at the start of that element until the <see cref="IEnumerable{T}"/>
		/// is advanced to the next item, or data is read for that element.
		/// </summary>
		internal static IEnumerable<Type> VisitChildren(this OpenXmlReader reader)
		{
			if (!reader.ReadFirstChild())
				yield break;

			do
			{
				yield return reader.ElementType;
			}
			while (reader.ReadNextSibling());
		}

		/// <summary>
		/// Enumerates through the children of the <paramref name="reader"/>'s current element.
		/// Each child is evaluated to see whether or not it is of type <typeparamref name="T"/>.
		/// If it is, <paramref name="visitChild"/> will be called while <paramref name="reader"/>
		/// is positioned at the start of that child. If not, the child is skipped.
		/// </summary>
		internal static void VisitChildren<T>(this OpenXmlReader reader, [InstantHandle] Action visitChild)
		{
			foreach (var elementType in reader.VisitChildren())
			{
				if (elementType != typeof(T))
					continue;

				visitChild();
			}
		}

		#endregion

		#region Attribute.Is

		internal static bool Is<T>(this OpenXmlAttribute attribute, string localName, out T value) where T : struct
		{
			value = default;

			if (attribute.LocalName != localName || attribute.Value is null)
				return false;

			if (typeof(T).IsEnum)
			{
				// Try to parse as EnumValue<T>

				var enumValue = new EnumValue<T> { InnerText = attribute.Value };

				if (enumValue.HasValue)
				{
					value = enumValue.Value;
					return true;
				}

				return false;
			}

			try
			{
				value = (T) Convert.ChangeType(attribute.Value, typeof(T));
				return true;
			}
			catch
			{
				return false;
			}
		}

		#endregion

		#region Writer.WriteElement

		/// <summary>
		/// Writes the start element for a new element of type <typeparamref name="T"/>, then calls
		/// <paramref name="writeElementChildren"/> to write the children of that element. Finishes
		/// by writing the matching end element for the initial start element.
		/// </summary>
		internal static void WriteElement<T>(this OpenXmlWriter writer, [InstantHandle] Action writeElementChildren)
			where T : OpenXmlElement, new()
		{
			writer.WriteStartElement(new T());
			{
				writeElementChildren();
			}
			writer.WriteEndElement();
		}

		/// <summary>
		/// Writes the start element for the provided <paramref name="element"/>, then calls
		/// <paramref name="writeElementChildren"/> to write the children of that element. Finishes
		/// by writing the matching end element for the initial start element.
		/// </summary>
		internal static void WriteElement<T>(this OpenXmlWriter writer, T element, [InstantHandle] Action writeElementChildren)
			where T : OpenXmlElement
		{
			writer.WriteStartElement(element);
			{
				writeElementChildren();
			}
			writer.WriteEndElement();
		}

		#endregion
	}
}