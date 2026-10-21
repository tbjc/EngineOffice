using System;
using DocumentFormat.OpenXml;

namespace OfficeEngine.Word;

/// <summary>
/// Inserts a child at the position a schema sequence requires.
/// </summary>
/// <remarks>
/// The Open XML SDK does not reorder the children of an element, and many WordprocessingML composite
/// types are schema sequences, so appending a property can produce a package that fails validation.
/// Property writers in this project therefore insert following an explicit order table instead.
/// A local name that is not in the table sorts last, which keeps unknown markup before the new child
/// rather than after it.
/// </remarks>
internal static class SchemaOrder
{
    /// <summary>Inserts <paramref name="element"/> into <paramref name="parent"/> following
    /// <paramref name="order"/>.</summary>
    /// <param name="parent">The composite element to insert into.</param>
    /// <param name="element">The element to insert.</param>
    /// <param name="order">The XML local names of the allowed children, in schema order.</param>
    internal static void Insert(OpenXmlCompositeElement parent, OpenXmlElement element, string[] order)
    {
        if (element.Parent is not null)
        {
            element.Remove();
        }

        int rank = RankOf(order, element.LocalName);
        foreach (OpenXmlElement child in parent.ChildElements)
        {
            if (RankOf(order, child.LocalName) > rank)
            {
                parent.InsertBefore(element, child);
                return;
            }
        }

        parent.AppendChild(element);
    }

    private static int RankOf(string[] order, string localName)
    {
        int index = Array.IndexOf(order, localName);
        return index < 0 ? int.MaxValue : index;
    }
}
