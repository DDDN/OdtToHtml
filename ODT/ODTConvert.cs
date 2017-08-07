/*
* DDDN.Office.ODT.ODTConvert
* 
* Copyright(C) 2017 Lukasz Jaskiewicz
* Author: Lukasz Jaskiewicz (lukasz@jaskiewicz.de, devdone@outlook.com)
*
* This program is free software; you can redistribute it and/or modify it under the terms of the
* GNU General Public License as published by the Free Software Foundation; version 2 of the License.
*
* This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
* warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License along with this program; if not, write
* to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace DDDN.Office.ODT
{
	public class ODTConvert : IODTConvert, IDisposable
	{
		private ZipArchive ODTZipArchive;

		private static readonly Dictionary<string, string> Tags = new Dictionary<string, string>()
		{
			["text"] = "article",
			["p"] = "p",
			["span"] = "span",
			["s"] = "span",
			["table"] = "table",
			["table-columns"] = "tr",
			["table-column"] = "th",
			["table-row"] = "tr",
			["table-cell"] = "td",
			["list"] = "ul",
			["list-item"] = "li"
		};

		private static readonly Dictionary<string, string> Attrs =
			 new Dictionary<string, string>()
			 {
				 ["style-name"] = "class"
			 };

		public ODTConvert(ZipArchive odtDocument)
		{
			ODTZipArchive = odtDocument ?? throw new System.ArgumentNullException(nameof(odtDocument));
		}

		public string GetCss()
		{
			throw new System.NotImplementedException();
		}

		public string GetHtml()
		{
			var contentEntry = ODTZipArchive.Entries
				 .Where(p => p.Name.Equals("content.xml", StringComparison.InvariantCultureIgnoreCase))
				 .FirstOrDefault();
			XDocument contentXDoc = null;

			using (var contentStream = contentEntry.Open())
			{
				contentXDoc = XDocument.Load(contentStream);
			}

			var contentEle = contentXDoc.Root
				 .Elements(XName.Get("body", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
				 .Elements(XName.Get("text", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
				 .First();

			var htmlEle = new XElement(Tags[contentEle.Name.LocalName]);
			NodeWalker(contentEle.Nodes(), htmlEle);

			return htmlEle.ToString(SaveOptions.DisableFormatting);
		}

		private void NodeWalker(IEnumerable<XNode> odNode, XElement htmlElement)
		{
			var childHtmlEle = htmlElement;

			foreach (var node in odNode)
			{
				if (node.NodeType == XmlNodeType.Text)
				{
					var textNode = node as XText;
					childHtmlEle.SetValue(childHtmlEle.Value + textNode.Value);
				}
				else if (node.NodeType == XmlNodeType.Element)
				{
					var elementNode = node as XElement;

					if (elementNode.Name.Equals(XName.Get("s", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")))
					{
						AddNbsp(elementNode, htmlElement);
					}
					else if (Tags.TryGetValue(elementNode.Name.LocalName, out string htmlTag))
					{
						childHtmlEle = new XElement(htmlTag);
						CopyAttributes(elementNode, childHtmlEle);
						htmlElement.Add(childHtmlEle);
						NodeWalker(elementNode.Nodes(), childHtmlEle);
					}
				}
			}
		}

		private static void AddNbsp(XElement odElement, XElement htmlElement)
		{
			var spacesValue = odElement.Attribute(XName.Get("c", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"))?.Value;
			int.TryParse(spacesValue, out int spacesCount);

			if (spacesCount == 0)
			{
				spacesCount++;
			}

			for (int i = 0; i < spacesCount; i++)
			{
				htmlElement.SetValue(htmlElement.Value + "&nbsp;");
			}
		}

		private static void CopyAttributes(XElement odElement, XElement htmlElement)
		{
			if (odElement.HasAttributes)
			{
				foreach (var attr in odElement.Attributes())
				{
					if (Attrs.TryGetValue(attr.Name.LocalName, out string htmlAttrName))
					{
						var htmlAttr = new XAttribute(htmlAttrName, attr.Value);
						htmlElement.Add(htmlAttr);
					}
				}
			}
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					// TODO: dispose managed state (managed objects).
					ODTZipArchive.Dispose();
				}

				// TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
				// TODO: set large fields to null.

				disposedValue = true;
			}
		}

		// TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
		// ~ODTConvert() {
		//   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
		//   Dispose(false);
		// }

		// This code added to correctly implement the disposable pattern.
		public void Dispose()
		{
			// Do not change this code. Put cleanup code in Dispose(bool disposing) above.
			Dispose(true);
			// TODO: uncomment the following line if the finalizer is overridden above.
			// GC.SuppressFinalize(this);
		}
		#endregion
	}
}
