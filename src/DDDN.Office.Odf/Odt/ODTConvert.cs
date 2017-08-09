/*
* DDDN.Office.Odf.Odt.ODTConvert
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

using DDDN.Logging.Messages;
using DDDN.Office.Odf.Style;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DDDN.Office.Odf.Odt
{
	public class ODTConvert : IODTConvert, IDisposable
	{
		private ZipArchive ODTZipArchive;

		private static readonly Dictionary<string, string> HtmlTags = new Dictionary<string, string>()
		{
			["text"] = "article",
			["h"] = "p",
			["p"] = "p",
			["span"] = "span",
			["s"] = "span",
			["table"] = "table",
			["table-columns"] = "tr",
			["table-column"] = "th",
			["table-row"] = "tr",
			["table-cell"] = "td",
			["list"] = "ul",
			["list-item"] = "li",
			["automatic-styles"] = "style"
		};

		private static readonly Dictionary<string, string> CssNames =
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
			XDocument contentXDoc = GetZipArchiveEntryAsXDocument("content.xml");
			//XDocument stylesXDoc = GetZipArchiveEntryAsXDocument("styles.xml);

			var styleElements = contentXDoc.Root
				.Elements(XName.Get("automatic-styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
				.Elements()
				.Where(p => p.Name.Equals(XName.Get("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")));

			if (!styleElements.Any())
			{
				return String.Empty;
			}

			List<IODTStyle> Styles = new List<IODTStyle>();
			StylesWalker(styleElements, Styles);
			var css = RenderCss(Styles);
			return css;
		}

		private string RenderCss(List<IODTStyle> styles)
		{
			var builder = new StringBuilder(1024);

			foreach (var style in styles)
			{
				builder.Append($"{Environment.NewLine}.{style.Name} {{{Environment.NewLine}");

				foreach (var attr in style.Attributes)
				{
					builder.Append($"{attr.Key}: {attr.Value};{Environment.NewLine}");
				}
				builder.Append("}");
			}

			return builder.ToString();
		}

		private void StylesWalker(IEnumerable<XNode> node, List<IODTStyle> styles)
		{
			foreach (var n in node)
			{
				if (n.NodeType == XmlNodeType.Element)
				{
					var elementNode = n as XElement;
					IODTStyle style = new ODTStyle(elementNode);
					styles.Add(style);
					StyleNodesWalker(elementNode.Nodes(), style);
				}
			}
		}

		private void StyleNodesWalker(IEnumerable<XNode> node, IODTStyle style)
		{
			foreach (var n in node)
			{
				if (n.NodeType == XmlNodeType.Element)
				{
					var elementNode = n as XElement;
					style.AddAttributes(elementNode);
					StyleNodesWalker(elementNode.Nodes(), style);
				}
			}
		}

		public string GetHtml()
		{
			XDocument contentXDoc = GetZipArchiveEntryAsXDocument("content.xml");

			var contentEle = contentXDoc.Root
				  .Elements(XName.Get("body", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
				  .Elements(XName.Get("text", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
				  .First();

			var htmlEle = new XElement(HtmlTags[contentEle.Name.LocalName]);
			ContentNodesWalker(contentEle.Nodes(), htmlEle);

			return htmlEle.ToString(SaveOptions.DisableFormatting);
		}

		private XDocument GetZipArchiveEntryAsXDocument(string entryName)
		{
			if (string.IsNullOrWhiteSpace(entryName))
			{
				throw new ArgumentException(LogMsg.StrArgNullOrWhite, nameof(entryName));
			}

			var contentEntry = ODTZipArchive.Entries
								  .Where(p => p.Name.Equals(entryName, StringComparison.InvariantCultureIgnoreCase))
								  .FirstOrDefault();

			XDocument contentXDoc = null;

			using (var contentStream = contentEntry.Open())
			{
				contentXDoc = XDocument.Load(contentStream);
			}

			return contentXDoc;
		}

		private void ContentNodesWalker(IEnumerable<XNode> odNode, XElement htmlElement)
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
					else if (HtmlTags.TryGetValue(elementNode.Name.LocalName, out string htmlTag))
					{
						childHtmlEle = new XElement(htmlTag);
						CopyAttributes(elementNode, childHtmlEle);
						htmlElement.Add(childHtmlEle);
						ContentNodesWalker(elementNode.Nodes(), childHtmlEle);
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
					if (CssNames.TryGetValue(attr.Name.LocalName, out string htmlAttrName))
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
