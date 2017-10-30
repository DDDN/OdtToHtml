/*
DDDN.OdtToHtml.OdtNode
Copyright(C) 2017 Lukasz Jaskiewicz(lukasz @jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace DDDN.OdtToHtml
{
	public class OdtNode
	{
		public string HtmlTag { get; }
		public string OdtTag { get; }
		public string OdtElementClassName { get; set; }
		public string InnerText { get; set; }
		public OdtNode ParentNode { get; }
		public List<OdtNode> ChildNodes { get; set; } = new List<OdtNode>();
		public Dictionary<string, List<string>> Attrs { get; } = new
			Dictionary<string, List<string>>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, string> CssProps = new
			Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
		public List<(string type, string position)> TabStops = new
			List<(string type, string position)>();

		private OdtNode()
		{
		}

		public OdtNode(string odtTagName, string htmlTagName, string className, OdtNode parent = null)
		{
			if (string.IsNullOrWhiteSpace(odtTagName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(odtTagName));
			}

			if (string.IsNullOrWhiteSpace(htmlTagName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(htmlTagName));
			}

			OdtTag = odtTagName;
			HtmlTag = htmlTagName;
			OdtElementClassName = className;
			ParentNode = parent;

			if (ParentNode != null)
			{
				parent.ChildNodes.Add(this);
			}
		}

		public static void AddCssPropertyValue(OdtNode odtNode, string propName, string propValue)
		{
			if (odtNode.CssProps.ContainsKey(propName))
			{
				odtNode.CssProps[propName] = propValue;
			}
			else
			{
				odtNode.CssProps.Add(propName, propValue);
			}
		}
		public static void EnsureClassName(OdtNode odtNode)
		{

			if (string.IsNullOrWhiteSpace(odtNode.OdtElementClassName)
				&& !odtNode.CssProps.ContainsKey("class"))
			{
				var className = odtNode.HtmlTag + Guid.NewGuid().ToString("N");
				OdtNode.AddAttrValue(odtNode, "class", className);
			}
		}

		public static void AddTabStop(OdtNode odtNode, XElement tabStop)
		{
			var typeAttrVal = tabStop.Attribute(XName.Get("type", OdtXmlNamespaces.Style))?.Value;
			var positionAttrVal = tabStop.Attribute(XName.Get("position", OdtXmlNamespaces.Style))?.Value;

			if (positionAttrVal == null)
			{
				return;
			}

			if (typeAttrVal == null)
			{
				typeAttrVal = "left";
			}

			odtNode.TabStops.Add((typeAttrVal, positionAttrVal));
		}

		public static IEnumerable<string> GetAttrValuesOrDefault(OdtNode odtNode, string attrName)
		{
			if (odtNode.Attrs.TryGetValue(attrName, out List<string> values))
			{
				return values;
			}
			else
			{
				return default(IEnumerable<string>);
			}
		}

		public static string RenderCss(OdtNode odtRootNode)
		{
			if (odtRootNode == null)
			{
				throw new ArgumentNullException(nameof(odtRootNode));
			}

			var builder = new StringBuilder(16384);
			CssRenderingWalker(odtRootNode.ChildNodes, builder);
			return builder.ToString();
		}

		public static void CssRenderingWalker(IEnumerable<OdtNode> odtNodes, StringBuilder builder)
		{
			foreach (var childNode in odtNodes)
			{
				RenderCssStyle(childNode, builder);
				CssRenderingWalker(childNode.ChildNodes, builder);
			}
		}

		private static void RenderCssStyle(OdtNode odtNode, StringBuilder builder)
		{
			if (!odtNode.Attrs.ContainsKey("class") || odtNode.CssProps.Count == 0)
			{
				return;
			}

			builder
				.Append(Environment.NewLine)
				.Append(".")
				.Append(String.Join(" ", odtNode.Attrs["class"]))
				.Append(" {")
				.Append(Environment.NewLine);
			RenderCssStyleProperties(odtNode, builder);
			builder.Append(" }");
		}

		private static void RenderCssStyleProperties(OdtNode odtNode, StringBuilder builder)
		{
			foreach (var prop in odtNode.CssProps)
			{
				builder
					.Append(prop.Key)
					.Append(": ")
					.Append(prop.Value)
					.Append(";")
					.Append(Environment.NewLine);
			}
		}

		public static void AddAttrValue(OdtNode odtNode, string attrName, string attrVal)
		{
			if (attrName == null)
			{
				throw new ArgumentNullException(nameof(attrName));
			}

			if (odtNode.Attrs.ContainsKey(attrName))
			{
				odtNode.Attrs[attrName].Add(attrVal);
			}
			else
			{
				odtNode.Attrs[attrName] = new List<string> { attrVal };
			}
		}

		public static string RenderHtml(OdtNode odtNode)
		{
			if (odtNode == null)
				throw new ArgumentNullException(nameof(odtNode));

			var htmlBuilder = new StringBuilder(8192);
			HtmlNodeWalker(odtNode, htmlBuilder);
			return htmlBuilder.ToString();
		}

		private static void HtmlNodeWalker(OdtNode odtNode, StringBuilder builder)
		{
			if (odtNode.ChildNodes.Count > 0 || !string.IsNullOrWhiteSpace(odtNode.InnerText))
			{
				builder.Append("<")
					.Append(odtNode.HtmlTag)
					.Append(RenderAtributes(odtNode))
					.Append(">")
					.Append(String.Concat(odtNode.InnerText));

				foreach (var child in odtNode.ChildNodes)
				{
					HtmlNodeWalker(child, builder);
				}

				builder.Append("</")
					.Append(odtNode.HtmlTag)
					.Append(">");
			}
			else
			{
				builder.Append("<")
					.Append(odtNode.HtmlTag)
					.Append(RenderAtributes(odtNode))
					.Append("/>");
			}
		}

		private static string RenderAtributes(OdtNode odtNode)
		{
			var builder = new StringBuilder(96);

			foreach (var attr in odtNode.Attrs)
			{
				builder
					.Append(" ")
					.Append(attr.Key)
					.Append("=\"");
				bool firstName = true;

				foreach (var val in attr.Value)
				{
					if (firstName)
					{
						builder.Append(val);
						firstName = false;
					}
					else
					{
						builder
							.Append(" ")
							.Append(val);
					}
				}

				builder.Append("\"");
			}

			return builder.ToString();
		}
	}
}
