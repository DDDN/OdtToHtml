/*
DDDN.OdtToHtml.OdtNode
Copyright(C) 2017-2018 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace DDDN.OdtToHtml
{
	public class OdtNode
	{
		public static class CssPrefix
		{
			public const string Element = "";
			public const string Id = "#";
			public const string Class = ".";
		}

		public string HtmlTag { get; }
		public string OdtTag { get; }
		public string OdtElementClassName { get; }
		public string InnerText { get; set; }
		public OdtNode ParentNode { get; }
		public OdtNode PreviousSameHierarchyNode { get; }
		public List<OdtNode> ChildNodes { get; } = new List<OdtNode>();
		public Dictionary<string, List<string>> Attrs { get; } = new
			Dictionary<string, List<string>>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, Dictionary<string, string>> CssProps { get; } = new
			Dictionary<string, Dictionary<string, string>>(StringComparer.InvariantCultureIgnoreCase);
		public List<(string type, string position)> TabStops = new
			List<(string type, string position)>();

		private OdtNode()
		{
		}

		public OdtNode(string odtTagName, string htmlTagName, string odtClassName, OdtNode parentNode = null)
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
			ParentNode = parentNode;

			var className = "";

			if (string.IsNullOrWhiteSpace(odtClassName))
			{
				className = $"{HtmlTag}_{Guid.NewGuid().ToString("N")}";
				OdtElementClassName = "";

			}
			else
			{
				className = odtClassName;
				OdtElementClassName = className;
			}

			OdtNode.AddAttrValue(this, "class", className);

			if (ParentNode != null)
			{
				PreviousSameHierarchyNode = parentNode.ChildNodes.LastOrDefault();
				parentNode.ChildNodes.Add(this);
			}
		}

		public static void AddCssPropertyValue(
			OdtNode odtNode,
			string cssName,
			string propName,
			string propValue,
			string prefix = CssPrefix.Class)
		{
			if (odtNode == null)
			{
				throw new ArgumentNullException(nameof(odtNode));
			}

			if (string.IsNullOrWhiteSpace(cssName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(cssName));
			}

			if (string.IsNullOrWhiteSpace(propName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(propName));
			}

			if (string.IsNullOrWhiteSpace(propValue))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(propValue));
			}

			if (prefix == null)
			{
				throw new ArgumentException(nameof(prefix));
			}

			odtNode.CssProps.TryGetValue(prefix + cssName, out Dictionary<string, string> cssProperties);

			if (cssProperties == null)
			{
				cssProperties = new Dictionary<string, string>();
				odtNode.CssProps.Add(prefix + cssName, cssProperties);
			}

			if (cssProperties.ContainsKey(propName))
			{
				cssProperties[propName] = propValue;
			}
			else
			{
				cssProperties.Add(propName, propValue);
			}
		}

		public static void AddTabStop(OdtNode odtNode, XElement tabStop)
		{
			var typeAttrVal = tabStop.Attribute(XName.Get("type", OdtXmlNs.Style))?.Value;
			var positionAttrVal = tabStop.Attribute(XName.Get("position", OdtXmlNs.Style))?.Value;

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

		public string GetClassName()
		{
			Attrs.TryGetValue("class", out List<string> className);
			return className.FirstOrDefault();
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

		private static string NormalizeClassName(string name)
		{
			return name.Trim().Replace(".", "");
		}

		private static void RenderCssStyle(OdtNode odtNode, StringBuilder builder)
		{
			if (odtNode.CssProps.Count == 0)
			{
				return;
			}

			foreach (var cssProp in odtNode.CssProps)
			{
				builder
				.Append(Environment.NewLine)
				.Append(cssProp.Key)
				.Append(" {")
				.Append(Environment.NewLine);
				RenderCssStyleProperties(cssProp.Value, builder);
				builder.Append(" }");
			}
		}

		private static void RenderCssStyleProperties(Dictionary<string, string> cssProperties, StringBuilder builder)
		{
			foreach (var prop in cssProperties)
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
			if (string.IsNullOrWhiteSpace(attrName))
			{
				throw new ArgumentNullException(nameof(attrName));
			}

			if (string.IsNullOrWhiteSpace(attrVal))
			{
				return;
			}

			// TODO adding and overiding values have to be implemented example: style= adding, src = overriding

			if (attrName.Equals("class", StringComparison.InvariantCultureIgnoreCase))
			{
				attrVal = NormalizeClassName(attrVal);
			}

			if (odtNode.Attrs.ContainsKey(attrName))
			{
				if (!odtNode.Attrs[attrName].Contains(attrVal))
				{
					odtNode.Attrs[attrName].Add(attrVal);
				}
			}
			else
			{
				odtNode.Attrs[attrName] = new List<string> { attrVal };
			}
		}

		public static string RenderHtml(OdtNode odtNode)
		{
			if (odtNode == null)
			{
				throw new ArgumentNullException(nameof(odtNode));
			}

			var htmlBuilder = new StringBuilder(8192);
			HtmlNodeWalker(odtNode, htmlBuilder);
			return htmlBuilder.ToString();
		}

		private static void HtmlNodeWalker(OdtNode odtNode, StringBuilder builder)
		{
			if ((odtNode.ChildNodes.Count > 0 || !string.IsNullOrWhiteSpace(odtNode.InnerText)))
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
