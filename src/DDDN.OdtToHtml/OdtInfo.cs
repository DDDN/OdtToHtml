/*
DDDN.OdtToHtml.OdtInfo
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
	public class OdtInfo
	{
		public enum PseudoClass
		{
			None,
			Before
		}

		public string HtmlNode { get; }
		public string OdtNode { get; }
		public string ClassName { get; }
		public string Label { get; set; }
		public OdtInfo ParentNode { get; }
		public OdtInfo PreviousSibling { get; }
		public List<OdtInfo> ChildNodes { get; } = new List<OdtInfo>();
		public Dictionary<string, List<string>> CssAttrs { get; } = new Dictionary<string, List<string>>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, string> OdtAttrs { get; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, Dictionary<string, string>> CssProps { get; } = new Dictionary<string, Dictionary<string, string>>(StringComparer.InvariantCultureIgnoreCase);
		public List<(string type, string position)> TabStops = new List<(string type, string position)>();

		private OdtInfo()
		{
		}

		public OdtInfo(string innerText, OdtInfo parentNode = null)
		{
			if (string.IsNullOrWhiteSpace(innerText))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(innerText));
			}

			OdtNode = "xtext";
			HtmlNode = innerText;
			ParentNode = parentNode;

			if (ParentNode != null)
			{
				PreviousSibling = parentNode.ChildNodes.LastOrDefault();
				parentNode.ChildNodes.Add(this);
			}
		}

		public OdtInfo(string odtTagName, string htmlTagName, string odtClassName, OdtInfo parentNode = null)
		{
			if (string.IsNullOrWhiteSpace(odtTagName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(odtTagName));
			}

			if (string.IsNullOrWhiteSpace(htmlTagName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(htmlTagName));
			}

			OdtNode = odtTagName;
			HtmlNode = htmlTagName;
			ParentNode = parentNode;

			if (string.IsNullOrWhiteSpace(odtClassName))
			{
				ClassName = $"{HtmlNode}_{Guid.NewGuid().ToString("N")}";
			}
			else
			{
				ClassName = NormalizeClassName(odtClassName);
			}

			AddCssAttrValue(this, "class", ClassName);

			if (ParentNode != null)
			{
				PreviousSibling = parentNode.ChildNodes.LastOrDefault();
				parentNode.ChildNodes.Add(this);
			}
		}

		public static void AddCssPropertyValue(OdtInfo odtInfo, string propName, string propValue, PseudoClass pseudoClass = PseudoClass.None)
		{
			if (odtInfo == null)
			{
				throw new ArgumentNullException(nameof(odtInfo));
			}

			if (string.IsNullOrWhiteSpace(propName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(propName));
			}

			if (string.IsNullOrWhiteSpace(propValue))
			{
				return;
			}

			string className = null;

			if (pseudoClass != PseudoClass.None)
			{
				className = $"{odtInfo.ClassName}:{pseudoClass}";
			}
			else
			{
				className = odtInfo.ClassName;
			}

			odtInfo.CssProps.TryGetValue(className, out Dictionary<string, string> cssProperties);

			if (cssProperties == null)
			{
				cssProperties = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
				odtInfo.CssProps.Add(className, cssProperties);
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

		public static void AddTabStop(OdtInfo odtInfo, XElement tabStop)
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

			odtInfo.TabStops.Add((typeAttrVal, positionAttrVal));
		}

		public static IEnumerable<string> GetAttrValuesOrDefault(OdtInfo odtInfo, string attrName)
		{
			if (odtInfo.CssAttrs.TryGetValue(attrName, out List<string> values))
			{
				return values;
			}
			else
			{
				return default(IEnumerable<string>);
			}
		}

		public static string RenderCss(OdtInfo odtRootNode)
		{
			if (odtRootNode == null)
			{
				throw new ArgumentNullException(nameof(odtRootNode));
			}

			var builder = new StringBuilder(16384);
			CssRenderingWalker(odtRootNode.ChildNodes, builder);
			return builder.ToString();
		}

		public static void CssRenderingWalker(IEnumerable<OdtInfo> odtNodes, StringBuilder builder)
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

		private static void RenderCssStyle(OdtInfo odtInfo, StringBuilder builder)
		{
			if (odtInfo.CssProps.Count == 0)
			{
				return;
			}

			foreach (var cssProp in odtInfo.CssProps)
			{
				builder
				.Append(Environment.NewLine)
				.Append(".")
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

		public static void AddCssAttrValue(OdtInfo odtInfo, string attrName, string attrVal)
		{
			if (odtInfo == null)
				throw new ArgumentNullException(nameof(odtInfo));

			if (string.IsNullOrWhiteSpace(attrName))
				throw new ArgumentNullException(nameof(attrName));

			if (string.IsNullOrWhiteSpace(attrVal))
				return;

			if (attrName.Equals("class", StringComparison.InvariantCultureIgnoreCase)
				&& odtInfo.CssAttrs.ContainsKey(attrName))
			{
				return;
			}

			if (odtInfo.CssAttrs.ContainsKey(attrName))
			{
				if (!odtInfo.CssAttrs[attrName].Contains(attrVal))
				{
					odtInfo.CssAttrs[attrName].Add(attrVal);
				}
			}
			else
			{
				odtInfo.CssAttrs[attrName] = new List<string> { attrVal };
			}
		}

		public static void AddOdtAttrValue(OdtInfo odtInfo, string attrName, string attrVal)
		{
			if (string.IsNullOrWhiteSpace(attrName))
			{
				throw new ArgumentNullException(nameof(attrName));
			}

			if (string.IsNullOrWhiteSpace(attrVal))
			{
				return;
			}

			if (odtInfo.OdtAttrs.ContainsKey(attrName))
			{
				odtInfo.OdtAttrs[attrName] = attrVal;
			}
			else
			{
				odtInfo.OdtAttrs.Add(attrName, attrVal);
			}
		}

		public static string RenderHtml(OdtInfo odtInfo)
		{
			if (odtInfo == null)
			{
				throw new ArgumentNullException(nameof(odtInfo));
			}

			var htmlBuilder = new StringBuilder(8192);
			HtmlNodeWalker(odtInfo, htmlBuilder);
			return htmlBuilder.ToString();
		}

		private static void HtmlNodeWalker(OdtInfo odtInfo, StringBuilder builder)
		{
			if (odtInfo.ChildNodes.Count > 0)
			{
				builder.Append("<")
					.Append(odtInfo.HtmlNode)
					.Append(RenderAtributes(odtInfo))
					.Append(">");

				foreach (var child in odtInfo.ChildNodes)
				{
					HtmlNodeWalker(child, builder);
				}

				builder.Append("</")
					.Append(odtInfo.HtmlNode)
					.Append(">");
			}
			else if (odtInfo.OdtNode.Equals("xtext"))
			{
				builder.Append(odtInfo.HtmlNode);
			}
			else
			{
				builder.Append("<")
					.Append(odtInfo.HtmlNode)
					.Append(RenderAtributes(odtInfo))
					.Append("/>");
			}
		}

		private static string RenderAtributes(OdtInfo odtInfo)
		{
			var builder = new StringBuilder(96);

			foreach (var attr in odtInfo.CssAttrs)
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
