/*
DDDN.OdtToHtml.OdtHtmlInfo
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
	public class OdtHtmlInfo : IOdtHtmlNode
	{
		public const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;

		public enum ClassKind
		{
			Odt,
			Own,
			PseudoBefore,
			PseudoAfter
		}

		public string HtmlTag { get; }
		public string OdtTag { get; }
		public string OdtClass { get; }
		public string OwnClass { get; }
		public OdtHtmlInfo ParentNode { get; }
		public IOdtHtmlNode PreviousSibling { get; }
		public List<IOdtHtmlNode> ChildNodes { get; } = new List<IOdtHtmlNode>();
		public Dictionary<string, string> OdtAttrs { get; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, List<string>> HtmlAttrs { get; } = new Dictionary<string, List<string>>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, Dictionary<string, string>> CssProps { get; } = new Dictionary<string, Dictionary<string, string>>(StringComparer.InvariantCultureIgnoreCase);
		public List<(string type, string position)> TabStops = new List<(string type, string position)>();

		private OdtHtmlInfo()
		{
		}

		public OdtHtmlInfo(XElement odtElement, OdtTagToHtml odtTagToHtml, OdtHtmlInfo parentHtmlNode = null)
		{
			ParentNode = parentHtmlNode;
			OdtTag = odtElement?.Name.LocalName;
			HtmlTag = String.IsNullOrWhiteSpace(odtTagToHtml?.HtmlName) ? OdtTag : odtTagToHtml.HtmlName;
			OdtClass = odtElement?.Attributes().FirstOrDefault(p => p.Name.LocalName.Equals("style-name", StrCompICIC))?.Value;
			OwnClass = $"{OdtTag}_{Guid.NewGuid().ToString("N")}";
			PreviousSibling = ParentNode?.ChildNodes.LastOrDefault();
			ParentNode?.ChildNodes.Add(this);

			TryCopyElementAttributes(this, odtElement);
			TryApplyDefaultCssStyleProperties(this, odtTagToHtml?.DefaultCssProperties);
			TryAddHtmlAttrValue(this, "class", OdtClass);
			TryAddHtmlAttrValue(this, "class", OwnClass);
		}

		public static bool TryCopyElementAttributes(OdtHtmlInfo odtInfo, XElement odtElement)
		{
			if (odtElement == null
				|| odtInfo == null)
			{
				return false;
			}

			foreach (var attr in odtElement.Attributes())
			{
				var styleAndAttrLocalName = $"{odtElement.Name.LocalName}.{attr.Name.LocalName}";

				if (OdtTrans.OdtAttrToHtmlAttr.TryGetValue(styleAndAttrLocalName, out string htmlAttrName))
				{
					TryAddHtmlAttrValue(odtInfo, htmlAttrName, attr.Value);
				}

				if (OdtTrans.OdtAttr.Any(p => p.Equals(attr.Name.LocalName, StrCompICIC)))
				{
					odtInfo.OdtAttrs.Add(attr.Name.LocalName, attr.Value);
				}
			}

			return true;
		}

		private static bool TryApplyDefaultCssStyleProperties(OdtHtmlInfo odtInfo, Dictionary<string, string> defaultProperties)
		{
			if (odtInfo == null
				|| defaultProperties == null)
			{
				return false;
			}

			foreach (var prop in defaultProperties)
			{
				if (!String.IsNullOrWhiteSpace(odtInfo.OdtClass))
				{
					TryAddCssPropertyValue(odtInfo, prop.Key, prop.Value, ClassKind.Odt);
				}
				else
				{
					TryAddCssPropertyValue(odtInfo, prop.Key, prop.Value, ClassKind.Own);
				}
			}

			return true;
		}

		public static bool TryAddCssPropertyValue(OdtHtmlInfo odtInfo, string propName, string propValue, ClassKind classKind)
		{
			if (odtInfo == null
				|| string.IsNullOrWhiteSpace(propName)
				|| string.IsNullOrWhiteSpace(propValue))
			{
				return false;
			}

			string className = null;

			switch (classKind)
			{
				case ClassKind.Odt:
					className = odtInfo.OdtClass ?? odtInfo.OwnClass;
					break;
				case ClassKind.Own:
					className = odtInfo.OwnClass;
					break;
				case ClassKind.PseudoBefore:
					className = $"{odtInfo.OwnClass}:before";
					break;
				case ClassKind.PseudoAfter:
					className = $"{odtInfo.OwnClass}:after";
					break;
				default:
					return false;
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

			return true;
		}

		public static void AddTabStop(OdtHtmlInfo odtInfo, XElement tabStopElement)
		{
			var typeAttrVal = tabStopElement.Attribute(XName.Get("type", OdtXmlNs.Style))?.Value;
			var positionAttrVal = tabStopElement.Attribute(XName.Get("position", OdtXmlNs.Style))?.Value;

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

		public static bool TryAddHtmlAttrValue(OdtHtmlInfo htmlTagNode, string attrName, string attrVal)
		{
			if (htmlTagNode == null
				|| string.IsNullOrWhiteSpace(attrName)
				|| string.IsNullOrWhiteSpace(attrVal))
			{
				return false;
			}

			if (htmlTagNode.HtmlAttrs.ContainsKey(attrName))
			{
				if (!htmlTagNode.HtmlAttrs[attrName].Contains(attrVal))
				{
					htmlTagNode.HtmlAttrs[attrName].Add(attrVal);
				}
			}
			else
			{
				htmlTagNode.HtmlAttrs[attrName] = new List<string> { attrVal };
			}

			return true;
		}

		public static bool TryGetHtmlAttrValues(OdtHtmlInfo odtInfo, string attrName, out IEnumerable<string> attrValues)
		{
			odtInfo.HtmlAttrs.TryGetValue(attrName, out List<string> values);
			attrValues = values;
			return values != null;
		}

		public static string RenderCss(OdtHtmlInfo odtRootNode)
		{
			if (odtRootNode == null)
			{
				return null;
			}

			var builder = new StringBuilder(16384);
			CssStylesWalker(odtRootNode.ChildNodes, builder);
			return builder.ToString();
		}

		public static void CssStylesWalker(IEnumerable<IOdtHtmlNode> odtNodes, StringBuilder builder)
		{
			foreach (var childNode in odtNodes)
			{
				RenderCssStyle(childNode, builder);

				if (childNode is OdtHtmlInfo htmlInfo)
				{
					CssStylesWalker(htmlInfo.ChildNodes, builder);
				}
			}
		}

		private static void RenderCssStyle(IOdtHtmlNode htmlNode, StringBuilder builder)
		{
			if (htmlNode is OdtHtmlInfo htmlInfo
				&& htmlInfo.CssProps.Count != 0)
			{
				foreach (var cssProp in htmlInfo.CssProps.Where(p => p.Value.Values.Count > 0))
				{
					builder
						.Append(Environment.NewLine)
						.Append(".")
						.Append(NormalizeClassName(cssProp.Key))
						.Append(" {")
						.Append(Environment.NewLine);
					RenderCssStyleProperties(cssProp.Value, builder);
					builder.Append(" }");
				}
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

		public static string RenderHtml(OdtHtmlInfo odtHtmlInfo)
		{
			if (odtHtmlInfo == null)
			{
				return null;
			}

			var htmlBuilder = new StringBuilder(8192);
			HtmlNodesWalker(odtHtmlInfo, htmlBuilder);
			return htmlBuilder.ToString();
		}

		private static void HtmlNodesWalker(IOdtHtmlNode htmlNode, StringBuilder builder)
		{
			if (htmlNode is OdtHtmlInfo htmlInfo)
			{
				var htmlTag = ComplementHtmlTagName(htmlInfo);

				builder.Append("<")
					.Append(htmlTag)
					.Append(RenderHtmlNodeAttributes(htmlInfo));

				if (htmlInfo.ChildNodes.Count > 0)
				{
					builder.Append(">");

					foreach (var child in htmlInfo.ChildNodes)
					{
						HtmlNodesWalker(child, builder);
					}

					builder.Append("</")
						.Append(htmlTag)
						.Append(">");
				}
				else
				{
					builder.Append("/>");
				}
			}
			else if (htmlNode is OdtHtmlText odtHtmlText)
			{
				builder.Append(odtHtmlText.InnerText);
			}
		}

		private static string ComplementHtmlTagName(OdtHtmlInfo htmlInfo)
		{
			if (htmlInfo.OdtTag.Equals("h", StrCompICIC)
				&& htmlInfo.OdtAttrs.TryGetValue("outline-level", out string outlineLevel)
				&& double.TryParse(outlineLevel, out double outlineLevelNo))
			{
				if (outlineLevelNo < 7)
				{
					return htmlInfo.HtmlTag + outlineLevel;
				}
				else
				{
					return "h6";
				}
			}

			return htmlInfo.HtmlTag;
		}

		private static string RenderHtmlNodeAttributes(OdtHtmlInfo odtInfo)
		{
			var builder = new StringBuilder(96);

			foreach (var attr in odtInfo.HtmlAttrs)
			{
				builder
					.Append(" ")
					.Append(attr.Key)
					.Append("=\"");
				bool firstName = true;

				foreach (var val in attr.Value)
				{
					var normalizedValue = val;

					if (attr.Key.Equals("class", StrCompICIC))
					{
						normalizedValue = NormalizeClassName(val);

						if ((!odtInfo.CssProps.TryGetValue(val, out Dictionary<string, string> props)
								|| props.Count == 0)
							&& odtInfo.ParentNode != null)
						{
							continue;
						}
					}

					if (firstName)
					{
						builder.Append(normalizedValue);
						firstName = false;
					}
					else
					{
						builder
							.Append(" ")
							.Append(normalizedValue);
					}
				}

				builder.Append("\"");
			}

			return builder.ToString();
		}

		private static string NormalizeClassName(string name)
		{
			return name?.Trim().Replace(".", "_").Replace("-", "_");
		}
	}
}
