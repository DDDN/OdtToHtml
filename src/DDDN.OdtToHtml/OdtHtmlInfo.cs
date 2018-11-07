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
using DDDN.OdtToHtml.Exceptions;
using DDDN.OdtToHtml.Transformation;
using static DDDN.OdtToHtml.OdtListLevel;

namespace DDDN.OdtToHtml
{
	public class OdtHtmlInfo : IOdtHtmlNode
	{
		public const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;

		public class OdtListInfo
		{
			public OdtHtmlInfo RootListInfo { get; set; }
			public int ListLevel { get; set; }
			public string ListItemContent { get; set; }
			public string ListItemContentPrefix { get; set; }
			public string ListItemContentSuffix { get; set; }
			public int DisplayLevels { get; set; }
			public ListKind ListKind { get; set; }
		}

		public enum ClassKind
		{
			Odt,
			Own,
			PseudoBefore,
			PseudoAfter
		}

		public string HtmlTag { get; }
		public string OdtTag { get; }
		public string OdtStyleName { get; }
		public string OwnCssClassName { get; }
		public OdtListInfo ListInfo { get; } = new OdtListInfo();
		public OdtHtmlInfo ParentNode { get; }
		public IOdtHtmlNode PreviousSibling { get; }
		public List<IOdtHtmlNode> ChildNodes { get; } = new List<IOdtHtmlNode>();
		public Dictionary<string, string> OdtAttrs { get; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, List<string>> HtmlAttrs { get; } = new Dictionary<string, List<string>>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<ClassKind, Dictionary<string, string>> CssProps { get; } = new Dictionary<ClassKind, Dictionary<string, string>>();

		private OdtHtmlInfo()
		{
		}

		public OdtHtmlInfo(XElement odtElement, OdtTransTagToTag odtTagToHtml, OdtHtmlInfo parentHtmlNode = null)
		{
			ParentNode = parentHtmlNode;
			OdtTag = odtElement?.Name.LocalName;
			HtmlTag = String.IsNullOrWhiteSpace(odtTagToHtml?.HtmlName) ? OdtTag : odtTagToHtml.HtmlName;
			OdtStyleName = odtElement?.Attributes().FirstOrDefault(p => p.Name.LocalName.Equals("style-name", StrCompICIC))?.Value;
			OwnCssClassName = $"{OdtTag}_{Guid.NewGuid().ToString("N")}";
			PreviousSibling = ParentNode?.ChildNodes.LastOrDefault();
			ParentNode?.ChildNodes.Add(this);

			TryCopyElementAttributes(this, odtElement);
			SetListInfo(this, parentHtmlNode);
		}

		private static void SetListInfo(OdtHtmlInfo odtHtmlInfo, OdtHtmlInfo parentHtmlNode)
		{
			if (odtHtmlInfo.OdtTag.Equals("list", StrCompICIC)
				&& parentHtmlNode?.OdtTag.Equals("list-item", StrCompICIC) == true)
			{
				odtHtmlInfo.ListInfo.RootListInfo = parentHtmlNode.ListInfo.RootListInfo;
				odtHtmlInfo.ListInfo.ListLevel = parentHtmlNode.ListInfo.ListLevel + 1;
			}
			else if (odtHtmlInfo.OdtTag.Equals("list", StrCompICIC)
				&& parentHtmlNode?.OdtTag.Equals("list-item", StrCompICIC) == false)
			{
				odtHtmlInfo.ListInfo.ListLevel = 1;
				odtHtmlInfo.ListInfo.RootListInfo = odtHtmlInfo;
			}
			else
			{
				odtHtmlInfo.ListInfo.RootListInfo = parentHtmlNode?.ListInfo.RootListInfo;
				odtHtmlInfo.ListInfo.ListLevel = parentHtmlNode?.ListInfo.ListLevel ?? 0;
			}
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

		public static void AddListContent(OdtHtmlInfo odtInfo, ListKind listKind, int displayLevels, string content, string prefix, string suffix)
		{
			odtInfo.ListInfo.ListItemContent = content;
			odtInfo.ListInfo.ListItemContentPrefix = prefix;
			odtInfo.ListInfo.ListItemContentSuffix = suffix;
			odtInfo.ListInfo.ListKind = listKind;
			odtInfo.ListInfo.DisplayLevels = displayLevels;
		}

		public static string GetListContent(in OdtHtmlInfo odtInfo)
		{
			var content = "";
			var parent = odtInfo;
			var displayLevel = parent.ListInfo?.DisplayLevels;

			do
			{
				if (parent.OdtTag.Equals("list-item"))
				{
					content = parent.ListInfo.ListItemContentPrefix + parent.ListInfo.ListItemContent + parent.ListInfo.ListItemContentSuffix + content;
					displayLevel--;
				}

				parent = parent.ParentNode;

			} while (displayLevel > 0 && parent != null);

			return content;
		}

		public static bool TryAddCssPropertyValue(OdtHtmlInfo odtInfo, string propName, string propValue, ClassKind classKind)
		{
			if (odtInfo == null
				|| string.IsNullOrWhiteSpace(propName)
				|| string.IsNullOrWhiteSpace(propValue))
			{
				return false;
			}

			if (classKind == ClassKind.Odt && string.IsNullOrWhiteSpace(odtInfo.OdtStyleName))
			{
				classKind = ClassKind.Own;
			}

			odtInfo.CssProps.TryGetValue(classKind, out Dictionary<string, string> cssProperties);

			if (cssProperties == null)
			{
				cssProperties = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
				odtInfo.CssProps.Add(classKind, cssProperties);
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
				if (childNode is OdtHtmlInfo htmlInfo)
				{
					Dictionary<string, string> listItemCssProps = null;
					string listItemFontFamilyCssPropValue = "";
					Dictionary<string, string> firstChildCssProps = null;
					string firstChildFontFamilyCssPropValue = "";

					if (htmlInfo.HtmlTag.Equals("li", StrCompICIC)
						&& htmlInfo.CssProps?.TryGetValue(ClassKind.PseudoBefore, out listItemCssProps) == true
						&& (listItemCssProps?.TryGetValue("font-family", out listItemFontFamilyCssPropValue) == false
							|| (listItemCssProps?.TryGetValue("font-family", out listItemFontFamilyCssPropValue) == true
								&& string.IsNullOrWhiteSpace(listItemFontFamilyCssPropValue)))
						&& htmlInfo.ChildNodes?.OfType<OdtHtmlInfo>().FirstOrDefault()?.CssProps?.TryGetValue(ClassKind.Odt, out firstChildCssProps) == true
						&& firstChildCssProps?.TryGetValue("font-family", out firstChildFontFamilyCssPropValue) == true
						&& !string.IsNullOrWhiteSpace(firstChildFontFamilyCssPropValue))
					{
						TryAddCssPropertyValue(htmlInfo, "font-family", firstChildFontFamilyCssPropValue, ClassKind.Odt);
					}

					RenderCssStyle(childNode, builder);

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
					var odtClassName = NormalizeClassName(htmlInfo.OdtStyleName);
					var ownClassName = NormalizeClassName(htmlInfo.OwnCssClassName);
					var className = "";

					switch (cssProp.Key)
					{
						case ClassKind.Odt:
							className = odtClassName ?? ownClassName;
							break;
						case ClassKind.Own:
							className = ownClassName;
							break;
						case ClassKind.PseudoBefore:
							className = $"{odtClassName ?? ownClassName}:before";
							break;
						case ClassKind.PseudoAfter:
							className = $"{odtClassName ?? ownClassName}:after";
							break;
						default:
							throw new UnknownCssClassKind("Unknown CSS class kind.");
					}

					RenderCssStyle(builder, className, ".", cssProp.Value);
				}
			}
		}

		private static void RenderCssStyle(StringBuilder builder, string styleName, string styleNamePrefix, Dictionary<string, string> styleProperties)
		{
			builder
				.Append(Environment.NewLine)
				.Append(styleNamePrefix)
				.Append(styleName)
				.Append(" {")
				.Append(Environment.NewLine);
			RenderCssStyleProperties(styleProperties, builder);
			builder.Append(" }");
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

			if (odtInfo.CssProps.ContainsKey(ClassKind.Odt)
				&& !string.IsNullOrWhiteSpace(odtInfo.OdtStyleName))
			{
				TryAddHtmlAttrValue(odtInfo, "class", NormalizeClassName(odtInfo.OdtStyleName));
			}

			if (odtInfo.CssProps.ContainsKey(ClassKind.Own))
			{
				TryAddHtmlAttrValue(odtInfo, "class", NormalizeClassName(odtInfo.OwnCssClassName));
			}

			foreach (var attr in odtInfo.HtmlAttrs)
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

		private static string NormalizeClassName(string name)
		{
			return name?.Trim().Replace(".", "_").Replace("-", "_");
		}
	}
}
