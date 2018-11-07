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
using DDDN.OdtToHtml.Conversion;
using DDDN.OdtToHtml.Transformation;
using static DDDN.OdtToHtml.OdtList;

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
		public Dictionary<string, string> OwnCssProps { get; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, string> BeforeCssProps { get; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);

		private OdtHtmlInfo()
		{
		}

		public static void AddOwnCssProps(OdtHtmlInfo htmlInfo, string name, string value)
		{
			if (htmlInfo.OwnCssProps.ContainsKey(name))
			{
				htmlInfo.OwnCssProps[name] = value;
			}
			else
			{
				htmlInfo.OwnCssProps.Add(name, value);
			}
		}

		public static void AddBeforeCssProps(OdtHtmlInfo htmlInfo, string name, string value)
		{
			if (htmlInfo.BeforeCssProps.ContainsKey(name))
			{
				htmlInfo.BeforeCssProps[name] = value;
			}
			else
			{
				htmlInfo.BeforeCssProps.Add(name, value);
			}
		}

		public OdtHtmlInfo(XElement odtElement, OdtTransTagToTag odtTagToHtml, OdtHtmlInfo parentHtmlNode = null)
		{
			ParentNode = parentHtmlNode;
			OdtTag = odtElement?.Name.LocalName;
			HtmlTag = string.IsNullOrWhiteSpace(odtTagToHtml?.HtmlName) ? OdtTag : odtTagToHtml.HtmlName;
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

		public static string RenderCss(OdtContext ctx, OdtHtmlInfo odtRootNode)
		{
			if (odtRootNode == null)
			{
				return null;
			}

			var builder = new StringBuilder(16384);
			RenderTagToTagCss(builder);
			RenderOdtStyles(ctx, builder);
			RenderElementCss(odtRootNode.ChildNodes, builder);
			return builder.ToString();
		}

		private static void RenderOdtStyles(OdtContext ctx, StringBuilder builder)
		{
			foreach (var style in ctx.UsedStyles.Values.Where(p => p.Props?.Any() == true))
			{
				RenderCssStyle(builder, OdtCssHelper.NormalizeClassName(style.Name), ".", style.Props);
			}
		}

		private static void RenderTagToTagCss(StringBuilder builder)
		{
			foreach (var tagToTag in OdtTrans.TagToTag.Where(p => p.DefaultCssProperties?.Any() == true))
			{
				RenderCssStyle(builder, "t2t_" + tagToTag.OdtName, ".", tagToTag.DefaultCssProperties);
			}
		}

		public static void RenderElementCss(IEnumerable<IOdtHtmlNode> odtNodes, StringBuilder builder)
		{
			foreach (var childNode in odtNodes)
			{
				if (childNode is OdtHtmlInfo htmlInfo)
				{
					string listItemFontFamilyCssPropValue = "";
					string firstChildFontFamilyCssPropValue = "";

					if (htmlInfo.HtmlTag.Equals("li", StrCompICIC)
						&& (htmlInfo.BeforeCssProps?.TryGetValue("font-family", out listItemFontFamilyCssPropValue) == false
							|| (htmlInfo.BeforeCssProps?.TryGetValue("font-family", out listItemFontFamilyCssPropValue) == true
								&& string.IsNullOrWhiteSpace(listItemFontFamilyCssPropValue)))
						&& htmlInfo.ChildNodes?.OfType<OdtHtmlInfo>().FirstOrDefault()?.OwnCssProps?.TryGetValue("font-family", out firstChildFontFamilyCssPropValue) == true
						&& !string.IsNullOrWhiteSpace(firstChildFontFamilyCssPropValue))
					{
						OdtHtmlInfo.AddOwnCssProps(htmlInfo, "font-family", firstChildFontFamilyCssPropValue);
					}

					if (htmlInfo.OwnCssProps.Count > 0)
					{
						RenderCssStyle(builder, htmlInfo.OwnCssClassName, ".", htmlInfo.OwnCssProps);
					}

					if (htmlInfo.BeforeCssProps.Count > 0)
					{
						RenderCssStyle(builder, htmlInfo.OwnCssClassName + ":before", ".", htmlInfo.BeforeCssProps);
					}

					RenderElementCss(htmlInfo.ChildNodes, builder);
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

			var transTagToTag = OdtTrans.TagToTag.Find(p => p.OdtName.Equals(odtInfo.OdtTag, StrCompICIC));

			if (transTagToTag?.DefaultCssProperties?.Any() == true)
			{
				TryAddHtmlAttrValue(odtInfo, "class", "t2t_" + odtInfo.OdtTag);
			}

			if (!string.IsNullOrWhiteSpace(odtInfo.OdtStyleName))
			{
				TryAddHtmlAttrValue(odtInfo, "class", OdtCssHelper.NormalizeClassName(odtInfo.OdtStyleName));
			}

			if (odtInfo.OwnCssProps?.Any() == true
				|| odtInfo.BeforeCssProps?.Any() == true)
			{
				TryAddHtmlAttrValue(odtInfo, "class", odtInfo.OwnCssClassName);
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
	}
}
