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
using System.Threading;
using System.Xml.Linq;
using DDDN.OdtToHtml.Conversion;
using DDDN.OdtToHtml.Transformation;
using static DDDN.OdtToHtml.OdtListStyle;

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

		private static int NodeCounter;
		public int NodeNo { get; }
		public string HtmlTag { get; }
		public string OdtTag { get; }
		public string OdtStyleName { get; }
		public OdtListInfo ListInfo { get; } = new OdtListInfo();
		public OdtHtmlInfo ParentNode { get; }
		public IOdtHtmlNode PreviousSibling { get; }
		public List<IOdtHtmlNode> ChildNodes { get; } = new List<IOdtHtmlNode>();
		public Dictionary<string, string> OdtAttrs { get; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, List<string>> HtmlAttrs { get; } = new Dictionary<string, List<string>>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, string> OwnCssProps { get; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, string> BeforeCssProps { get; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);

		private OdtHtmlInfo(OdtHtmlInfo parentHtmlInfo)
		{
			if (parentHtmlInfo == null)
			{
				NodeCounter = 0;
				NodeNo = 0;
			}
			else
			{
				NodeNo = Interlocked.Increment(ref NodeCounter);
			}
		}

		public OdtHtmlInfo(XElement docElement, string htmlTag, OdtHtmlInfo parentHtmlInfo) : this(parentHtmlInfo)
		{
			ParentNode = parentHtmlInfo;
			ParentNode?.ChildNodes.Add(this);
			PreviousSibling = ParentNode?.ChildNodes.LastOrDefault();

			OdtTag = docElement?.Name.LocalName;
			HtmlTag = htmlTag;
			OdtStyleName = docElement?.Attributes().FirstOrDefault(p => p.Name.LocalName.Equals("style-name", StrCompICIC))?.Value;

			TryCopyElementAttributes(this, docElement);
			SetListInfo(this, ParentNode);
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

		private static void SetListInfo(OdtHtmlInfo htmlInfo, OdtHtmlInfo parentHtmlInfo)
		{
			if (htmlInfo.OdtTag.Equals("list", StrCompICIC)
				&& parentHtmlInfo?.OdtTag.Equals("list-item", StrCompICIC) == true)
			{
				htmlInfo.ListInfo.RootListInfo = parentHtmlInfo.ListInfo.RootListInfo;
				htmlInfo.ListInfo.ListLevel = parentHtmlInfo.ListInfo.ListLevel + 1;
			}
			else if (htmlInfo.OdtTag.Equals("list", StrCompICIC)
				&& parentHtmlInfo?.OdtTag.Equals("list-item", StrCompICIC) == false)
			{
				htmlInfo.ListInfo.ListLevel = 1;
				htmlInfo.ListInfo.RootListInfo = htmlInfo;
			}
			else
			{
				htmlInfo.ListInfo.RootListInfo = parentHtmlInfo?.ListInfo.RootListInfo;
				htmlInfo.ListInfo.ListLevel = parentHtmlInfo?.ListInfo.ListLevel ?? 0;
			}
		}

		public static bool TryCopyElementAttributes(OdtHtmlInfo htmlInfo, XElement docElement)
		{
			if (docElement == null
				|| htmlInfo == null)
			{
				return false;
			}

			foreach (var attr in docElement.Attributes())
			{
				var styleAndAttrLocalName = $"{docElement.Name.LocalName}.{attr.Name.LocalName}";

				if (OdtTrans.OdtAttrToHtmlAttr.TryGetValue(styleAndAttrLocalName, out string htmlAttrName))
				{
					TryAddHtmlAttrValue(htmlInfo, htmlAttrName, attr.Value);
				}

				if (OdtTrans.OdtAttr.Any(p => p.Equals(attr.Name.LocalName, StrCompICIC)))
				{
					htmlInfo.OdtAttrs.Add(attr.Name.LocalName, attr.Value);
				}
			}

			return true;
		}

		public static void AddListContent(OdtHtmlInfo htmlInfo, ListKind listKind, int displayLevels, string content, string prefix, string suffix)
		{
			htmlInfo.ListInfo.ListItemContent = content;
			htmlInfo.ListInfo.ListItemContentPrefix = prefix;
			htmlInfo.ListInfo.ListItemContentSuffix = suffix;
			htmlInfo.ListInfo.ListKind = listKind;
			htmlInfo.ListInfo.DisplayLevels = displayLevels;
		}

		public static string GetListContent(in OdtHtmlInfo htmlInfo)
		{
			var content = "";
			var parent = htmlInfo;
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

		public static bool TryAddHtmlAttrValue(OdtHtmlInfo htmlInfo, string attrName, string attrVal)
		{
			if (htmlInfo == null
				|| string.IsNullOrWhiteSpace(attrName)
				|| string.IsNullOrWhiteSpace(attrVal))
			{
				return false;
			}

			if (htmlInfo.HtmlAttrs.ContainsKey(attrName))
			{
				if (!htmlInfo.HtmlAttrs[attrName].Contains(attrVal))
				{
					htmlInfo.HtmlAttrs[attrName].Add(attrVal);
				}
			}
			else
			{
				htmlInfo.HtmlAttrs[attrName] = new List<string> { attrVal };
			}

			return true;
		}

		public static string RenderCss(OdtContext ctx, OdtHtmlInfo htmlInfo)
		{
			if (htmlInfo == null)
			{
				return null;
			}

			var builder = new StringBuilder(16384);
			RenderTagToTagCss(builder);
			RenderOdtStyles(ctx, builder);
			RenderElementCss(ctx, htmlInfo.ChildNodes, builder);
			return builder.ToString();
		}

		private static void RenderOdtStyles(OdtContext ctx, StringBuilder builder)
		{
			foreach (var style in ctx.UsedStyles.Values.Where(p => p.CssProps?.Any() == true))
			{
				RenderCssStyle(builder, OdtCssHelper.NormalizeClassName(style.Name), ".", style.CssProps);
			}
		}

		private static void RenderTagToTagCss(StringBuilder builder)
		{
			foreach (var tagToTag in OdtTrans.TagToTag.Where(p => p.DefaultCssProperties?.Any() == true))
			{
				RenderCssStyle(builder, "t2t_" + tagToTag.OdtTag, ".", tagToTag.DefaultCssProperties);
			}
		}

		public static void RenderElementCss(OdtContext ctx, IEnumerable<IOdtHtmlNode> odtNodes, StringBuilder builder)
		{
			foreach (var childNode in odtNodes)
			{
				if (childNode is OdtHtmlInfo htmlInfo)
				{
					string listItemFontFamilyCssPropValue = "";
					string firstChildFontFamilyCssPropValue = "";
					OdtStyle firstChildStyle = null;

					if (htmlInfo.HtmlTag.Equals("li", StrCompICIC))
					{
						if (htmlInfo.BeforeCssProps?.TryGetValue("font-family", out listItemFontFamilyCssPropValue) == false
							|| (htmlInfo.BeforeCssProps?.TryGetValue("font-family", out listItemFontFamilyCssPropValue) == true
								&& string.IsNullOrWhiteSpace(listItemFontFamilyCssPropValue)))
						{
							if ((htmlInfo.ChildNodes?.OfType<OdtHtmlInfo>().FirstOrDefault()?.OwnCssProps?.TryGetValue("font-family", out firstChildFontFamilyCssPropValue) == true
								|| (!string.IsNullOrWhiteSpace(htmlInfo.ChildNodes?.OfType<OdtHtmlInfo>().FirstOrDefault()?.OdtStyleName)
									&& ctx.UsedStyles.TryGetValue(htmlInfo.ChildNodes.OfType<OdtHtmlInfo>().FirstOrDefault().OdtStyleName, out firstChildStyle)
									&& firstChildStyle.CssProps.TryGetValue("font-family", out firstChildFontFamilyCssPropValue)))
									&& !string.IsNullOrWhiteSpace(firstChildFontFamilyCssPropValue))
							{
								AddBeforeCssProps(htmlInfo, "font-family", firstChildFontFamilyCssPropValue);
							}
						}
					}

					if (htmlInfo.OwnCssProps.Count > 0)
					{
						RenderCssStyle(builder, "nno-" + htmlInfo.NodeNo, ".", htmlInfo.OwnCssProps);
					}

					if (htmlInfo.BeforeCssProps.Count > 0)
					{
						RenderCssStyle(builder, "nno-" + htmlInfo.NodeNo + ":before", ".", htmlInfo.BeforeCssProps);
					}

					RenderElementCss(ctx, htmlInfo.ChildNodes, builder);
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

		public static string RenderHtml(OdtHtmlInfo htmlInfo)
		{
			if (htmlInfo == null)
			{
				return null;
			}

			var htmlBuilder = new StringBuilder(8192);
			HtmlNodesWalker(htmlInfo, htmlBuilder);
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
			else if (htmlNode is OdtHtmlText htmlText)
			{
				builder.Append(htmlText.InnerText);
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

		private static string RenderHtmlNodeAttributes(OdtHtmlInfo htmlInfo)
		{
			var builder = new StringBuilder(96);

			var transTagToTag = OdtTrans.TagToTag.Find(p => p.OdtTag.Equals(htmlInfo.OdtTag, StrCompICIC));

			if (transTagToTag?.DefaultCssProperties?.Any() == true)
			{
				TryAddHtmlAttrValue(htmlInfo, "class", "t2t_" + htmlInfo.OdtTag);
			}

			if (!string.IsNullOrWhiteSpace(htmlInfo.OdtStyleName))
			{
				TryAddHtmlAttrValue(htmlInfo, "class", OdtCssHelper.NormalizeClassName(htmlInfo.OdtStyleName));
			}

			if (htmlInfo.OwnCssProps?.Any() == true
				|| htmlInfo.BeforeCssProps?.Any() == true)
			{
				TryAddHtmlAttrValue(htmlInfo, "class", "nno-" + htmlInfo.NodeNo);
			}

			foreach (var attr in htmlInfo.HtmlAttrs)
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
