/*
DDDN.OdtToHtml.OdtContentHelper
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
using System.Xml;
using System.Xml.Linq;
using DDDN.OdtToHtml.Conversion;
using DDDN.OdtToHtml.Transformation;

namespace DDDN.OdtToHtml
{
	public static class OdtContentHelper
	{
		public const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;

		public static void OdtNodesWalker(OdtContext ctx, IEnumerable<XNode> childXNodes, OdtHtmlInfo parentHtmlInfo)
		{
			foreach (var childNode in childXNodes)
			{
				if (!HandleTextNode(childNode, parentHtmlInfo)
					&& !HandleNonBreakingSpaceNode(childNode, parentHtmlInfo))
				{
					var childHtmlInfo = HandleElementNode(ctx, childNode, parentHtmlInfo);

					if (childNode is XElement element)
					{
						OdtNodesWalker(ctx, element.Nodes(), childHtmlInfo ?? parentHtmlInfo);
					}
				}
			}
		}

		public static OdtHtmlInfo HandleElementNode(OdtContext ctx, XNode xNode, OdtHtmlInfo parentHtmlInfo)
		{
			if (!(xNode is XElement element))
			{
				return null;
			}

			var transTagToTag = OdtTrans.TagToTag.Find(p => p.OdtTag.Equals(element.Name.LocalName, StrCompICIC));

			if (transTagToTag == null)
			{
				return null;
			}

			var odtInfo = new OdtHtmlInfo(element, transTagToTag.HtmlTag, parentHtmlInfo);

			if (!string.IsNullOrEmpty(transTagToTag.DefaultValue)
				&& string.IsNullOrEmpty(element.Value))
			{
				new OdtHtmlText(transTagToTag.DefaultValue, odtInfo);
			}

			HandleEmptyParagraphElement(element, odtInfo);
			HandleTabElement(ctx, element, odtInfo, parentHtmlInfo);
			HandleImageElement(ctx, element, odtInfo);
			HandleListItemElement(ctx, element, odtInfo);
			HandleListItemNonlistChildElement(ctx, odtInfo);
			// TODO handle all other elements?

			return odtInfo;
		}

		public static void HandleListItemElement(OdtContext odtContext, XElement xElement, OdtHtmlInfo htmlInfo)
		{
			if (!xElement.Name.Equals(XName.Get("list-item", OdtXmlNs.Text))
				|| !OdtListStyle.TryGetListLevelInfo(odtContext, htmlInfo.ListInfo, out OdtListStyle odtListLevel))
			{
				return;
			}

			OdtHtmlInfo.AddBeforeCssProps(htmlInfo, "text-indent", odtListLevel.PosFirstLineTextIndent);
			OdtHtmlInfo.AddBeforeCssProps(htmlInfo, "margin-right", "1em");
			OdtHtmlInfo.AddBeforeCssProps(htmlInfo, "float", "left");
			OdtHtmlInfo.AddOwnCssProps(htmlInfo, "clear", "both");

			var fontName = "";

			if (!string.IsNullOrWhiteSpace(odtListLevel.StyleFontName))
			{
				fontName = odtListLevel.StyleFontName;
			}
			else if (!string.IsNullOrWhiteSpace(odtListLevel.TextFontName))
			{
				fontName = odtListLevel.TextFontName;
			}

			OdtHtmlInfo.AddBeforeCssProps(htmlInfo, "font-family", fontName);
			odtContext.UsedFontFamilies.Remove(fontName);
			odtContext.UsedFontFamilies.Add(fontName);

			if (xElement.Elements().FirstOrDefault()?.Name.LocalName.Equals("list", StrCompICIC) == false)
			{
				if (odtListLevel.KindOfList == OdtListStyle.ListKind.Bullet)
				{
					OdtHtmlInfo.AddListContent(htmlInfo, odtListLevel.KindOfList, odtListLevel.DisplayLevels, odtListLevel.BulletChar, odtListLevel.NumPrefix, odtListLevel.NumSuffix);
				}
				else if (odtListLevel.KindOfList == OdtListStyle.ListKind.Number)
				{
					OdtListStyle.TryGetListItemIndex(htmlInfo, out int listItemIndex);
					var numberLevelContent = OdtListStyle.GetNumberLevelContent(listItemIndex, OdtListStyle.IsKindOfNumber(odtListLevel));
					OdtHtmlInfo.AddListContent(htmlInfo, odtListLevel.KindOfList, odtListLevel.DisplayLevels, numberLevelContent, odtListLevel.NumPrefix, string.IsNullOrEmpty(odtListLevel.NumSuffix) ? "." : odtListLevel.NumSuffix);
				}

				OdtHtmlInfo.AddBeforeCssProps(htmlInfo, "content", $"\"{OdtHtmlInfo.GetListContent(htmlInfo) ?? "-"}\"");
			}
		}

		public static void HandleListItemNonlistChildElement(OdtContext context, OdtHtmlInfo htmlInfo)
		{
			if (!htmlInfo.ParentNode.OdtTag.Equals("list-item", StrCompICIC)
				|| htmlInfo.OdtTag.Equals("list", StrCompICIC)
				|| !OdtListStyle.TryGetListLevelInfo(context, htmlInfo.ListInfo, out OdtListStyle listLevelInfo))
			{
				return;
			}

			OdtHtmlInfo.AddOwnCssProps(htmlInfo, "margin-left", listLevelInfo.PosMarginLeft);
		}

		public static bool HandleNonBreakingSpaceNode(XNode xNode, OdtHtmlInfo parentHtmlInfo)
		{
			if (!(xNode is XElement element)
				|| !element.Name.Equals(XName.Get("s", OdtXmlNs.Text)))
			{
				return false;
			}

			var innerText = new StringBuilder(32);

			var spacesValue = element.Attribute(XName.Get("c", OdtXmlNs.Text))?.Value;
			int.TryParse(spacesValue, out int spacesCount);

			if (spacesCount == 0)
			{
				spacesCount++;
			}

			for (int i = 0; i < spacesCount; i++)
			{
				innerText.Append("&nbsp;");
			}

			new OdtHtmlText(innerText.ToString(), parentHtmlInfo);

			return true;
		}

		public static void HandleEmptyParagraphElement(XElement xElement, OdtHtmlInfo odtHtmlInfo)
		{
			if (xElement == null
				|| odtHtmlInfo == null
				|| !xElement.Name.Equals(XName.Get("p", OdtXmlNs.Text)))
			{
				return;
			}

			if (xElement.Nodes()?.Any() == false)
			{
				new OdtHtmlText("&nbsp;", odtHtmlInfo);
			}
		}

		public static void HandleTabElement(
			OdtContext ctx,
			XElement element,
			OdtHtmlInfo htmlInfo,
			OdtHtmlInfo parentHtmlInfo)
		{
			if (element?.Name.Equals(XName.Get("tab", OdtXmlNs.Text)) != true
				|| htmlInfo == null
				|| string.IsNullOrWhiteSpace(parentHtmlInfo?.OdtStyleName))
			{
				return;
			}

			var parentStyle = ctx.Styles.Find(p => p.Style.Equals(parentHtmlInfo.OdtStyleName, StrCompICIC));
			var tabLevel = parentHtmlInfo.ChildNodes.OfType<OdtHtmlInfo>().Count(p => p.OdtTag.Equals("tab", StrCompICIC));
			var tabStop = parentStyle.TabStops.ElementAtOrDefault(tabLevel);
			var lastTabStop = parentStyle.TabStops.ElementAtOrDefault(tabLevel - 1);

			if (tabStop == (null, null)) // TODO checking values separately?
			{
				OdtHtmlInfo.AddOwnCssProps(htmlInfo, "margin-left", ctx.ConvertSettings.DefaultTabSize);
			}
			else
			{
				var value = "";

				if (lastTabStop == (null, null))
				{
					value = tabStop.position;
				}
				else
				{
					value = $"calc({tabStop.position} - {lastTabStop.position})";
				}

				OdtHtmlInfo.AddOwnCssProps(htmlInfo, $"margin-{tabStop.type}", value);
			}
		}

		public static bool HandleTextNode(XNode xNode, OdtHtmlInfo parentOdtHtmlInfo)
		{
			if (xNode.NodeType != XmlNodeType.Text
				|| (xNode.Parent.NodeType == XmlNodeType.Element
					&& !OdtTrans.TextNodeParent.Contains(xNode.Parent.Name.LocalName, StringComparer.InvariantCultureIgnoreCase)))
			{
				return false;
			}

			new OdtHtmlText(((XText)xNode).Value, parentOdtHtmlInfo);

			return true;
		}

		public static void HandleImageElement(OdtContext ctx, XElement xElement, OdtHtmlInfo htmlInfo)
		{
			if (!xElement.Name.Equals(XName.Get("image", OdtXmlNs.Draw)))
			{
				return;
			}

			TryGetXAttrValue(xElement.Parent, "style-name", OdtXmlNs.Draw, out string frameStyleName);
			var frameStyleElement = ctx.Styles.Find(p => p.Style.Equals(frameStyleName, StrCompICIC));
			var horizontalPos = frameStyleElement?.Attrs.Find(p => p.LocalName.Equals("horizontal-pos", StrCompICIC))?.Value;
			var verticalPos = frameStyleElement?.Attrs.Find(p => p.LocalName.Equals("vertical-pos", StrCompICIC))?.Value;
			var maxWidth = frameStyleElement?.Attrs.Find(p => p.LocalName.Equals("width", StrCompICIC))?.Value;
			var maxHeight = frameStyleElement?.Attrs.Find(p => p.LocalName.Equals("height", StrCompICIC))?.Value;

			if (TryGetXAttrValue(xElement, "href", OdtXmlNs.XLink, out string hrefAttrVal))
			{
				OdtHtmlInfo.TryAddHtmlAttrValue(htmlInfo, "alt", "Image href attribute not found.");
			}
			else
			{
				var contentLink = GetEmbedContentLink(ctx, hrefAttrVal);

				if (contentLink != null)
				{
					OdtHtmlInfo.TryAddHtmlAttrValue(htmlInfo, "src", contentLink);
					OdtHtmlInfo.TryAddHtmlAttrValue(htmlInfo, "alt", contentLink);
				}
				else
				{
					OdtHtmlInfo.TryAddHtmlAttrValue(htmlInfo, "alt", "Embedded content not found.");
				}
			}

			OdtHtmlInfo.AddOwnCssProps(htmlInfo, "max-width", maxWidth);
			OdtHtmlInfo.AddOwnCssProps(htmlInfo, "max-height", maxHeight); // TODO correct or twice time width like before?

			string[] horizontalPosVal = { "left", "from-left", "right", "from-right" };

			if (horizontalPosVal.Contains(horizontalPos))
			{
				OdtHtmlInfo.AddOwnCssProps(htmlInfo, "float", horizontalPos.Replace("from-", ""));
			}
		}

		public static string GetEmbedContentLink(OdtContext odtContext, string odtAttrLink)
		{
			if (string.IsNullOrWhiteSpace(odtAttrLink))
			{
				return null;
			}

			var content = odtContext.EmbedContent
				.FirstOrDefault(p => p.ContentFullName.Equals(odtAttrLink, StrCompICIC));

			if (content == default(OdtEmbedContent))
			{
				return null;
			}

			if (!string.IsNullOrWhiteSpace(content.Link))
			{
				return content.Link;
			}

			var name = $"{content.Id}_{odtAttrLink.Replace('/', '_')}";
			var link = name;

			if (!string.IsNullOrWhiteSpace(odtContext.ConvertSettings.LinkUrlPrefix))
			{
				if (odtContext.ConvertSettings.LinkUrlPrefix.EndsWith("/", StrCompICIC))
				{
					link = odtContext.ConvertSettings.LinkUrlPrefix + link;
				}
				else
				{
					link = $"{odtContext.ConvertSettings.LinkUrlPrefix}/{link}";
				}
			}

			content.LinkName = name;
			content.Link = link;

			return link;
		}

		public static bool TryGetXAttrValue(XElement xElement, string attrName, string attrNamespace, out string value)
		{
			return (value = xElement?.Attribute(XName.Get(attrName, attrNamespace))?.Value) != null;
		}

		public static string GetOdtElementAttrValOrNull(XElement xElement, string attrName, string attrNamespace)
		{
			return xElement?.Attribute(XName.Get(attrName, attrNamespace))?.Value;
		}
	}
}
