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
using static DDDN.OdtToHtml.OdtHtmlInfo;

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

			if (transTagToTag == default(OdtTransTagToTag))
			{
				return null;
			}

			var odtInfo = new OdtHtmlInfo(element, transTagToTag.HtmlTag, parentHtmlInfo);

			if (!string.IsNullOrEmpty(transTagToTag.DefaultValue)
				&& string.IsNullOrEmpty(element.Value))
			{
				new OdtHtmlText(transTagToTag.DefaultValue, odtInfo);
			}

			OdtStyle.HandleOdtStyle(ctx, odtInfo);

			HandleEmptyParagraphElement(element, odtInfo);
			HandleTabElement(element, ctx.UsedStyles, odtInfo, parentHtmlInfo, ctx.ConvertSettings.DefaultTabSize);
			HandleImageElement(ctx, element, odtInfo);
			HandleListItemElement(ctx, element, odtInfo);
			HandleListItemNonlistChildElement(ctx, odtInfo);

			return odtInfo;
		}

		public static void HandleListItemElement(OdtContext odtContext, XElement xElement, OdtHtmlInfo htmlInfo)
		{
			if (!xElement.Name.Equals(XName.Get("list-item", OdtXmlNs.Text))
				|| !OdtList.TryGetListLevelInfo(odtContext, htmlInfo.ListInfo, out OdtList odtListLevel))
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
				if (odtListLevel.KindOfList == OdtList.ListKind.Bullet)
				{
					AddListContent(htmlInfo, odtListLevel.KindOfList, odtListLevel.DisplayLevels, odtListLevel.BulletChar, odtListLevel.NumPrefix, odtListLevel.NumSuffix);
				}
				else if (odtListLevel.KindOfList == OdtList.ListKind.Number)
				{
					OdtList.TryGetListItemIndex(htmlInfo, out int listItemIndex);
					var numberLevelContent = OdtList.GetNumberLevelContent(listItemIndex, OdtList.IsKindOfNumber(odtListLevel));
					AddListContent(htmlInfo, odtListLevel.KindOfList, odtListLevel.DisplayLevels, numberLevelContent, odtListLevel.NumPrefix, string.IsNullOrEmpty(odtListLevel.NumSuffix) ? "." : odtListLevel.NumSuffix);
				}

				OdtHtmlInfo.AddBeforeCssProps(htmlInfo, "content", $"\"{GetListContent(htmlInfo) ?? "-"}\"");
			}
		}

		public static void HandleListItemNonlistChildElement(OdtContext context, OdtHtmlInfo htmlInfo)
		{
			if (!htmlInfo.ParentNode.OdtTag.Equals("list-item", StrCompICIC)
				|| htmlInfo.OdtTag.Equals("list", StrCompICIC)
				|| !OdtList.TryGetListLevelInfo(context, htmlInfo.ListInfo, out OdtList listLevelInfo))
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
			XElement xElement,
			Dictionary<string, OdtStyle> styles,
			OdtHtmlInfo odtHtmlInfo,
			OdtHtmlInfo parentOdtHtmlInfo,
			string defaultTabSize)
		{
			if (xElement == null
				|| odtHtmlInfo == null
				|| parentOdtHtmlInfo == null
				|| string.IsNullOrWhiteSpace(parentOdtHtmlInfo.OdtStyleName)
				|| string.IsNullOrWhiteSpace(defaultTabSize)
				|| !xElement.Name.Equals(XName.Get("tab", OdtXmlNs.Text)))
			{
				return;
			}
			var parentStyle = styles[parentOdtHtmlInfo.OdtStyleName];
			var tabLevel = parentOdtHtmlInfo.ChildNodes.OfType<OdtHtmlInfo>().Count(p => p.OdtTag.Equals("tab", StrCompICIC));
			var lastTabStopValue = parentStyle.TabStops.ElementAtOrDefault(tabLevel - 1);
			var tabStopValue = parentStyle.TabStops.ElementAtOrDefault(tabLevel);

			if (tabStopValue.Equals((null, null)))
			{
				OdtHtmlInfo.AddOwnCssProps(odtHtmlInfo, "margin-left", defaultTabSize);
			}
			else
			{
				var value = "";

				if (lastTabStopValue.Equals((null, null)))
				{
					value = tabStopValue.position;
				}
				else
				{
					value = $"calc({tabStopValue.position} - {lastTabStopValue.position})";
				}

				OdtHtmlInfo.AddOwnCssProps(odtHtmlInfo, $"margin-{tabStopValue.type}", value);
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

		public static void HandleImageElement(OdtContext odtContext, XElement xElement, OdtHtmlInfo odtHtmlInfo)
		{
			if (!xElement.Name.Equals(XName.Get("image", OdtXmlNs.Draw)))
			{
				return;
			}

			var frameStyleName = GetOdtElementAttrValOrNull(xElement.Parent, "style-name", OdtXmlNs.Draw);
			var frameStyleElement = OdtStyle.FindStyleElementByNameAttr(frameStyleName, "style", odtContext.OdtStyles);
			var frameStyleElementGraphicProps = frameStyleElement?.Element(XName.Get("graphic-properties", OdtXmlNs.Style));
			var horizontalPos = GetOdtElementAttrValOrNull(frameStyleElementGraphicProps, "horizontal-pos", OdtXmlNs.Style);
			var verticalPos = GetOdtElementAttrValOrNull(frameStyleElementGraphicProps, "vertical-pos", OdtXmlNs.Style);

			var hrefAttrVal = xElement.Attribute(XName.Get("href", OdtXmlNs.XLink))?.Value;

			if (hrefAttrVal == null)
			{
				TryAddHtmlAttrValue(odtHtmlInfo, "alt", "Image href attribute not found.");
			}
			else
			{
				var contentLink = GetEmbedContentLink(odtContext, hrefAttrVal);

				if (contentLink != null)
				{
					TryAddHtmlAttrValue(odtHtmlInfo, "src", contentLink);
					TryAddHtmlAttrValue(odtHtmlInfo, "alt", contentLink);
				}
				else
				{
					TryAddHtmlAttrValue(odtHtmlInfo, "alt", "Embedded content not found.");
				}
			}

			var maxWidth = xElement.Parent.Attribute(XName.Get("width", OdtXmlNs.SvgCompatible))?.Value;
			var maxHeight = xElement.Parent.Attribute(XName.Get("height", OdtXmlNs.SvgCompatible))?.Value;
			OdtHtmlInfo.AddOwnCssProps(odtHtmlInfo, "max-width", maxWidth);
			OdtHtmlInfo.AddOwnCssProps(odtHtmlInfo, "max-height", maxHeight); // TODO richtig?

			string[] horizontalPosVal = { "left", "from-left", "right", "from-right" };

			if (horizontalPosVal.Contains(horizontalPos))
			{
				OdtHtmlInfo.AddOwnCssProps(odtHtmlInfo, "float", horizontalPos.Replace("from-", ""));
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

		public static string GetOdtElementAttrValOrNull(XElement xElement, string attrName, string attrNamespace)
		{
			return xElement?.Attribute(XName.Get(attrName, attrNamespace))?.Value;
		}
	}
}
