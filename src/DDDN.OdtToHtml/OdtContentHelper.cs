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

		public static void OdtTextNodeChildsWalker(OdtContext odtContext, IEnumerable<XNode> childXNodes, OdtHtmlInfo parentOdtHtmlInfo)
		{
			foreach (var childNode in childXNodes)
			{
				if (!HandleTextNode(childNode, parentOdtHtmlInfo)
					&& !HandleNonBreakingSpaceElement(childNode, parentOdtHtmlInfo))
				{
					HandleDocumentBodyNode(odtContext, childNode, parentOdtHtmlInfo);
				}
			}
		}

		public static void HandleDocumentBodyNode(OdtContext odtContext, XNode xNode, OdtHtmlInfo parentOdtHtmlInfo)
		{
			if (!(xNode is XElement element))
			{
				return;
			}

			var trans = OdtTrans.TagToTag.Find(p => p.OdtName.Equals(element.Name.LocalName, StrCompICIC));

			if (trans == default(OdtTransTagToTag))
			{
				OdtTextNodeChildsWalker(odtContext, element.Nodes(), parentOdtHtmlInfo);
				return;
			}

			var odtInfo = new OdtHtmlInfo(element, trans, parentOdtHtmlInfo);

			OdtStyleHelper.GetOdtStylesProperties(odtContext, odtInfo);

			HandleEmptyParagraph(element, odtInfo);
			HandleTabElement(element, odtInfo, parentOdtHtmlInfo, odtContext.ConvertSettings.DefaultTabSize);
			HandleImageElement(odtContext, element, odtInfo);
			HandleListItemElement(odtContext, element, odtInfo);
			HandleListItemNonlistChildElement(odtContext, odtInfo);

			OdtTextNodeChildsWalker(odtContext, element.Nodes(), odtInfo);
		}

		public static void HandleListItemElement(OdtContext odtContext, XElement xElement, OdtHtmlInfo odtHtmlInfo)
		{
			if (!xElement.Name.Equals(XName.Get("list-item", OdtXmlNs.Text))
				|| !OdtListHelper.TryGetListLevelInfo(odtContext, odtHtmlInfo.ListInfo, out OdtListLevel odtListLevel))
			{
				return;
			}

			TryAddCssPropertyValue(odtHtmlInfo, "text-indent", odtListLevel.PosFirstLineTextIndent, ClassKind.PseudoBefore);
			TryAddCssPropertyValue(odtHtmlInfo, "margin-right", "1em", ClassKind.PseudoBefore);
			TryAddCssPropertyValue(odtHtmlInfo, "float", "left", ClassKind.PseudoBefore);
			TryAddCssPropertyValue(odtHtmlInfo, "clear", "both", ClassKind.Own);

			var fontName = "";

			if (!String.IsNullOrWhiteSpace(odtListLevel.StyleFontName))
			{
				fontName = odtListLevel.StyleFontName;
			}
			else if (!String.IsNullOrWhiteSpace(odtListLevel.TextFontName))
			{
				fontName = odtListLevel.TextFontName;
			}

			TryAddCssPropertyValue(odtHtmlInfo, "font-family", fontName, ClassKind.PseudoBefore);
			odtContext.UsedFontFamilies.Remove(fontName);
			odtContext.UsedFontFamilies.Add(fontName);

			if (xElement.Elements().FirstOrDefault()?.Name.LocalName.Equals("list", StrCompICIC) == false)
			{
				if (odtListLevel.KindOfList == OdtListLevel.ListKind.Bullet)
				{
					AddListContent(odtHtmlInfo, odtListLevel.KindOfList, odtListLevel.DisplayLevels, odtListLevel.BulletChar, odtListLevel.NumPrefix, odtListLevel.NumSuffix);
				}
				else if (odtListLevel.KindOfList == OdtListLevel.ListKind.Number)
				{
					OdtListHelper.TryGetListItemIndex(odtHtmlInfo, out int listItemIndex);
					var numberLevelContent = OdtListHelper.GetNumberLevelContent(listItemIndex, OdtListLevel.IsKindOfNumber(odtListLevel));
					AddListContent(odtHtmlInfo, odtListLevel.KindOfList, odtListLevel.DisplayLevels, numberLevelContent, odtListLevel.NumPrefix, string.IsNullOrEmpty(odtListLevel.NumSuffix) ? "." : odtListLevel.NumSuffix);
				}

				TryAddCssPropertyValue(odtHtmlInfo, "content", $"\"{GetListContent(odtHtmlInfo) ?? "-"}\"", ClassKind.PseudoBefore);
			}
		}

		public static void HandleListItemNonlistChildElement(OdtContext context, OdtHtmlInfo htmlInfo)
		{
			if (!htmlInfo.ParentNode.OdtTag.Equals("list-item", StrCompICIC)
				|| htmlInfo.OdtTag.Equals("list", StrCompICIC)
				|| !OdtListHelper.TryGetListLevelInfo(context, htmlInfo.ListInfo, out OdtListLevel listLevelInfo))
			{
				return;
			}

			TryAddCssPropertyValue(htmlInfo, "margin-left", listLevelInfo.PosMarginLeft, ClassKind.Own);
		}

		public static bool HandleNonBreakingSpaceElement(XNode xNode, OdtHtmlInfo parentHtmlInfo)
		{
			var innerText = new StringBuilder(32);

			if (!(xNode is XElement element)
				|| !element.Name.Equals(XName.Get("s", OdtXmlNs.Text)))
			{
				return false;
			}

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

			var odtInfo = new OdtHtmlText(innerText.ToString(), xNode, parentHtmlInfo);

			return true;
		}

		public static void HandleEmptyParagraph(XElement xElement, OdtHtmlInfo odtHtmlInfo)
		{
			if (xElement == null
				|| odtHtmlInfo == null
				|| !xElement.Name.Equals(XName.Get("p", OdtXmlNs.Text)))
			{
				return;
			}

			if (xElement.Nodes()?.Any() == false)
			{
				new OdtHtmlText("&nbsp;", new XText("&nbsp;"), odtHtmlInfo);
			}
		}

		public static void HandleTabElement(XElement xElement, OdtHtmlInfo odtHtmlInfo, OdtHtmlInfo parentOdtHtmlInfo, string defaultTabSize)
		{
			if (xElement == null
				|| odtHtmlInfo == null
				|| parentOdtHtmlInfo == null
				|| String.IsNullOrWhiteSpace(defaultTabSize)
				|| !xElement.Name.Equals(XName.Get("tab", OdtXmlNs.Text)))
			{
				return;
			}

			var tabLevel = parentOdtHtmlInfo.ChildNodes.OfType<OdtHtmlInfo>().Count(p => p.OdtTag.Equals("tab", StrCompICIC));
			var lastTabStopValue = parentOdtHtmlInfo.TabStops.ElementAtOrDefault(tabLevel - 1);
			var tabStopValue = parentOdtHtmlInfo.TabStops.ElementAtOrDefault(tabLevel);

			if (tabStopValue.Equals((null, null)))
			{
				TryAddCssPropertyValue(odtHtmlInfo, "margin-left", defaultTabSize, ClassKind.Own);
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

				TryAddCssPropertyValue(odtHtmlInfo, $"margin-{tabStopValue.type}", value, ClassKind.Own);
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

			var odtInfo = new OdtHtmlText(((XText)xNode).Value, xNode, parentOdtHtmlInfo);

			return true;
		}

		public static void HandleImageElement(OdtContext odtContext, XElement xElement, OdtHtmlInfo odtHtmlInfo)
		{
			if (!xElement.Name.Equals(XName.Get("image", OdtXmlNs.Draw)))
			{
				return;
			}

			var frameStyleName = GetOdtElementAttrValOrNull(xElement.Parent, "style-name", OdtXmlNs.Draw);
			var frameStyleElement = OdtStyleHelper.FindStyleElementByNameAttr(frameStyleName, "style", odtContext.OdtStyles);
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
			TryAddCssPropertyValue(odtHtmlInfo, "max-width", maxWidth, ClassKind.Own);
			TryAddCssPropertyValue(odtHtmlInfo, "max-width", maxWidth, ClassKind.Own);

			string[] horizontalPosVal = { "left", "from-left", "right", "from-right" };

			if (horizontalPosVal.Contains(horizontalPos))
			{
				TryAddCssPropertyValue(odtHtmlInfo, "float", horizontalPos.Replace("from-", ""), ClassKind.Own);
			}
		}

		public static void HandleTabStopElement(XElement xElement, OdtHtmlInfo odtHtmlInfo)
		{
			if (xElement.Name.Equals(XName.Get("tab-stop", OdtXmlNs.Style)))
			{
				AddTabStop(odtHtmlInfo, xElement);
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
