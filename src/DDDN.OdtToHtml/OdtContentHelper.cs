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

namespace DDDN.OdtToHtml
{
	public static class OdtContentHelper
	{
		public const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;

		public static void DocumentBodyNodesWalker(OdtContext ctx, IEnumerable<XNode> childDocumentBodyNodes, OdtInfo parentOdtInfo)
		{
			foreach (var childNode in childDocumentBodyNodes)
			{
				if (!HandleTextNode(childNode, parentOdtInfo)
					&& !HandleNonBreakingSpaceElement(childNode, parentOdtInfo))
				{
					HandleDocumentBodyNode(ctx, childNode, parentOdtInfo);
				}
			}
		}

		public static void HandleDocumentBodyNode(OdtContext ctx, XNode node, OdtInfo parentOdtInfo)
		{
			var element = node as XElement;

			if (element == null)
			{
				return;
			}

			var tag = OdtTrans.TagToTag.Find(p => p.OdtName.Equals(element.Name.LocalName, StrCompICIC));

			if (tag == default(OdtTagToHtml))
			{
				DocumentBodyNodesWalker(ctx, element.Nodes(), parentOdtInfo);
				return;
			}

			var odtClassName = element.Attributes()
				.FirstOrDefault(p => p.Name.LocalName.Equals("style-name", StrCompICIC))?.Value;

			var odtInfo = new OdtInfo(element.Name.LocalName, tag.HtmlName, odtClassName, parentOdtInfo);

			OdtStyleHelper.ApplyDefaultCssStyleProperties(odtInfo, tag.DefaultProperty);
			CopyElementAttributes(element, odtInfo);
			OdtStyleHelper.GetStylesProperties(ctx, odtInfo);

			HandleTabElement(element, odtInfo, parentOdtInfo, ctx.ConvertSettings.DefaultTabSize);
			HandleImageElement(ctx, element, odtInfo);
			HandleListItemElement(ctx, element, odtInfo);
			HandleElementAfterListElement(ctx, element, odtInfo);
			HandleElementAfterListItemElement(ctx, odtInfo);

			DocumentBodyNodesWalker(ctx, element.Nodes(), odtInfo);
		}

		public static void HandleListItemElement(OdtContext ctx, XElement element, OdtInfo odtInfo)
		{
			if (!element.Name.Equals(XName.Get("list-item", OdtXmlNs.Text)))
			{
				return;
			}

			var listItemIndex = OdtListHelper.GetListItemIndex(odtInfo);
			var listLevel = OdtListHelper.GetListLevel(odtInfo.ParentNode);
			var listLevelInfo = OdtListHelper.GetListLevelInfo(ctx, OdtListHelper.GetListClassName(odtInfo), listLevel);

			if (listLevelInfo == null)
			{
				return;
			}

			OdtInfo.AddCssPropertyValue(odtInfo, "margin-left", listLevelInfo.PosSpaceBefore);
			//OdtInfo.AddCssPropertyValue(odtInfo, "display", "inline-block");
			//OdtInfo.AddCssPropertyValue(odtInfo, "margin-right", listLevelInfo.PosFirstLineMarginLeft, OdtInfo.PseudoClass.Before);
			OdtInfo.AddCssPropertyValue(odtInfo, "min-width", listLevelInfo.PosLabelWidth, OdtInfo.PseudoClass.Before);
			//OdtInfo.AddCssPropertyValue(odtInfo, "display", "inline-block", OdtInfo.PseudoClass.Before);
			OdtInfo.AddCssPropertyValue(odtInfo, "float", "left", OdtInfo.PseudoClass.Before);

			if (!String.IsNullOrWhiteSpace(listLevelInfo.TextFontName))
			{
				OdtInfo.AddCssPropertyValue(odtInfo, "font-family", listLevelInfo.TextFontName, OdtInfo.PseudoClass.Before);
				ctx.UsedFontFamilies.Remove(listLevelInfo.TextFontName);
				ctx.UsedFontFamilies.Add(listLevelInfo.TextFontName);
			}

			var nextElement = element.Elements().FirstOrDefault();

			if (nextElement != null)
			{
				var nextElementLocalName = nextElement.Name.LocalName;

				if (!nextElementLocalName.Equals("list", StrCompICIC))
				{
					if (listLevelInfo.KindOfList == OdtListLevel.ListKind.Bullet)
					{
						OdtInfo.AddCssPropertyValue(odtInfo, "content", $"\"{listLevelInfo.BulletChar}\"", OdtInfo.PseudoClass.Before);
					}
					else if (listLevelInfo.KindOfList == OdtListLevel.ListKind.Number)
					{
						var numberLevelContent = OdtListHelper.GetNumberLevelContent(listItemIndex, OdtListLevel.IsKindOfNumber(listLevelInfo));
						OdtInfo.AddCssPropertyValue(odtInfo, "content", $"\"{listLevelInfo.NumPrefix}{numberLevelContent}{listLevelInfo.NumSuffix}\"", OdtInfo.PseudoClass.Before);
					}
				}
			}
		}

		public static void HandleElementAfterListItemElement(OdtContext ctx, OdtInfo odtInfo)
		{
			if (!odtInfo.ParentNode.OdtNode.Equals("list-item", StrCompICIC)
				|| odtInfo.OdtNode.Equals("list", StrCompICIC))
			{
				return;
			}

			var listItemNode = odtInfo.ParentNode;
			var listLevel = OdtListHelper.GetListLevel(listItemNode);
			var listLevelInfo = OdtListHelper.GetListLevelInfo(ctx, OdtListHelper.GetListClassName(odtInfo), listLevel);

			if (listLevelInfo != null)
			{
				if (odtInfo.PreviousSibling == null)
				{
					OdtInfo.AddCssPropertyValue(odtInfo, "margin-left", listLevelInfo.PosFirstLineIndent);
				}
				else
				{
					OdtInfo.AddCssPropertyValue(odtInfo, "margin-left", listLevelInfo.PosTextIndent);
				}

				//if (odtInfo.PreviousSibling == null
				//	&& String.IsNullOrWhiteSpace(listLevelInfo.TextFontName)
				//	&& odtInfo.CssProps.TryGetValue($".{odtInfo.ClassName}", out Dictionary<string, string> style)
				//	&& style.TryGetValue("font-family", out string fontName)
				//	&& !String.IsNullOrWhiteSpace(fontName))
				//{
				//	OdtInfo.AddCssPropertyValue(listItemNode, "font-family", fontName, OdtInfo.PseudoClass.Before);
				//	ctx.UsedFontFamilies.Remove(listLevelInfo.TextFontName);
				//	ctx.UsedFontFamilies.Add(listLevelInfo.TextFontName);
				//}
			}
		}

		public static void HandleElementAfterListElement(OdtContext ctx, XElement element, OdtInfo odtInfo)
		{
			if (odtInfo.PreviousSibling == null
					|| !odtInfo.PreviousSibling.OdtNode.Equals("list", StrCompICIC)
					|| element.Name.Equals(XName.Get("list-item", OdtXmlNs.Text)))
			{
				return;
			}

			var listLevel = OdtListHelper.GetListLevel(odtInfo.PreviousSibling);
			var listLevelInfo = OdtListHelper.GetListLevelInfo(ctx, OdtListHelper.GetListClassName(odtInfo), listLevel);

			if (listLevelInfo == null)
			{
				return;
			}

			OdtInfo.AddCssPropertyValue(odtInfo, "margin-left", listLevelInfo.PosTextIndent);
		}

		public static bool HandleNonBreakingSpaceElement(XNode node, OdtInfo parentOdtInfo)
		{
			var element = node as XElement;
			var innerText = new StringBuilder(16);

			if (element == null
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

			var odtInfo = new OdtInfo(innerText.ToString(), parentOdtInfo);
			return true;
		}

		public static void HandleTabElement(XElement element, OdtInfo odtInfo, OdtInfo parentOdtInfo, string defaultTabSize)
		{
			if (!element.Name.Equals(XName.Get("tab", OdtXmlNs.Text)))
			{
				return;
			}

			var tabLevel = parentOdtInfo.ChildNodes.Count(p => p.OdtNode.Equals("tab", StrCompICIC));
			var lastTabStopValue = parentOdtInfo.TabStops.ElementAtOrDefault(tabLevel - 1);
			var tabStopValue = parentOdtInfo.TabStops.ElementAtOrDefault(tabLevel);

			if (tabStopValue.Equals((null, null)))
			{
				OdtInfo.AddCssPropertyValue(odtInfo, "margin-left", defaultTabSize);
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

				OdtInfo.AddCssPropertyValue(odtInfo, $"margin-{tabStopValue.type}", value);
			}
		}

		public static bool HandleTextNode(XNode node, OdtInfo parentOdtInfo)
		{
			if (node.NodeType != XmlNodeType.Text
				|| (node.Parent.NodeType == XmlNodeType.Element
					&& !OdtTrans.TextNodeParent.Contains(node.Parent.Name.LocalName, StringComparer.InvariantCultureIgnoreCase)))
			{
				return false;
			}

			var odtInfo = new OdtInfo(((XText)node).Value, parentOdtInfo);
			return true;
		}

		public static void HandleImageElement(OdtContext ctx, XElement element, OdtInfo odtInfo)
		{
			if (!element.Name.Equals(XName.Get("image", OdtXmlNs.Draw)))
			{
				return;
			}

			var hrefAttrVal = element.Attribute(XName.Get("href", OdtXmlNs.XLink))?.Value;

			if (hrefAttrVal == null)
			{
				OdtInfo.AddCssAttrValue(odtInfo, "alt", "Image href attribute not found.");
			}
			else
			{
				var contentLink = GetEmbedContentLink(ctx, hrefAttrVal);

				if (contentLink != null)
				{
					OdtInfo.AddCssAttrValue(odtInfo, "src", contentLink);
					OdtInfo.AddCssAttrValue(odtInfo, "alt", contentLink);
				}
				else
				{
					OdtInfo.AddCssAttrValue(odtInfo, "alt", "Embedded content not found.");
				}
			}

			var maxWidth = element.Parent.Attribute(XName.Get("width", OdtXmlNs.SvgCompatible))?.Value;
			var maxHeight = element.Parent.Attribute(XName.Get("height", OdtXmlNs.SvgCompatible))?.Value;
			OdtInfo.AddCssPropertyValue(odtInfo, "max-width", maxWidth);
			OdtInfo.AddCssPropertyValue(odtInfo, "max-height", maxHeight);
		}

		public static void HandleTabStopElement(XElement styleElement, OdtInfo odtInfo)
		{
			if (styleElement.Name.Equals(XName.Get("tab-stop", OdtXmlNs.Style)))
			{
				OdtInfo.AddTabStop(odtInfo, styleElement);
			}
		}

		public static string GetEmbedContentLink(OdtContext ctx, string odtAttrLink)
		{
			if (string.IsNullOrWhiteSpace(odtAttrLink))
			{
				return null;
			}

			var content = ctx.EmbedContent
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

			if (!string.IsNullOrWhiteSpace(ctx.ConvertSettings.LinkUrlPrefix))
			{
				if (ctx.ConvertSettings.LinkUrlPrefix.EndsWith("/", StrCompICIC))
				{
					link = ctx.ConvertSettings.LinkUrlPrefix + link;
				}
				else
				{
					link = $"{ctx.ConvertSettings.LinkUrlPrefix}/{link}";
				}
			}

			content.LinkName = name;
			content.Link = link;

			return link;
		}

		public static void CopyElementAttributes(XElement element, OdtInfo odtInfo)
		{
			if (element == null)
			{
				throw new ArgumentNullException(nameof(element));
			}

			if (odtInfo == null)
			{
				throw new ArgumentNullException(nameof(odtInfo));
			}

			foreach (var attr in element.Attributes())
			{
				OdtInfo.AddOdtAttrValue(odtInfo, attr.Name.LocalName, attr.Value);

				var styleAndAttrLocalName = $"{element.Name.LocalName}.{attr.Name.LocalName}";

				if (OdtTrans.AttrNameToAttrName.TryGetValue(styleAndAttrLocalName, out string htmlAttrName))
				{
					OdtInfo.AddCssAttrValue(odtInfo, htmlAttrName, attr.Value);
				}
			}
		}

		public static string GetOdtElementAttrValOrNull(XElement element, string attrName, string attrNamespace)
		{
			return element?.Attribute(XName.Get(attrName, attrNamespace))?.Value;
		}
	}
}
