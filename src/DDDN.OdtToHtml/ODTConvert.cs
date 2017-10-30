/*
DDDN.OdtToHtml.OdtConvert
Copyright(C) 2017 Lukasz Jaskiewicz(lukasz @jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace DDDN.OdtToHtml
{
	public class OdtConvert : IOdtConvert
	{
		private OdtContext Ctx;

		private void Init(IOdtFile odtFile, OdtConvertSettings convertSettings)
		{
			var embedContent = odtFile.GetZipArchiveEntries();

			var contentXDoc = OdtFile.GetZipArchiveEntryAsXDocument(embedContent.FirstOrDefault(p => p.ContentFullName.Equals("content.xml", StringComparison.InvariantCultureIgnoreCase))?.Data);
			var stylesXDoc = OdtFile.GetZipArchiveEntryAsXDocument(embedContent.FirstOrDefault(p => p.ContentFullName.Equals("styles.xml", StringComparison.InvariantCultureIgnoreCase))?.Data);

			var documentNodes = contentXDoc.Root
						 .Elements(XName.Get("body", OdtXmlNamespaces.Office))
						 .Elements(XName.Get("text", OdtXmlNamespaces.Office))
						 .Nodes();

			var odtStyles = contentXDoc.Root.Descendants(XName.Get("style", OdtXmlNamespaces.Style));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("style", OdtXmlNamespaces.Style)));
			odtStyles = odtStyles.Concat(contentXDoc.Root.Descendants(XName.Get("default-style", OdtXmlNamespaces.Style)));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("default-style", OdtXmlNamespaces.Style)));
			odtStyles = odtStyles.Concat(contentXDoc.Root.Descendants(XName.Get("page-layout", OdtXmlNamespaces.Style)));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("page-layout", OdtXmlNamespaces.Style)));
			odtStyles = odtStyles.Concat(contentXDoc.Root.Descendants(XName.Get("master-page", OdtXmlNamespaces.Style)));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("master-page", OdtXmlNamespaces.Style)));

			var (pageWidth, pageHeight) = GetPageInfo(odtStyles);

			Ctx = new OdtContext
			{
				DocumentNodes = documentNodes,
				OdtStyles = odtStyles,
				ConvertSettings = convertSettings,
				PageWidth = pageWidth,
				PageHeight = pageHeight,
				EmbedContent = embedContent
			};
		}

		public OdtConvertedData Convert(IOdtFile odtFile, OdtConvertSettings convertSettings)
		{
			if (odtFile == null)
			{
				throw new ArgumentNullException(nameof(odtFile));
			}

			if (convertSettings == null)
			{
				throw new ArgumentNullException(nameof(convertSettings));
			}

			Init(odtFile, convertSettings);

			var htmlTreeRootTag = GetHtmlTree(Ctx);
			var html = OdtNode.RenderHtml(htmlTreeRootTag);
			var css = OdtNode.RenderCss(htmlTreeRootTag);
			var firstHeader = GetFirstHeaderHtml(htmlTreeRootTag);
			var firstParagraph = GetFirstParagraphHtml(htmlTreeRootTag);
			var embedContent = Ctx.EmbedContent.Where(p => !string.IsNullOrWhiteSpace(p.Link));

			return new OdtConvertedData
			{
				Css = css,
				Html = html,
				DocumentFirstHeader = firstHeader,
				DocumentFirstParagraph = firstParagraph,
				EmbedContent = embedContent,
				PageWidth = Ctx.PageWidth,
				PageHeight = Ctx.PageHeight,
			};
		}

		private string GetEmbedContentLink(OdtContext ctx, string odtAttrLink)
		{
			if (odtAttrLink == null)
			{
				return null;
			}

			var content = ctx.EmbedContent
				.FirstOrDefault(p => p.ContentFullName.Equals(odtAttrLink, StringComparison.InvariantCultureIgnoreCase));

			if (!string.IsNullOrWhiteSpace(content.Link))
			{
				return content.Link;
			}

			var name = $"{content.Id}_{odtAttrLink.Replace('/', '_')}";
			var link = name;

			if (!string.IsNullOrWhiteSpace(ctx.ConvertSettings.LinkUrlPrefix))
			{
				if (ctx.ConvertSettings.LinkUrlPrefix.EndsWith("/", StringComparison.InvariantCultureIgnoreCase))
				{
					link = $"{ctx.ConvertSettings.LinkUrlPrefix}{link}";
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

		private (string pageWidth, string pageHeight) GetPageInfo(IEnumerable<XElement> odtStyles)
		{
			var masterStyle = odtStyles
				.FirstOrDefault(p =>
					p.Name.LocalName.Equals("master-page", StringComparison.InvariantCultureIgnoreCase));
			var pageLayoutStyleName = GetAttrValOrNull(masterStyle, "page-layout-name", OdtXmlNamespaces.Style);
			var pageLayoutStyleProperties = FindStyleByAttrName(pageLayoutStyleName, "page-layout", odtStyles)
				?.Element(XName.Get("page-layout-properties", OdtXmlNamespaces.Style));
			var pageWidth = pageLayoutStyleProperties?.Attribute(XName.Get("page-width", OdtXmlNamespaces.XslFoCompatible))?.Value;
			var pageHeight = pageLayoutStyleProperties?.Attribute(XName.Get("page-height", OdtXmlNamespaces.XslFoCompatible))?.Value;
			return (pageWidth, pageHeight);
		}

		private string GetFirstHeaderHtml(OdtNode htmlTreeRootTag)
		{
			var headerNode = htmlTreeRootTag.ChildNodes
				.FirstOrDefault(p => p.OdtTag.Equals("h", StringComparison.InvariantCultureIgnoreCase));

			if (headerNode == default(OdtNode))
			{
				return null;
			}

			return OdtNode.RenderHtml(headerNode);
		}

		private string GetFirstParagraphHtml(OdtNode htmlTreeRootTag)
		{
			var headerNode = htmlTreeRootTag.ChildNodes
				.FirstOrDefault(p => p.OdtTag.Equals("p", StringComparison.InvariantCultureIgnoreCase));

			if (headerNode == default(OdtNode))
			{
				return null;
			}

			return OdtNode.RenderHtml(headerNode);
		}

		private OdtNode GetHtmlTree(OdtContext ctx)
		{
			var rootNode = CreateHtmlRootNode(ctx.ConvertSettings);
			OdtBodyNodesWalker(ctx, ctx.DocumentNodes, rootNode);
			return rootNode;
		}

		private static OdtNode CreateHtmlRootNode(OdtConvertSettings convertSettings)
		{
			var rootNode = new OdtNode("text", convertSettings.RootElementTagName, null);

			if (!string.IsNullOrWhiteSpace(convertSettings.RootElementId))
			{
				OdtNode.AddAttrValue(rootNode, "id", convertSettings.RootElementId);
			}

			if (!string.IsNullOrWhiteSpace(convertSettings.RootElementClassNames))
			{
				OdtNode.AddAttrValue(rootNode, "class", convertSettings.RootElementClassNames);
			}

			return rootNode;
		}

		private void OdtBodyNodesWalker(OdtContext ctx, IEnumerable<XNode> childDocumentBodyNodes, OdtNode parentNode)
		{
			foreach (var childNode in childDocumentBodyNodes)
			{
				HandleTextNode(childNode, parentNode);
				HandleDocumentBodyNode(ctx, childNode, parentNode);
			}
		}

		private void HandleDocumentBodyNode(OdtContext ctx, XNode documentBodyNode, OdtNode parentNode)
		{
			if (documentBodyNode.NodeType != XmlNodeType.Element)
			{
				return;
			}

			HandleNonBreakingSpace(documentBodyNode as XElement, parentNode);
			HandleTabs(documentBodyNode as XElement, parentNode);
			HandleFrameElement(ctx, documentBodyNode as XElement, parentNode);
			HandleOtherDocumentElement(ctx, documentBodyNode as XElement, parentNode);
		}

		private void HandleFrameElement(OdtContext ctx, XElement frameEle, OdtNode parentNode)
		{
			if (!frameEle.Name.Equals(XName.Get("frame", OdtXmlNamespaces.Draw)))
			{
				return;
			}

			var imgEle = frameEle.Element(XName.Get("image", OdtXmlNamespaces.Draw));

			if (imgEle == null)
			{
				return;
			}

			var imgNode = new OdtNode("frame", "img", "", parentNode);
			OdtNode.EnsureClassName(imgNode);

			var maxWidth = frameEle.Attribute(XName.Get("width", OdtXmlNamespaces.SvgCompatible))?.Value;
			var maxHeight = frameEle.Attribute(XName.Get("height", OdtXmlNamespaces.SvgCompatible))?.Value;
			var hrefAttrVal = imgEle.Attribute(XName.Get("href", OdtXmlNamespaces.XLink))?.Value;
			hrefAttrVal = GetEmbedContentLink(ctx, hrefAttrVal);

			OdtNode.AddAttrValue(imgNode, "src", hrefAttrVal);
			OdtNode.AddCssPropertyValue(imgNode, "max-width", maxWidth);
			OdtNode.AddCssPropertyValue(imgNode, "max-height", maxHeight);
			OdtNode.AddCssPropertyValue(imgNode, "width", "100%");
			OdtNode.AddCssPropertyValue(imgNode, "height", "auto");
		}

		private void HandleOtherDocumentElement(OdtContext ctx, XElement documentBodyElement, OdtNode parentNode)
		{
			if (!OdtTrans.TagToTag.TryGetValue(documentBodyElement.Name.LocalName, out string htmlTag))
			{
				return;
			}

			var odtClassName = documentBodyElement
				.Attributes()
				.FirstOrDefault(p =>
					p.Name.LocalName.Equals("style-name", StringComparison.InvariantCultureIgnoreCase))?.Value;

			var childHtmlNode = new OdtNode(documentBodyElement.Name.LocalName, htmlTag, odtClassName, parentNode);

			CopyElementAttributes(ctx, documentBodyElement, childHtmlNode);
			GetCssStyle(ctx, childHtmlNode);
			OdtBodyNodesWalker(ctx, documentBodyElement.Nodes(), childHtmlNode);
		}

		private void GetCssStyle(OdtContext ctx, OdtNode odtNode)
		{
			if (string.IsNullOrWhiteSpace(odtNode.OdtElementClassName))
			{
				return;
			}

			var styleElement = FindStyleByAttrName(odtNode.OdtElementClassName, "style", ctx.OdtStyles);

			var parentStyleName = GetAttrValOrNull(styleElement, "parent-style-name", OdtXmlNamespaces.Style);
			var defaultStyleFamilyName = GetAttrValOrNull(styleElement, "family", OdtXmlNamespaces.Style);

			var parentStyleElement = FindStyleByAttrName(parentStyleName, "style", ctx.OdtStyles);
			var defaultStyle = FindDefaultStyle(defaultStyleFamilyName, ctx.OdtStyles);

			TransformOdtStyleElements(defaultStyle?.Elements(), odtNode);
			TransformOdtStyleElements(parentStyleElement?.Elements(), odtNode);
			TransformOdtStyleElements(styleElement?.Elements(), odtNode);
		}

		private void TransformOdtStyleElements(IEnumerable<XElement> odtStyleElements, OdtNode odtNode)
		{
			if (odtStyleElements == null)
			{
				return;
			}

			foreach (var odtStyleElement in odtStyleElements)
			{
				HandleTabStopElement(odtStyleElement, odtNode);
				HandleStyleTrasformation(odtStyleElement, odtNode);
				TransformOdtStyleElements(odtStyleElement.Elements(), odtNode);
			}
		}

		private void HandleStyleTrasformation(XElement odtStyleElement, OdtNode odtNode)
		{
			foreach (var attr in odtStyleElement.Attributes())
			{
				var trans = OdtTrans.StyleToStyle
				.Find(p =>
					p.OdtAttrName.Equals(attr.Name.LocalName, StringComparison.InvariantCultureIgnoreCase)
					&& p.StyleTypes.Contains(odtStyleElement.Name.LocalName, StringComparer.InvariantCultureIgnoreCase));

				if (trans == null)
				{
					continue;
				}

				var attrVal = attr.Value;

				if (trans.ValueToValue?.ContainsKey(attr.Value) == true)
				{
					attrVal = trans.ValueToValue[attr.Value];
				}

				OdtNode.AddCssPropertyValue(odtNode, trans.CssPropName, attrVal);
			}
		}

		private void HandleTabStopElement(XElement odtStyleElement, OdtNode odtNode)
		{
			if (odtStyleElement.Name.Equals(XName.Get("tab-stop", OdtXmlNamespaces.Style)))
			{
				OdtNode.AddTabStop(odtNode, odtStyleElement);
			}
		}

		private XElement FindStyleByAttrName(string attrName, string styleLocalName, IEnumerable<XElement> odtStyles)
		{
			return odtStyles
				.FirstOrDefault(p =>
				p.Name.LocalName.Equals(styleLocalName, StringComparison.InvariantCultureIgnoreCase)
				&& p.Attribute(XName.Get("name", OdtXmlNamespaces.Style)).Value
					.Equals(attrName, StringComparison.InvariantCultureIgnoreCase));
		}

		private XElement FindDefaultStyle(string family, IEnumerable<XElement> odtStyles)
		{
			return odtStyles
				.FirstOrDefault(p =>
					p.Name.LocalName.Equals("default-style", StringComparison.InvariantCultureIgnoreCase)
					&& p.Attribute(XName.Get("family", OdtXmlNamespaces.Style)).Value
						.Equals(family, StringComparison.InvariantCultureIgnoreCase));
		}

		private static string GetAttrValOrNull(XElement odElement, string attrName, string attrNamespace)
		{
			return odElement?.Attribute(XName.Get(attrName, attrNamespace))?.Value;
		}

		private void CopyElementAttributes(OdtContext ctx, XElement odElement, OdtNode htmlElement)
		{
			foreach (var attr in odElement.Attributes())
			{
				var styleAndAttrLocalName = $"{odElement.Name.LocalName}.{attr.Name.LocalName}";

				if (OdtTrans.AttrNameToAttrName.TryGetValue(styleAndAttrLocalName, out string htmlAttrName))
				{
					OdtNode.AddAttrValue(htmlElement, htmlAttrName, attr.Value);
				}
			}
		}

		private static void HandleNonBreakingSpace(XElement element, OdtNode odtNode)
		{
			if (!element.Name.Equals(XName.Get("s", OdtXmlNamespaces.Text)))
			{
				return;
			}

			var spacesValue = element.Attribute(XName.Get("c", OdtXmlNamespaces.Text))?.Value;
			int.TryParse(spacesValue, out int spacesCount);

			if (spacesCount == 0)
			{
				spacesCount++;
			}

			var childNode = new OdtNode("span", "span", null, odtNode);

			for (int i = 0; i < spacesCount; i++)
			{
				childNode.InnerText += ("&nbsp;");
			}
		}

		private static void HandleTabs(XElement element, OdtNode odtParentNode)
		{
			if (!element.Name.Equals(XName.Get("tab", OdtXmlNamespaces.Text)))
			{
				return;
			}

			var tabLevel = odtParentNode.ChildNodes.Count(p => p.OdtTag.Equals("tab", StringComparison.InvariantCultureIgnoreCase));
			var lastTabStopValue = odtParentNode.TabStops.ElementAtOrDefault(tabLevel - 1);
			var tabStopValue = odtParentNode.TabStops.ElementAtOrDefault(tabLevel);
			var tabNodeOdtClassName = $"{odtParentNode.OdtElementClassName}tab{tabLevel}";
			var tabNode = new OdtNode("tab", "span", tabNodeOdtClassName, odtParentNode);
			OdtNode.AddAttrValue(tabNode, "class", tabNodeOdtClassName);

			if (tabStopValue.Equals((null, null)))
			{
				OdtNode.AddCssPropertyValue(tabNode, "margin-left", "2rem");
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

				OdtNode.AddCssPropertyValue(tabNode, $"margin-{tabStopValue.type}", value);
			}
		}

		private static void GetLevel(ref int level, string elementTag, string elementNameSpace, OdtNode odtNode)
		{
			if (odtNode != null)
			{
				if (odtNode.OdtTag.Equals("tab", StringComparison.InvariantCultureIgnoreCase))
				{
					level++;
				}

				GetLevel(ref level, elementTag, elementNameSpace, odtNode.ParentNode);
			}
		}

		private static void HandleTextNode(XNode textNode, OdtNode parentNode)
		{
			if (textNode.NodeType != XmlNodeType.Text)
			{
				return;
			}

			var childNode = new OdtNode("span", "span", null, parentNode)
			{
				InnerText = ((XText)textNode).Value
			};
		}

		private static bool IsCssNumberValue(string value)
		{
			return Regex.IsMatch(value, @"^[+-]?[0-9]+.?([0-9]+)?(px|em|ex|%|in|cm|mm|pt|pc)$");
		}

		private static bool IsCssColorValue(string value)
		{
			return Regex.IsMatch(value, @"^(#[0-9a-f]{3}|#(?:[0-9a-f]{2}){2,4}|(rgb|hsl)a?\((-?\d+%?[,\s]+){2,3}\s*[\d\.]+%?\))$");
		}
	}
}