/*
DDDN.OdtToHtml.OdtConvert
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
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace DDDN.OdtToHtml
{
	public class OdtConvert : IOdtConvert
	{
		private OdtContext Ctx;
		private const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;

		private void Init(IOdtFile odtFile, OdtConvertSettings convertSettings)
		{
			var embedContent = odtFile.GetZipArchiveEntries();

			var contentXDoc = OdtFile.GetZipArchiveEntryAsXDocument(embedContent
				.FirstOrDefault(p => p.ContentFullName.Equals("content.xml", StrCompICIC))?.Data);
			var stylesXDoc = OdtFile.GetZipArchiveEntryAsXDocument(embedContent
				.FirstOrDefault(p => p.ContentFullName.Equals("styles.xml", StrCompICIC))?.Data);

			var documentNodes = contentXDoc.Root
						 .Elements(XName.Get("body", OdtXmlNs.Office))
						 .Elements(XName.Get("text", OdtXmlNs.Office))
						 .Nodes();

			var odtStyles = contentXDoc.Root.Descendants(XName.Get("style", OdtXmlNs.Style));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("style", OdtXmlNs.Style)));

			odtStyles = odtStyles.Concat(contentXDoc.Root.Descendants(XName.Get("default-style", OdtXmlNs.Style)));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("default-style", OdtXmlNs.Style)));

			var listStyles = contentXDoc.Root.Descendants(XName.Get("list-style", OdtXmlNs.Text));
			listStyles = listStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("list-style", OdtXmlNs.Text)));
			odtStyles.Concat(listStyles);

			odtStyles = odtStyles.Concat(contentXDoc.Root.Descendants(XName.Get("page-layout", OdtXmlNs.Style)));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("page-layout", OdtXmlNs.Style)));

			odtStyles = odtStyles.Concat(contentXDoc.Root.Descendants(XName.Get("master-page", OdtXmlNs.Style)));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("master-page", OdtXmlNs.Style)));

			var pageIngo = GetPageInfo(odtStyles);
			var odtListsLevelInfo = CreateListLevelInfos(listStyles, pageIngo);

			Ctx = new OdtContext
			{
				DocumentNodes = documentNodes,
				OdtStyles = odtStyles,
				ConvertSettings = convertSettings,
				EmbedContent = embedContent,
				PageInfo = pageIngo,
				OdtListsLevelInfo = odtListsLevelInfo
			};
		}

		private Dictionary<string, Dictionary<string, OdtListLevel>> CreateListLevelInfos(IEnumerable<XElement> listStyleELements, OdtPageInfo pageInfo)
		{
			var listStyleInfos = new Dictionary<string, Dictionary<string, OdtListLevel>>(StringComparer.InvariantCultureIgnoreCase);

			foreach (var listStyleElement in listStyleELements)
			{
				var listStyleName = GetAttrValOrNull(listStyleElement, "name", OdtXmlNs.Style);

				if (listStyleName == null)
				{
					continue;
				}

				listStyleInfos.TryGetValue(listStyleName, out Dictionary<string, OdtListLevel> styleLevelInfos);

				if (styleLevelInfos == null)
				{
					styleLevelInfos = new Dictionary<string, OdtListLevel>(StringComparer.InvariantCultureIgnoreCase);
					listStyleInfos.Add(listStyleName, styleLevelInfos);
					AddStyleListLevels(pageInfo, listStyleName, listStyleElement, styleLevelInfos);
				}
			}

			return listStyleInfos;
		}

		private void AddStyleListLevels(OdtPageInfo pageInfo, string listStyleName, XElement listStyleElement, Dictionary<string, OdtListLevel> styleLevelInfos)
		{
			OdtListLevel parentLevel = null;
			foreach (var styleLevelElement in listStyleElement.Elements())
			{
				var listLevelInfo = CreateListLevelInfo(listStyleName, styleLevelElement);
				styleLevelInfos.Add(listLevelInfo.level, listLevelInfo.listLevelInfo);
				CalculateListSpaces(listLevelInfo.listLevelInfo, parentLevel, pageInfo);
				parentLevel = listLevelInfo.listLevelInfo;
			}
		}

		private static (string level, OdtListLevel listLevelInfo) CreateListLevelInfo(string styleName, XElement levelElement)
		{
			if (string.IsNullOrWhiteSpace(styleName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(styleName));
			}

			if (levelElement == null)
			{
				throw new ArgumentNullException(nameof(levelElement));
			}

			var listLevelPropertiesElement = levelElement.Element(XName.Get("list-level-properties", OdtXmlNs.Style));
			var listLevelLabelAlignmentElement = listLevelPropertiesElement.Element(XName.Get("list-level-label-alignment", OdtXmlNs.Style));
			var textPropertiesElement = levelElement.Element(XName.Get("text-properties", OdtXmlNs.Style));

			var level = levelElement?.Attribute(XName.Get("level", OdtXmlNs.Text))?.Value;
			var bulletChar = levelElement?.Attribute(XName.Get("bullet-char", OdtXmlNs.Text))?.Value;
			var numFormat = levelElement?.Attribute(XName.Get("num-format", OdtXmlNs.Style))?.Value;
			var spaceBefore = listLevelPropertiesElement?.Attribute(XName.Get("space-before", OdtXmlNs.Text))?.Value;
			var labelMarginLeft = listLevelLabelAlignmentElement?.Attribute(XName.Get("margin-left", OdtXmlNs.XslFoCompatible))?.Value;
			var textIndent = listLevelLabelAlignmentElement?.Attribute(XName.Get("text-indent", OdtXmlNs.XslFoCompatible))?.Value;
			var fontName = textPropertiesElement?.Attribute(XName.Get("font-name", OdtXmlNs.Style))?.Value;

			return (level, new OdtListLevel(styleName, levelElement, level)
			{
				FontName = fontName,
				SpaceBefore = spaceBefore,
				MarginLeft = labelMarginLeft,
				TextIndent = textIndent,
				BulletChar = bulletChar,
				NumFormat = numFormat,
			});
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
				PageInfo = Ctx.PageInfo
			};
		}

		private string GetEmbedContentLink(OdtContext ctx, string odtAttrLink)
		{
			if (string.IsNullOrWhiteSpace(odtAttrLink))
			{
				return "empty_odt_file_content_link";
			}

			var content = ctx.EmbedContent
				.FirstOrDefault(p => p.ContentFullName.Equals(odtAttrLink, StrCompICIC));

			if (content == default(OdtEmbedContent))
			{
				return "content_not_found_in_odt_file";
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

		private OdtPageInfo GetPageInfo(IEnumerable<XElement> odtStyles)
		{
			var masterStyle = odtStyles
				.FirstOrDefault(p =>
					p.Name.LocalName.Equals("master-page", StrCompICIC));
			var pageLayoutStyleName = GetAttrValOrNull(masterStyle, "page-layout-name", OdtXmlNs.Style);
			var pageLayoutStyleProperties = FindStyleByNameAttr(pageLayoutStyleName, "page-layout", odtStyles)
				?.Element(XName.Get("page-layout-properties", OdtXmlNs.Style));

			var width = pageLayoutStyleProperties?.Attribute(XName.Get("page-width", OdtXmlNs.XslFoCompatible))?.Value;
			var height = pageLayoutStyleProperties?.Attribute(XName.Get("page-height", OdtXmlNs.XslFoCompatible))?.Value;
			var marginTop = pageLayoutStyleProperties?.Attribute(XName.Get("margin-top", OdtXmlNs.XslFoCompatible))?.Value;
			var marginLeft = pageLayoutStyleProperties?.Attribute(XName.Get("margin-left", OdtXmlNs.XslFoCompatible))?.Value;
			var marginBottom = pageLayoutStyleProperties?.Attribute(XName.Get("margin-bottom", OdtXmlNs.XslFoCompatible))?.Value;
			var marginRight = pageLayoutStyleProperties?.Attribute(XName.Get("margin-right", OdtXmlNs.XslFoCompatible))?.Value;

			var provider = new NumberFormatInfo { NumberDecimalSeparator = "." };
			double.TryParse(GetRealNumber(width), NumberStyles.Any, provider, out double widthNo);
			double.TryParse(GetRealNumber(height), NumberStyles.Any, provider, out double heightNo);
			double.TryParse(GetRealNumber(marginTop), NumberStyles.Any, provider, out double marginTopNo);
			double.TryParse(GetRealNumber(marginLeft), NumberStyles.Any, provider, out double marginLeftNo);
			double.TryParse(GetRealNumber(marginBottom), NumberStyles.Any, provider, out double marginBottomNo);
			double.TryParse(GetRealNumber(marginRight), NumberStyles.Any, provider, out double marginRightNo);

			var widthNettoNo = widthNo - (marginBottomNo + marginRightNo);
			var heightNettoNo = heightNo - (marginTopNo + marginBottomNo);

			var widthNetto = widthNettoNo.ToString(provider) + GetNumberUnit(width);
			var heightNetto = heightNettoNo.ToString(provider) + GetNumberUnit(height);

			return new OdtPageInfo
			{
				WidthBrutto = width,
				HeightBrutto = height,
				WidthNetto = widthNetto,
				HeightNetto = heightNetto,
				MarginTop = marginTop,
				MarginLeft = marginLeft,
				MarginBottom = marginBottom,
				MarginRight = marginRight,
			};
		}

		private string GetFirstHeaderHtml(OdtNode htmlTreeRootTag)
		{
			var headerNode = htmlTreeRootTag.ChildNodes
				.FirstOrDefault(p => p.OdtTag.Equals("h", StrCompICIC));

			if (headerNode == default(OdtNode))
			{
				return null;
			}

			return OdtNode.RenderHtml(headerNode);
		}

		private string GetFirstParagraphHtml(OdtNode htmlTreeRootTag)
		{
			var headerNode = htmlTreeRootTag.ChildNodes
				.FirstOrDefault(p => p.OdtTag.Equals("p", StrCompICIC));

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

			var documentBodyElement = documentBodyNode as XElement;

			var tag = OdtTrans.TagToTag
				.FirstOrDefault(p =>
					p.OdtName.Equals(documentBodyElement.Name.LocalName, StrCompICIC));

			if (tag == default(OdtTagToHtml))
			{
				OdtBodyNodesWalker(ctx, documentBodyElement.Nodes(), parentNode);
				return;
			}

			var odtClassName = documentBodyElement.Attributes()
				.FirstOrDefault(p => p.Name.LocalName.Equals("style-name", StrCompICIC))?.Value;

			var childHtmlNode = new OdtNode(documentBodyElement.Name.LocalName, tag.HtmlName, odtClassName, parentNode);

			ApplyDefaultStyleProperties(childHtmlNode, tag.DefaultProperty);
			CopyElementAttributes(documentBodyElement, childHtmlNode);
			GetCssStyleProperties(ctx, childHtmlNode);

			HandleNonBreakingSpaceElement(documentBodyElement, childHtmlNode);
			HandleTabElement(documentBodyElement, childHtmlNode, parentNode, ctx.ConvertSettings.DefaultTabSize);
			HandleImageElement(ctx, documentBodyElement, childHtmlNode);
			HandleElementAfterListElement(ctx, documentBodyElement, childHtmlNode);
			HandleListItemChildElement(ctx, documentBodyElement, childHtmlNode);
			HandleListItemElement(ctx, documentBodyElement, childHtmlNode);

			OdtBodyNodesWalker(ctx, documentBodyElement.Nodes(), childHtmlNode);
		}

		private void HandleListItemChildElement(OdtContext ctx, XElement element, OdtNode odtNode)
		{
			if (!odtNode.ParentNode.OdtTag.Equals("list-item", StrCompICIC))
			{
				return;
			}

			OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), "margin-left", "0");
		}

		private void HandleElementAfterListElement(OdtContext ctx, XElement element, OdtNode odtNode)
		{
			if (odtNode.PreviousSameHierarchyNode == null
					|| !odtNode.PreviousSameHierarchyNode.OdtTag.Equals("list", StrCompICIC)
					|| element.Name.Equals(XName.Get("list-item", OdtXmlNs.Text)))
			{
				return;
			}

			var listLevel = GetListLevel(odtNode.PreviousSameHierarchyNode);
			var listLevelInfo = ctx.OdtListsLevelInfo[GetListClassName(odtNode.PreviousSameHierarchyNode)][listLevel];
			listLevelInfo.MarginLeftPercent = GetCssValuePercentValueRelativeToPage(ctx.PageInfo, OdtStyleToStyle.RelativeTo.Width, listLevelInfo.MarginLeft);
			OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), "padding-left", listLevelInfo.MarginLeftPercent);

			//if (odtNode.OdtTag.Equals("list-item", StrCompICIC))
			//{
			//	OdtNode.AddCssPropertyValue(odtNode, $"{odtNode.HtmlTag}.{odtNode.GetClassName()}:before", "padding-right", listLevelInfo.SpaceBefore, OdtNode.CssPrefix.Element);
			//	OdtNode.AddCssPropertyValue(odtNode, $"{odtNode.HtmlTag}.{odtNode.GetClassName()}:before", "float", "left", OdtNode.CssPrefix.Element);

			//	var nextElement = element.Elements().FirstOrDefault();

			//	if (nextElement != null)
			//	{
			//		if (!nextElement.Name.LocalName.Equals("list", StrCompICIC))
			//		{
			//			OdtNode.AddCssPropertyValue(odtNode, $"{odtNode.HtmlTag}.{odtNode.GetClassName()}:before", "content", "\"■\"", OdtNode.CssPrefix.Element);
			//		}
			//	}
			//}			
		}

		private static void CalculateListSpaces(OdtListLevel listLevelInfo, OdtListLevel prevListLevelInfo, OdtPageInfo pageInfo)
		{
			if (listLevelInfo == null)
			{
				throw new ArgumentNullException(nameof(listLevelInfo));
			}

			var provider = new NumberFormatInfo { NumberDecimalSeparator = "." };

			double.TryParse(GetRealNumber(listLevelInfo.MarginLeft), NumberStyles.Any, provider, out double marginLeft);
			double.TryParse(GetRealNumber(prevListLevelInfo?.MarginLeft), NumberStyles.Any, provider, out double prevMarginLeft);

			double.TryParse(GetRealNumber(listLevelInfo.SpaceBefore), NumberStyles.Any, provider, out double spaceBefore);
			double.TryParse(GetRealNumber(prevListLevelInfo?.SpaceBefore), NumberStyles.Any, provider, out double prevSpaceBefore);

			double.TryParse(GetRealNumber(listLevelInfo.TextIndent), NumberStyles.Any, provider, out double textIntend);
			double.TryParse(GetRealNumber(prevListLevelInfo?.TextIndent), NumberStyles.Any, provider, out double textIntendBefore);

			marginLeft -= (prevMarginLeft - textIntend);
			spaceBefore -= prevSpaceBefore;

			var marginLeftStr = marginLeft.ToString(provider) + GetNumberUnit(listLevelInfo.MarginLeft);
			var spaceBeforeStr = spaceBefore.ToString(provider) + GetNumberUnit(listLevelInfo.SpaceBefore);

			listLevelInfo.MarginLeftPercent = GetCssValuePercentValueRelativeToPage(pageInfo, OdtStyleToStyle.RelativeTo.Width, marginLeftStr);
			listLevelInfo.SpaceBeforePercent = GetCssValuePercentValueRelativeToPage(pageInfo, OdtStyleToStyle.RelativeTo.Width, spaceBeforeStr);
		}

		private static string GetListClassName(OdtNode odtNode)
		{
			if (odtNode == null)
			{
				throw new ArgumentNullException(nameof(odtNode));
			}

			do
			{
				if (odtNode.OdtTag.Equals("list", StrCompICIC)
					&& !string.IsNullOrWhiteSpace(odtNode.OdtElementClassName))
				{
					return odtNode.OdtElementClassName;
				}

				odtNode = odtNode.ParentNode;
			}
			while (odtNode != null);

			return null;
		}

		private void HandleListItemElement(OdtContext ctx, XElement element, OdtNode odtNode)
		{
			if (!element.Name.Equals(XName.Get("list-item", OdtXmlNs.Text)))
			{
				return;
			}

			var listLevel = GetListLevel(odtNode.ParentNode);
			var listLevelInfo = ctx.OdtListsLevelInfo[GetListClassName(odtNode)][listLevel];

			OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), "padding-left", listLevelInfo.SpaceBeforePercent);
			OdtNode.AddCssPropertyValue(odtNode, $"{odtNode.HtmlTag}.{odtNode.GetClassName()}:before", "padding-right", listLevelInfo.MarginLeftPercent, OdtNode.CssPrefix.Element);
			OdtNode.AddCssPropertyValue(odtNode, $"{odtNode.HtmlTag}.{odtNode.GetClassName()}:before", "float", "left", OdtNode.CssPrefix.Element);

			var nextElement = element.Elements().FirstOrDefault();

			if (nextElement != null)
			{
				var nextElementLocalName = nextElement.Name.LocalName;

				if (!nextElementLocalName.Equals("list", StrCompICIC))
				{
					OdtNode.AddCssPropertyValue(odtNode, $"{odtNode.HtmlTag}.{odtNode.GetClassName()}:before", "content", "\"■\"", OdtNode.CssPrefix.Element);
				}
			}
		}

		private static string GetListLevel(OdtNode listNode)
		{
			int listLevel = 1;

			var parentNode = listNode.ParentNode;

			while (parentNode != null)
			{
				if (parentNode.OdtTag.Equals("list", StrCompICIC))
				{
					listLevel++;
				}

				parentNode = parentNode.ParentNode;
			}

			return listLevel.ToString();
		}

		private void HandleImageElement(OdtContext ctx, XElement element, OdtNode odtNode)
		{
			if (!element.Name.Equals(XName.Get("image", OdtXmlNs.Draw)))
			{
				return;
			}

			var hrefAttrVal = element.Attribute(XName.Get("href", OdtXmlNs.XLink))?.Value;
			hrefAttrVal = GetEmbedContentLink(ctx, hrefAttrVal);
			OdtNode.AddAttrValue(odtNode, "src", hrefAttrVal);

			var maxWidth = element.Parent.Attribute(XName.Get("width", OdtXmlNs.SvgCompatible))?.Value;
			var maxHeight = element.Parent.Attribute(XName.Get("height", OdtXmlNs.SvgCompatible))?.Value;
			OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), "max-width", maxWidth);
			OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), "max-height", maxHeight);
		}

		private void ApplyDefaultStyleProperties(OdtNode odtNode, Dictionary<string, string> defaultProps)
		{
			if (defaultProps == null)
			{
				return;
			}

			foreach (var prop in defaultProps)
			{
				OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), prop.Key, prop.Value);
			}
		}

		private void GetCssStyleProperties(OdtContext ctx, OdtNode odtNode)
		{
			if (string.IsNullOrWhiteSpace(odtNode.GetClassName()))
			{
				return;
			}

			var styleElement = FindStyleByNameAttr(odtNode.GetClassName(), "style", ctx.OdtStyles);

			var parentStyleName = GetAttrValOrNull(styleElement, "parent-style-name", OdtXmlNs.Style);
			var defaultStyleFamilyName = GetAttrValOrNull(styleElement, "family", OdtXmlNs.Style);

			var parentStyleElement = FindStyleByNameAttr(parentStyleName, "style", ctx.OdtStyles);
			var defaultStyle = FindDefaultStyle(defaultStyleFamilyName, ctx.OdtStyles);

			TransformOdtStyleElements(ctx, defaultStyle?.Elements(), odtNode);
			TransformOdtStyleElements(ctx, parentStyleElement?.Elements(), odtNode);
			TransformOdtStyleElements(ctx, styleElement?.Elements(), odtNode);
		}

		private void TransformOdtStyleElements(OdtContext ctx, IEnumerable<XElement> odtStyleElements, OdtNode odtNode)
		{
			if (odtStyleElements == null)
			{
				return;
			}

			foreach (var odtStyleElement in odtStyleElements)
			{
				HandleTabStopElement(odtStyleElement, odtNode);
				HandleStyleTrasformation(ctx, odtStyleElement, odtNode);
				TransformOdtStyleElements(ctx, odtStyleElement.Elements(), odtNode);
			}
		}

		private void HandleStyleTrasformation(OdtContext ctx, XElement odtStyleElement, OdtNode odtNode)
		{
			foreach (var attr in odtStyleElement.Attributes())
			{
				var trans = OdtTrans.StyleToStyle
				.Find(p =>
					p.OdtAttrName.Equals(attr.Name.LocalName, StrCompICIC)
					&& p.StyleTypes.Contains(odtStyleElement.Name.LocalName, StringComparer.InvariantCultureIgnoreCase));

				if (trans == null)
				{
					continue;
				}

				if (trans.ValueToValue != null)
				{
					var valueFound = false;

					foreach (var valToVal in trans.ValueToValue)
					{
						if (valToVal.OdtStyleAttr.TryGetValue(attr.Name.LocalName, out string value))
						{
							if (value.Equals(attr.Value, StrCompICIC))
							{
								foreach (var cssProp in valToVal.CssProp)
								{
									var cssPropValue = GetCssValuePercentValueRelativeToPage(ctx.PageInfo, trans.AsPercentage, cssProp.Value);
									OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), cssProp.Key, cssPropValue);
								}

								valueFound = true;
								break;
							}
						}
					}

					if (!valueFound)
					{
						var cssPropVal = GetCssValuePercentValueRelativeToPage(ctx.PageInfo, trans.AsPercentage, attr.Value);
						OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), trans.CssPropName, cssPropVal);
					}
				}
				else
				{
					var cssPropVal = GetCssValuePercentValueRelativeToPage(ctx.PageInfo, trans.AsPercentage, attr.Value);
					OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), trans.CssPropName, cssPropVal);
				}
			}
		}

		private static string GetCssValuePercentValueRelativeToPage(OdtPageInfo pageInfo, OdtStyleToStyle.RelativeTo relativeTo, string value)
		{
			if (relativeTo == OdtStyleToStyle.RelativeTo.None || !IsCssNumberValue(value))
			{
				return value;
			}

			var provider = new NumberFormatInfo
			{
				NumberDecimalSeparator = "."
			};

			string pageValueString = "0";

			if (relativeTo == OdtStyleToStyle.RelativeTo.Width)
			{
				pageValueString = GetRealNumber(pageInfo.WidthBrutto);
			}
			else if (relativeTo == OdtStyleToStyle.RelativeTo.Height)
			{
				pageValueString = GetRealNumber(pageInfo.HeightBrutto);
			}

			var relativeValueString = GetRealNumber(value);

			double.TryParse(pageValueString, NumberStyles.Any, provider, out double pageValue);
			double.TryParse(relativeValueString, NumberStyles.Any, provider, out double relativeValue);

			if (pageValue > 0 && relativeValue > 0 && pageValue > relativeValue)
			{
				return ((relativeValue / pageValue) * 100).ToString(provider) + "%";
			}
			else
			{
				return value;
			}
		}

		private void HandleTabStopElement(XElement odtStyleElement, OdtNode odtNode)
		{
			if (odtStyleElement.Name.Equals(XName.Get("tab-stop", OdtXmlNs.Style)))
			{
				OdtNode.AddTabStop(odtNode, odtStyleElement);
			}
		}

		private XElement FindStyleByNameAttr(string attrName, string styleLocalName, IEnumerable<XElement> odtStyles)
		{
			return odtStyles
				.FirstOrDefault(p =>
				p.Name.LocalName.Equals(styleLocalName, StrCompICIC)
				&& (p.Attribute(XName.Get("name", OdtXmlNs.Style)) != null
					&& p.Attribute(XName.Get("name", OdtXmlNs.Style)).Value.Equals(attrName, StrCompICIC)));
		}

		private XElement FindDefaultStyle(string family, IEnumerable<XElement> odtStyles)
		{
			return odtStyles
				.FirstOrDefault(p =>
					p.Name.LocalName.Equals("default-style", StrCompICIC)
					&& p.Attribute(XName.Get("family", OdtXmlNs.Style)).Value
						.Equals(family, StrCompICIC));
		}

		private static string GetAttrValOrNull(XElement odElement, string attrName, string attrNamespace)
		{
			return odElement?.Attribute(XName.Get(attrName, attrNamespace))?.Value;
		}

		private void CopyElementAttributes(XElement odElement, OdtNode htmlElement)
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

		private static void HandleNonBreakingSpaceElement(XElement element, OdtNode odtNode)
		{
			if (!element.Name.Equals(XName.Get("s", OdtXmlNs.Text)))
			{
				return;
			}

			var spacesValue = element.Attribute(XName.Get("c", OdtXmlNs.Text))?.Value;
			int.TryParse(spacesValue, out int spacesCount);

			if (spacesCount == 0)
			{
				spacesCount++;
			}

			for (int i = 0; i < spacesCount; i++)
			{
				odtNode.InnerText += ("&nbsp;");
			}
		}

		private static void HandleTabElement(XElement element, OdtNode odtNode, OdtNode odtParentNode, string defaultTabSize)
		{
			if (!element.Name.Equals(XName.Get("tab", OdtXmlNs.Text)))
			{
				return;
			}

			var tabLevel = odtParentNode.ChildNodes.Count(p => p.OdtTag.Equals("tab", StrCompICIC));
			var lastTabStopValue = odtParentNode.TabStops.ElementAtOrDefault(tabLevel - 1);
			var tabStopValue = odtParentNode.TabStops.ElementAtOrDefault(tabLevel);
			var tabNodeOdtClassName = $"{odtParentNode.GetClassName()}tab{tabLevel}";

			OdtNode.AddAttrValue(odtNode, "class", tabNodeOdtClassName);

			if (tabStopValue.Equals((null, null)))
			{
				OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), "margin-left", defaultTabSize);
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

				OdtNode.AddCssPropertyValue(odtNode, odtNode.GetClassName(), $"margin-{tabStopValue.type}", value);
			}
		}

		private static void GetLevel(ref int level, string elementTag, string elementNameSpace, OdtNode odtNode)
		{
			if (odtNode != null)
			{
				if (odtNode.OdtTag.Equals("tab", StrCompICIC))
				{
					level++;
				}

				GetLevel(ref level, elementTag, elementNameSpace, odtNode.ParentNode);
			}
		}

		private static void HandleTextNode(XNode node, OdtNode odtNode)
		{
			if (node.NodeType != XmlNodeType.Text)
			{
				return;
			}

			var parentName = node.Parent.Name.LocalName;

			if (!OdtTrans.TextNodeParent.Contains(parentName, StringComparer.InvariantCultureIgnoreCase))
			{
				return;
			}

			var childNode = new OdtNode(parentName, "span", null, odtNode)
			{
				InnerText = ((XText)node).Value
			};
		}

		private static bool IsCssNumberValue(string value)
		{
			return Regex.IsMatch(value, "^[+-]?[0-9]+.?([0-9]+)?(px|em|ex|%|in|cm|mm|pt|pc)$");
		}

		private static string GetRealNumber(string value, bool nullOrWhiteSpaceToZero = true)
		{
			if (String.IsNullOrWhiteSpace(value))
			{
				if (nullOrWhiteSpaceToZero)
				{
					value = "0";
				}
				else
				{
					throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(value));
				}
			}

			return Regex.Match(value, "[+-]?([0-9]*[.])?[0-9]+").Value;
		}

		private static string GetNumberUnit(string value)
		{
			return Regex.Replace(value, @"[\d.]", "");
		}

		private static bool IsCssColorValue(string value)
		{
			return Regex.IsMatch(value, @"^(#[0-9a-f]{3}|#(?:[0-9a-f]{2}){2,4}|(rgb|hsl)a?\((-?\d+%?[,\s]+){2,3}\s*[\d\.]+%?\))$");
		}
	}
}