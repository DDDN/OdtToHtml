﻿/*
DDDN.OdtToHtml.Conversion.OdtConvert
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
using System.Xml.Linq;
using DDDN.OdtToHtml.Exceptions;

namespace DDDN.OdtToHtml.Conversion
{
	public class OdtConvert : IOdtConvert
	{
		private OdtContext Ctx;
		private const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;
		private static readonly NumberFormatInfo NumberFormat = new NumberFormatInfo { NumberDecimalSeparator = "." };

		public OdtConversionOutput Convert(IOdtFile odtFile, OdtConvertSettings convertSettings)
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
			var html = OdtHtmlInfo.RenderHtml(htmlTreeRootTag);
			var css = OdtHtmlInfo.RenderCss(htmlTreeRootTag);
			var firstHeader = GetFirstHeaderHtml(htmlTreeRootTag);
			var firstParagraph = GetFirstParagraphHtml(htmlTreeRootTag);
			var embedContent = Ctx.EmbedContent.Where(p => !string.IsNullOrWhiteSpace(p.Link));

			return new OdtConversionOutput
			{
				Css = css,
				Html = html,
				DocumentFirstHeader = firstHeader,
				DocumentFirstParagraph = firstParagraph,
				EmbedContent = embedContent,
				PageInfo = Ctx.PageInfo,
				UsedFontFamilies = Ctx.UsedFontFamilies
			};
		}

		private void Init(IOdtFile odtFile, OdtConvertSettings convertSettings)
		{
			if (odtFile == null)
				throw new ArgumentNullException(nameof(odtFile));

			if (convertSettings == null)
				throw new ArgumentNullException(nameof(convertSettings));

			var embedContent = odtFile.GetZipArchiveEntries();

			var contentXDoc = OdtFile.GetZipArchiveEntryAsXDocument(embedContent
				.FirstOrDefault(p => p.ContentFullName.Equals("content.xml", StrCompICIC))?.Data);
			var stylesXDoc = OdtFile.GetZipArchiveEntryAsXDocument(embedContent
				.FirstOrDefault(p => p.ContentFullName.Equals("styles.xml", StrCompICIC))?.Data);

			var version = contentXDoc.Root.Attribute(XName.Get("version", OdtXmlNs.Office))?.Value;

			if (version != null
				&& String.CompareOrdinal(version, "1.2") < 0)
			{
				throw new WrongOdtDocumentVersion("Only ODT version 1.2 and higher supported.");
			}

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
			odtStyles = odtStyles.Concat(listStyles);

			odtStyles = odtStyles.Concat(contentXDoc.Root.Descendants(XName.Get("page-layout", OdtXmlNs.Style)));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("page-layout", OdtXmlNs.Style)));

			odtStyles = odtStyles.Concat(contentXDoc.Root.Descendants(XName.Get("master-page", OdtXmlNs.Style)));
			odtStyles = odtStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("master-page", OdtXmlNs.Style)));

			var fontStyles = contentXDoc.Root.Descendants(XName.Get("font-face", OdtXmlNs.Style));
			fontStyles = fontStyles.Concat(stylesXDoc.Root.Descendants(XName.Get("font-face", OdtXmlNs.Style)));
			odtStyles = odtStyles.Concat(fontStyles);

			var (pageInfo, pageInfoCalc) = GetPageInfo(odtStyles);
			var odtListsLevelInfo = OdtListHelper.CreateListLevelInfos(odtStyles);

			Ctx = new OdtContext
			{
				DocumentNodes = documentNodes,
				OdtStyles = odtStyles,
				ConvertSettings = convertSettings,
				EmbedContent = embedContent,
				PageInfo = pageInfo,
				PageInfoCalc = pageInfoCalc,
				OdtListsLevelInfo = odtListsLevelInfo
			};
		}

		private static (OdtPageInfo pageInfo, OdtPageInfoCalc pageInfoCalc) GetPageInfo(IEnumerable<XElement> odtStyles)
		{
			var masterStyle = odtStyles
				.FirstOrDefault(p =>
					p.Name.LocalName.Equals("master-page", StrCompICIC));
			var pageLayoutStyleName = OdtContentHelper.GetOdtElementAttrValOrNull(masterStyle, "page-layout-name", OdtXmlNs.Style);
			var pageLayoutStyleProperties = OdtStyleHelper.FindStyleElementByNameAttr(pageLayoutStyleName, "page-layout", odtStyles)
				?.Element(XName.Get("page-layout-properties", OdtXmlNs.Style));

			var width = pageLayoutStyleProperties?.Attribute(XName.Get("page-width", OdtXmlNs.XslFoCompatible))?.Value;
			var height = pageLayoutStyleProperties?.Attribute(XName.Get("page-height", OdtXmlNs.XslFoCompatible))?.Value;
			var marginTop = pageLayoutStyleProperties?.Attribute(XName.Get("margin-top", OdtXmlNs.XslFoCompatible))?.Value;
			var marginLeft = pageLayoutStyleProperties?.Attribute(XName.Get("margin-left", OdtXmlNs.XslFoCompatible))?.Value;
			var marginBottom = pageLayoutStyleProperties?.Attribute(XName.Get("margin-bottom", OdtXmlNs.XslFoCompatible))?.Value;
			var marginRight = pageLayoutStyleProperties?.Attribute(XName.Get("margin-right", OdtXmlNs.XslFoCompatible))?.Value;

			double.TryParse(OdtCssHelper.GetRealNumber(width), NumberStyles.Any, NumberFormat, out double widthNo);
			double.TryParse(OdtCssHelper.GetRealNumber(height), NumberStyles.Any, NumberFormat, out double heightNo);
			double.TryParse(OdtCssHelper.GetRealNumber(marginTop), NumberStyles.Any, NumberFormat, out double marginTopNo);
			double.TryParse(OdtCssHelper.GetRealNumber(marginLeft), NumberStyles.Any, NumberFormat, out double marginLeftNo);
			double.TryParse(OdtCssHelper.GetRealNumber(marginBottom), NumberStyles.Any, NumberFormat, out double marginBottomNo);
			double.TryParse(OdtCssHelper.GetRealNumber(marginRight), NumberStyles.Any, NumberFormat, out double marginRightNo);

			var widthNettoNo = widthNo - (marginBottomNo + marginRightNo);
			var heightNettoNo = heightNo - (marginTopNo + marginBottomNo);

			var widthNetto = widthNettoNo.ToString(NumberFormat) + OdtCssHelper.GetNumberUnit(width);
			var heightNetto = heightNettoNo.ToString(NumberFormat) + OdtCssHelper.GetNumberUnit(height);

			var pageInfoCalc = new OdtPageInfoCalc
			{
				Width = widthNo,
				WidthUnit = OdtCssHelper.GetNumberUnit(width),
				Height = heightNo,
				HeightUnit = OdtCssHelper.GetNumberUnit(height),
			};

			var pageInfo = new OdtPageInfo
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

			return (pageInfo, pageInfoCalc);
		}

		private static string GetFirstHeaderHtml(OdtHtmlInfo htmlTreeRootTag)
		{
			var headerNode = htmlTreeRootTag.ChildNodes
				.OfType<OdtHtmlInfo>()
				.FirstOrDefault(p => p.OdtTag.Equals("h", StrCompICIC));

			if (headerNode == default(OdtHtmlInfo))
			{
				return null;
			}

			return OdtHtmlInfo.RenderHtml(headerNode);
		}

		private static string GetFirstParagraphHtml(OdtHtmlInfo htmlTreeRootTag)
		{
			var headerNode = htmlTreeRootTag.ChildNodes
				.OfType<OdtHtmlInfo>()
				.FirstOrDefault(p => p.OdtTag.Equals("p", StrCompICIC));

			if (headerNode == default(OdtHtmlInfo))
			{
				return null;
			}

			return OdtHtmlInfo.RenderHtml(headerNode);
		}

		private static OdtHtmlInfo GetHtmlTree(OdtContext ctx)
		{
			var rootNode = CreateHtmlRootNode(ctx.ConvertSettings);
			OdtContentHelper.OdtTextNodeChildsWalker(ctx, ctx.DocumentNodes, rootNode);
			return rootNode;
		}

		private static OdtHtmlInfo CreateHtmlRootNode(OdtConvertSettings convertSettings)
		{
			if (convertSettings == null)
			{
				throw new ArgumentNullException(nameof(convertSettings));
			}

			if (string.IsNullOrWhiteSpace(convertSettings.RootElementTagName))
			{
				throw new OdtToHtmlException("RootElementTagName not provided.");
			}

			var rootElement = new XElement(convertSettings.RootElementTagName);

			if (!string.IsNullOrWhiteSpace(convertSettings.RootElementClassName))
			{
				rootElement.Add(new XAttribute("style-name", convertSettings.RootElementClassName));
			}

			var rootHtmlInfo = new OdtHtmlInfo(rootElement, null);

			if (!string.IsNullOrWhiteSpace(convertSettings.RootElementId))
			{
				rootHtmlInfo.HtmlAttrs.Add("id", new List<string>() { convertSettings.RootElementId });
			}

			return rootHtmlInfo;
		}
	}
}