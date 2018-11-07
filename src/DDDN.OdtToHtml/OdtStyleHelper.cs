/*
DDDN.OdtToHtml.OdtStyleHelper
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
using System.Xml.Linq;
using DDDN.OdtToHtml.Conversion;
using DDDN.OdtToHtml.Transformation;
using static DDDN.OdtToHtml.OdtHtmlInfo;

namespace DDDN.OdtToHtml
{
	public static class OdtStyleHelper
	{
		private const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;

		public static void TransformOdtStyleElements(OdtContext ctx, IEnumerable<XElement> odtStyles, OdtHtmlInfo odtInfo)
		{
			if (odtStyles == null)
			{
				return;
			}

			foreach (var odtStyleElement in odtStyles)
			{
				OdtContentHelper.HandleTabStopElement(ctx.UsedStyles[odtInfo.OdtStyleName], odtStyleElement);
				HandleStyleTrasformation(ctx, odtStyleElement, odtInfo);
				TransformOdtStyleElements(ctx, odtStyleElement.Elements(), odtInfo);
			}
		}

		/// <summary>
		/// Transforms the ODT styles into CSS Styles
		/// </summary>
		/// <param name="ctx">The main context.</param>
		/// <param name="odtInfo">Structure containing transformed information for single HTML tag transformed from ODT tag.</param>
		public static void GetOdtStyle(OdtContext ctx, OdtTransTagToTag transTagToTrag, OdtHtmlInfo odtInfo)
		{
			TryApplyDefaultCssStyleProperties(odtInfo, transTagToTrag?.DefaultCssProperties);

			if (string.IsNullOrWhiteSpace(odtInfo.OdtStyleName))
			{
				return;
			}
			// TODO add default/tag css properties to elements without a style-name

			if (!ctx.UsedStyles.ContainsKey(odtInfo.OdtStyleName))
			{
				var style = new OdtStyle(odtInfo.OdtStyleName);
				ctx.UsedStyles.Add(odtInfo.OdtStyleName, style);
			}

			var styleElement = FindStyleElementByNameAttr(odtInfo.OdtStyleName, "style", ctx.OdtStyles);

			var familyStyleFamilyName = OdtContentHelper.GetOdtElementAttrValOrNull(styleElement, "family", OdtXmlNs.Style);
			var parentStyleName = OdtContentHelper.GetOdtElementAttrValOrNull(styleElement, "parent-style-name", OdtXmlNs.Style);
			var listStyleName = OdtContentHelper.GetOdtElementAttrValOrNull(styleElement, "list-style-name", OdtXmlNs.Style);

			var familyStyle = FindDefaultOdtStyleElement(familyStyleFamilyName, ctx.OdtStyles);
			var parentStyleElement = FindStyleElementByNameAttr(parentStyleName, "style", ctx.OdtStyles);
			var listStyleElement = FindStyleElementByNameAttr(listStyleName, "list-style", ctx.OdtStyles);

			TransformOdtStyleElements(ctx, familyStyle?.Elements(), odtInfo);
			TransformOdtStyleElements(ctx, parentStyleElement?.Elements(), odtInfo);
			TransformOdtStyleElements(ctx, styleElement?.Elements(), odtInfo);
			TransformOdtStyleElements(ctx, listStyleElement?.Elements(), odtInfo);
		}

		private static bool TryApplyDefaultCssStyleProperties(OdtHtmlInfo odtInfo, Dictionary<string, string> defaultProperties)
		{
			if (odtInfo == null
				|| defaultProperties == null)
			{
				return false;
			}

			foreach (var prop in defaultProperties)
			{
				TryAddCssPropertyValue(odtInfo, prop.Key, prop.Value, ClassKind.Odt);
			}

			return true;
		}

		public static void HandleStyleTrasformation(OdtContext ctx, XElement odtStyleElement, OdtHtmlInfo odtInfo)
		{
			string cssPropVal = null;

			var attrNames = odtStyleElement.Attributes()
				.Select(p => $"{odtStyleElement.Name.LocalName}.{p.Name.LocalName}")
				.ToList();

			foreach (var attr in odtStyleElement.Attributes())
			{
				var attrName = $"{odtStyleElement.Name.LocalName}.{attr.Name.LocalName}";
				var styleTostyle = OdtTrans.StyleToStyle.Find(p => p.OdtAttrNames.Contains(attrName));

				if (styleTostyle == null
					|| attrNames.Intersect(styleTostyle.OverridableBy).Any())
				{
					continue;
				}

				if (attr.Name.LocalName.Equals("font-name", StrCompICIC))
				{
					var fontFamily = HandleFontFamilyStyle(ctx.OdtStyles, attr.Value);

					if (!String.IsNullOrWhiteSpace(fontFamily))
					{
						TryAddCssPropertyValue(odtInfo, "font-family", fontFamily, ClassKind.Odt);
						ctx.UsedFontFamilies.Remove(fontFamily);
						ctx.UsedFontFamilies.Add(fontFamily);
					}
				}
				else if (styleTostyle.ValueToValue.Count > 0)
				{
					var valueFound = false;

					foreach (var valToVal in styleTostyle.ValueToValue)
					{
						if (valToVal.OdtStyleAttr.TryGetValue(attr.Name.LocalName, out string value)
							&& value.Equals(attr.Value, StrCompICIC))
						{
							foreach (var cssProp in valToVal.CssProp)
							{
								TryAddCssPropertyValue(odtInfo, cssProp.Key, cssProp.Value, ClassKind.Odt);
							}

							valueFound = true;
							break;
						}
					}

					if (!valueFound)
					{
						if (styleTostyle.AsPercentageTo != OdtTransStyleToStyle.RelativeTo.None)
						{
							cssPropVal = OdtCssHelper.GetCssPercentValueRelativeToPage(attr.Value, ctx.PageInfoCalc, styleTostyle.AsPercentageTo);
						}
						else
						{
							cssPropVal = attr.Value;
						}

						TryAddCssPropertyValue(odtInfo, styleTostyle.CssPropName, cssPropVal, ClassKind.Odt);
					}
				}
				else
				{
					if (styleTostyle.AsPercentageTo != OdtTransStyleToStyle.RelativeTo.None)
					{
						cssPropVal = OdtCssHelper.GetCssPercentValueRelativeToPage(attr.Value, ctx.PageInfoCalc, styleTostyle.AsPercentageTo);
					}
					else
					{
						cssPropVal = attr.Value;
					}

					TryAddCssPropertyValue(odtInfo, styleTostyle.CssPropName, cssPropVal, ClassKind.Odt);
				}
			}
		}

		public static string HandleFontFamilyStyle(IEnumerable<XElement> styles, string styleName)
		{
			if (String.IsNullOrWhiteSpace(styleName))
			{
				return null;
			}

			var fontStyle = FindStyleElementByNameAttr(styleName, "font-face", styles);
			var fontFamily = OdtContentHelper.GetOdtElementAttrValOrNull(fontStyle, "font-family", OdtXmlNs.SvgCompatible);
			var fontFamilyGeneric = OdtContentHelper.GetOdtElementAttrValOrNull(fontStyle, "font-family-generic", OdtXmlNs.Style);
			return $"\"{fontFamily?.Replace("\"", "").Replace("'", "")}\", \"{fontFamilyGeneric?.Replace("\"", "").Replace("'", "")}\""; // TODO
		}

		public static XElement FindStyleElementByNameAttr(string attrName, string styleLocalName, IEnumerable<XElement> odtStyles)
		{
			return odtStyles
				.FirstOrDefault(p =>
				p.Name.LocalName.Equals(styleLocalName, StrCompICIC)
				&& (p.Attribute(XName.Get("name", OdtXmlNs.Style))?.Value.Equals(attrName, StrCompICIC) == true));
		}

		public static XElement FindDefaultOdtStyleElement(string family, IEnumerable<XElement> odtStyles)
		{
			return odtStyles
				.FirstOrDefault(p =>
					p.Name.LocalName.Equals("default-style", StrCompICIC)
					&& p.Attribute(XName.Get("family", OdtXmlNs.Style)).Value
						.Equals(family, StrCompICIC));
		}
	}
}
