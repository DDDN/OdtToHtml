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

namespace DDDN.OdtToHtml
{
	public static class OdtStyleHelper
	{
		private const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;

		public static void ApplyDefaultCssStyleProperties(OdtInfo odtInfo, Dictionary<string, string> defaultProps)
		{
			if (odtInfo == null
				|| defaultProps == null)
			{
				return;
			}

			foreach (var prop in defaultProps)
			{
				OdtInfo.AddCssPropertyValue(odtInfo, prop.Key, prop.Value);
			}
		}

		public static void TransformOdtStyleElements(OdtContext ctx, IEnumerable<XElement> odtStyles, OdtInfo odtInfo)
		{
			if (odtStyles == null)
			{
				return;
			}

			foreach (var odtStyleElement in odtStyles)
			{
				OdtContentHelper.HandleTabStopElement(odtStyleElement, odtInfo);
				HandleStyleTrasformation(ctx, odtStyleElement, odtInfo);
				TransformOdtStyleElements(ctx, odtStyleElement.Elements(), odtInfo);
			}
		}

		public static void GetStylesProperties(OdtContext ctx, OdtInfo odtInfo)
		{
			if (string.IsNullOrWhiteSpace(odtInfo.ClassName))
			{
				return;
			}

			var styleElement = FindStyleElementByNameAttr(odtInfo.ClassName, "style", ctx.OdtStyles);

			var defaultStyleFamilyName = OdtContentHelper.GetOdtElementAttrValOrNull(styleElement, "family", OdtXmlNs.Style);
			var parentStyleName = OdtContentHelper.GetOdtElementAttrValOrNull(styleElement, "parent-style-name", OdtXmlNs.Style);
			//var listStyleName = GetOdtElementAttrValOrNull(styleElement, "list-style-name", OdtXmlNs.Style);

			var defaultStyle = FindDefaultOdtStyleElement(defaultStyleFamilyName, ctx.OdtStyles);
			var parentStyleElement = FindStyleElementByNameAttr(parentStyleName, "style", ctx.OdtStyles);
			//var listStyleElement = FindStyleElementByNameAttr(listStyleName, "list-style", ctx.OdtStyles);

			//TransformOdtStyleElements(ctx, listStyleElement?.Elements(), odtInfo);
			TransformOdtStyleElements(ctx, defaultStyle?.Elements(), odtInfo);
			TransformOdtStyleElements(ctx, parentStyleElement?.Elements(), odtInfo);
			TransformOdtStyleElements(ctx, styleElement?.Elements(), odtInfo);
		}

		public static void HandleStyleTrasformation(OdtContext ctx, XElement odtStyleElement, OdtInfo odtInfo)
		{
			string cssPropVal = null;

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

				if (attr.Name.LocalName.Equals("font-name", StrCompICIC))
				{
					var fontFamily = HandleFontFamilyStyle(ctx.OdtStyles, attr.Value);

					if (!String.IsNullOrWhiteSpace(fontFamily))
					{
						OdtInfo.AddCssPropertyValue(odtInfo, "font-family", fontFamily);
						ctx.UsedFontFamilies.Remove(fontFamily);
						ctx.UsedFontFamilies.Add(fontFamily);
					}
				}
				else if (trans.ValueToValue != null)
				{
					var valueFound = false;

					foreach (var valToVal in trans.ValueToValue)
					{
						if (valToVal.OdtStyleAttr.TryGetValue(attr.Name.LocalName, out string value)
							&& value.Equals(attr.Value, StrCompICIC))
						{
							foreach (var cssProp in valToVal.CssProp)
							{
								OdtInfo.AddCssPropertyValue(odtInfo, cssProp.Key, cssProp.Value);
							}

							valueFound = true;
							break;
						}
					}

					if (!valueFound)
					{
						if (trans.AsPercentageTo != OdtStyleToStyle.RelativeTo.None)
						{
							cssPropVal = OdtCssHelper.GetCssValuePercentValueRelativeToPage(attr.Value, ctx.PageInfoCalc, trans.AsPercentageTo);
						}
						else
						{
							cssPropVal = attr.Value;
						}

						OdtInfo.AddCssPropertyValue(odtInfo, trans.CssPropName, cssPropVal);
					}
				}
				else
				{
					if (trans.AsPercentageTo != OdtStyleToStyle.RelativeTo.None)
					{
						cssPropVal = OdtCssHelper.GetCssValuePercentValueRelativeToPage(attr.Value, ctx.PageInfoCalc, trans.AsPercentageTo);
					}
					else
					{
						cssPropVal = attr.Value;
					}

					OdtInfo.AddCssPropertyValue(odtInfo, trans.CssPropName, cssPropVal);
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
			return $"\"{fontFamily}\", \"{fontFamilyGeneric}\"";
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
