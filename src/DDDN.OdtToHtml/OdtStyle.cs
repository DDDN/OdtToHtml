/*
DDDN.OdtToHtml.OdtStyle
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

namespace DDDN.OdtToHtml
{
	public class OdtStyle
	{
		private const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;
		public struct StyleType
		{
			public const string default_style = "default-style";
			public const string list_style = "list-style";
			public const string page_layout = "page-layout";
			public const string master_page = "master-page";
			public const string font_face = "font-face";
			public const string style = "style";
		}

		private static readonly string[] SupportedListStyles = {"list-level-style-bullet", "list-level-style-number", "text:list-level-style-image"};
		public string Name { get; set; }
		public string Type { get; set; }
		public Dictionary<string, string> Attrs { get; set; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
		public Dictionary<string, string> CssProps { get; set; } = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
		public List<Dictionary<string, string>> ListLevelProps { get; set; } = new List<Dictionary<string, string>>();

		public List<(string type, string position)> TabStops = new List<(string type, string position)>();

		private OdtStyle()
		{

		}

		public OdtStyle(string name, string type)
		{
			Name = name;
			Type = type;
		}

		public static void HandleTabStopStyle(OdtStyle style, XElement xElement)
		{
			if (!xElement.Name.Equals(XName.Get("tab-stop", OdtXmlNs.Style)))
			{
				return;
			}

			var typeAttrVal = xElement.Attribute(XName.Get("type", OdtXmlNs.Style))?.Value;
			var positionAttrVal = xElement.Attribute(XName.Get("position", OdtXmlNs.Style))?.Value;

			if (positionAttrVal == null)
			{
				return;
			}

			if (typeAttrVal == null)
			{
				typeAttrVal = "left";
			}

			style.TabStops.Add((typeAttrVal, positionAttrVal));
		}

		public static void HandleOdtStyleElements(OdtContext ctx, IEnumerable<XElement> odtStyles, OdtHtmlInfo odtInfo)
		{
			if (odtStyles?.Any() != true)
			{
				return;
			}

			foreach (var odtStyleElement in odtStyles)
			{
				HandleTabStopStyle(ctx.UsedStyles[odtInfo.OdtStyleName], odtStyleElement);
				HandleStyleTrasformation(ctx, odtStyleElement, odtInfo);

				HandleOdtStyleElements(ctx, odtStyleElement.Elements(), odtInfo);
			}
		}

		/// <summary>
		/// Transforms the ODT styles into CSS Styles
		/// </summary>
		/// <param name="ctx">The main context.</param>
		/// <param name="odtInfo">Structure containing transformed information for single HTML tag transformed from ODT tag.</param>
		public static void HandleOdtStyle(OdtContext ctx, OdtHtmlInfo odtInfo)
		{
			if (string.IsNullOrWhiteSpace(odtInfo.OdtStyleName)
				|| ctx.UsedStyles.ContainsKey(odtInfo.OdtStyleName))
			{
				return;
			}

			var style = new OdtStyle(odtInfo.OdtStyleName, "style");
			ctx.UsedStyles.Add(odtInfo.OdtStyleName, style);

			var styleElement = FindStyleElementByNameAttr(odtInfo.OdtStyleName, StyleType.style, ctx.OdtStyles);
			var familyStyleFamilyName = OdtContentHelper.GetOdtElementAttrValOrNull(styleElement, "family", OdtXmlNs.Style);
			var parentStyleName = OdtContentHelper.GetOdtElementAttrValOrNull(styleElement, "parent-style-name", OdtXmlNs.Style);
			//style.ListStyleName = OdtContentHelper.GetOdtElementAttrValOrNull(styleElement, "list-style-name", OdtXmlNs.Style);

			var familyStyle = FindDefaultOdtStyleElement(familyStyleFamilyName, ctx.OdtStyles);
			var parentStyleElement = FindStyleElementByNameAttr(parentStyleName, StyleType.style, ctx.OdtStyles);
			//var listStyleElement = FindStyleElementByNameAttr(listStyleName, StyleType.list_style, ctx.OdtStyles);

			HandleOdtStyleElements(ctx, familyStyle?.Elements(), odtInfo);
			HandleOdtStyleElements(ctx, parentStyleElement?.Elements(), odtInfo);
			HandleOdtStyleElements(ctx, styleElement?.Elements(), odtInfo);
			//HandleOdtStyleElements(ctx, listStyleElement?.Elements(), odtInfo);
		}

		public static void HandleStyleTrasformation(OdtContext ctx, XElement styleElement, OdtHtmlInfo htmlInfo)
		{
			string cssPropVal = null;

			var attrNames = styleElement.Attributes()
				.Select(p => $"{styleElement.Name.LocalName}.{p.Name.LocalName}")
				.ToList();

			foreach (var attr in styleElement.Attributes())
			{
				var attrName = $"{styleElement.Name.LocalName}.{attr.Name.LocalName}";
				var styleTostyle = OdtTrans.StyleToStyle.Find(p => p.OdtAttrNames.Contains(attrName));

				if (styleTostyle == null
					|| attrNames.Intersect(styleTostyle.OverridableBy).Any())
				{
					continue;
				}

				if (attr.Name.LocalName.Equals("font-name", StrCompICIC))
				{
					var fontFamily = HandleFontFamilyStyle(ctx.OdtStyles, attr.Value);

					if (!string.IsNullOrWhiteSpace(fontFamily))
					{
						ctx.UsedStyles[htmlInfo.OdtStyleName].CssProps["font-family"] = fontFamily;
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
								ctx.UsedStyles[htmlInfo.OdtStyleName].CssProps[cssProp.Key] = cssProp.Value;
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

						ctx.UsedStyles[htmlInfo.OdtStyleName].CssProps[styleTostyle.CssPropName] = cssPropVal;
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

					ctx.UsedStyles[htmlInfo.OdtStyleName].CssProps[styleTostyle.CssPropName] = cssPropVal;
				}
			}
		}

		public static IDictionary<string, OdtStyle> GetOdtStylesStyles(IEnumerable<XElement> odtStyles)
		{
			var styles = new Dictionary<string, OdtStyle>(StringComparer.InvariantCultureIgnoreCase);

			foreach (var odtStyle in odtStyles)
			{
				var odtStyleName = odtStyle?.Attribute(XName.Get("name", OdtXmlNs.Style))?.Value;

				if (string.IsNullOrWhiteSpace(odtStyleName)
					|| styles.ContainsKey(odtStyleName))
				{
					continue;
				}

				var style = new OdtStyle(odtStyleName, odtStyle.Name.LocalName);
				styles.Add(style.Name, style);

				GetOdtStyleAttributes(odtStyle, style);
			}

			return styles;
		}

		private static void GetOdtStyleAttributes(XElement odtStyle, OdtStyle style)
		{
			foreach (var odtAttr in odtStyle.Attributes())
			{
				style.Attrs[odtAttr.Name.LocalName] = odtAttr.Value;
			}

			foreach (var styleProperty in odtStyle.Elements())
			{
				if (SupportedListStyles.Contains(styleProperty.Name.LocalName, StringComparer.InvariantCultureIgnoreCase))
				{
					style.ListLevelProps.Add(new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase));
					GetListLevelAttributes(styleProperty, style);
				}
				else
				{
					GetOdtStyleAttributes(styleProperty, style);
				}
			}
		}

		private static void GetListLevelAttributes(XElement styleProp, OdtStyle style)
		{
			foreach (var odtAttr in styleProp.Attributes())
			{
				style.ListLevelProps.Last()[odtAttr.Name.LocalName] = odtAttr.Value;
			}

			foreach (var stylePropChildProp in styleProp.Elements())
			{
				GetListLevelAttributes(stylePropChildProp, style);
			}
		}

		public static string HandleFontFamilyStyle(IEnumerable<XElement> styles, string styleName)
		{
			if (string.IsNullOrWhiteSpace(styleName))
			{
				return null;
			}

			var fontStyle = FindStyleElementByNameAttr(styleName, StyleType.font_face, styles);
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
					p.Name.LocalName.Equals(StyleType.default_style, StrCompICIC)
					&& p.Attribute(XName.Get("family", OdtXmlNs.Style)).Value
						.Equals(family, StrCompICIC));
		}
	}
}
