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
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DDDN.OdtToHtml.Conversion;
using DDDN.OdtToHtml.Transformation;

namespace DDDN.OdtToHtml
{
	[DebuggerDisplay("{StyleType,nq} {StyleOrder}")]
	public class OdtStyle
	{
		private const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;

		public struct StyleTypes
		{
			public const string default_style = "default-style";
			public const string list_style = "list-style";
			public const string page_layout = "page-layout";
			public const string master_page = "master-page";
			public const string font_face = "font-face";
			public const string style = "style";
		}

		private static readonly string[] SupportedListStyles = { "list-level-style-bullet", "list-level-style-number", "text:list-level-style-image" };
		public string Style { get; }
		public string ParentStyle { get; }
		public string ListStyle { get; }
		public string FamilyStyle { get; }
		public string FontStyle { get; }
		public string StyleType { get; }
		public string StyleOrder { get; }
		public List<OdtStyleAttr> Attrs { get; set; } = new List<OdtStyleAttr>();
		public List<OdtStyle> ListLevels { get; set; } = new List<OdtStyle>();
		public List<(string type, string position)> TabStops { get; set; } = new List<(string type, string position)>();

		private OdtStyle()
		{

		}

		public OdtStyle(XElement styleElement)
		{
			StyleType = styleElement.Name.LocalName;
			Style = styleElement.Attribute(XName.Get("name", OdtXmlNs.Style))?.Value;
			ParentStyle = styleElement.Attribute(XName.Get("parent-style-name", OdtXmlNs.Style))?.Value;
			ListStyle = styleElement.Attribute(XName.Get("list-style-name", OdtXmlNs.Style))?.Value;
			FamilyStyle = styleElement.Attribute(XName.Get("family", OdtXmlNs.Style))?.Value;
			FontStyle = styleElement.Element(XName.Get("text-properties", OdtXmlNs.Style))?.Attribute(XName.Get("font-name", OdtXmlNs.Style))?.Value;

			StyleOrder = $"{FontStyle} {FamilyStyle} {ParentStyle} {Style} {ListStyle}";
		}

		public static string RenderCss(OdtContext ctx, OdtHtmlInfo htmlInfo)
		{
			if (htmlInfo == null)
			{
				return null;
			}

			var builder = new StringBuilder(16384);
			RenderTagToTagCss(builder);
			RenderOdtStylesCss(ctx, builder);
			RenderElementCss(ctx, htmlInfo.ChildNodes, builder);
			return builder.ToString();
		}

		private static void RenderOdtStylesCss(OdtContext ctx, StringBuilder builder)
		{
			foreach (var style in ctx.Styles.Where(p => p.Attrs.Count > 0))
			{
				var cssProps = HandleStyleTrasformation(ctx, style);
				RenderCssStyle(builder, OdtCssHelper.NormalizeClassName(style.Style), ".", cssProps);
			}
		}

		public static IDictionary<string, string> HandleStyleTrasformation(OdtContext ctx, OdtStyle style)
		{
			var cssProps2 = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
			var attrs = new List<(string localName, string propName, string value)>();
			var attrNames = style.Attrs.Select(p => $"{p.PropertyName}.{p.LocalName}").ToList();

			foreach (var attr in style.Attrs)
			{
				var styleToStyle = OdtTrans.StyleToStyle.Find(p => string.Equals(p.LocalName, attr.LocalName, StrCompICIC)
					&& p.PropNames.Contains(attr.PropertyName, StringComparer.InvariantCultureIgnoreCase));

				if (styleToStyle == null
					|| styleToStyle.Values == null
					|| styleToStyle.Values.Count == 0)
				{
					continue;
				}

				foreach (var (odtAttrs, cssProps) in styleToStyle.Values)
				{

					foreach (var odtAttr in odtAttrs)
					{
						GetOdtStyles value
					}
				}

					var transAttr = valToVal.OdtAttrs;

				foreach (var valToVal in styleToStyle.ValueToValue)
				{
					
					var odtAttrs = style.Attrs.Where(p => transAttr.Keys.Contains(p.LocalName, StringComparer.InvariantCultureIgnoreCase) && p.PropertyName.Equals();

					  new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);

					foreach (var odtAttr in valToVal.OdtAttrs)
					{
						var addAttr =  style.Attrs.Find(p => p.LocalName.Equals(odtAttr.Key, StrCompICIC));

						if (addAttr != null)
						{

						}
					}

					if (valToVal.OdtAttrs.Count == valToVal.CssProps.Count)
					{
					}
					else if (valToVal.OdtAttrs.Count < valToVal.CssProps.Count)
					{
						foreach (var item in valToVal.OdtAttrs)
						{

						}
					}
					else if (valToVal.OdtAttrs.Count > valToVal.CssProps.Count)
					{
						foreach (var item in valToVal.OdtAttrs)
						{

						}
					}
				}

				if (styleToStyle.ValueToValue.Count > 0)
				{
					var valueFound = false;

					foreach (var valToVal in styleToStyle.ValueToValue)
					{
						if (valToVal.OdtStyleAttr.TryGetValue(attr.LocalName, out string value)
							&& value.Equals(attr.Value, StrCompICIC))
						{
							foreach (var cssProp in valToVal.CssProp)
							{
								cssProps[cssProp.Key] = cssProp.Value;
							}

							valueFound = true;
							break;
						}
					}

					if (!valueFound)
					{
						if (styleToStyle.AsPercentageTo != OdtTransStyleToStyle.RelativeToPage.None)
						{
							cssPropVal = OdtCssHelper.GetCssPercentValueRelativeToPage(attr.Value, ctx.PageInfoCalc, styleToStyle.AsPercentageTo);
						}
						else
						{
							cssPropVal = attr.Value;
						}
						cssProps[styleToStyle.CssPropName] = cssPropVal;
					}
				}
				else
				{
					if (styleToStyle.AsPercentageTo != OdtTransStyleToStyle.RelativeToPage.None)
					{
						cssPropVal = OdtCssHelper.GetCssPercentValueRelativeToPage(attr.Value, ctx.PageInfoCalc, styleToStyle.AsPercentageTo);
					}
					else
					{
						cssPropVal = attr.Value;
					}

					cssProps[styleToStyle.CssPropName] = cssPropVal;
				}
			}

			return cssProps;
		}

		private static void RenderTagToTagCss(StringBuilder builder)
		{
			foreach (var tagToTag in OdtTrans.TagToTag.Where(p => p.DefaultCssProperties?.Any() == true))
			{
				RenderCssStyle(builder, "t2t_" + tagToTag.OdtTag, ".", tagToTag.DefaultCssProperties);
			}
		}

		public static void RenderElementCss(OdtContext ctx, IEnumerable<IOdtHtmlNode> odtNodes, StringBuilder builder)
		{
			foreach (var childNode in odtNodes)
			{
				if (childNode is OdtHtmlInfo htmlInfo)
				{
					//string listItemFontFamilyCssPropValue = "";
					//string firstChildFontFamilyCssPropValue = "";
					//OdtStyle firstChildStyle = null;

					//if (htmlInfo.HtmlTag.Equals("li", StrCompICIC))
					//{
					//	if (htmlInfo.BeforeCssProps?.TryGetValue("font-family", out listItemFontFamilyCssPropValue) == false
					//		|| (htmlInfo.BeforeCssProps?.TryGetValue("font-family", out listItemFontFamilyCssPropValue) == true
					//			&& string.IsNullOrWhiteSpace(listItemFontFamilyCssPropValue)))
					//	{
					//		if ((htmlInfo.ChildNodes?.OfType<OdtHtmlInfo>().FirstOrDefault()?.OwnCssProps?.TryGetValue("font-family", out firstChildFontFamilyCssPropValue) == true
					//			|| (!string.IsNullOrWhiteSpace(htmlInfo.ChildNodes?.OfType<OdtHtmlInfo>().FirstOrDefault()?.OdtStyleName)
					//				&& ctx.UsedStyles.TryGetValue(htmlInfo.ChildNodes.OfType<OdtHtmlInfo>().FirstOrDefault().OdtStyleName, out firstChildStyle)
					//				&& firstChildStyle.CssProps.TryGetValue("font-family", out firstChildFontFamilyCssPropValue)))
					//				&& !string.IsNullOrWhiteSpace(firstChildFontFamilyCssPropValue))
					//		{
					//			AddBeforeCssProps(htmlInfo, "font-family", firstChildFontFamilyCssPropValue);
					//		}
					//	}
					//}

					if (htmlInfo.OwnCssProps.Count > 0)
					{
						RenderCssStyle(builder, "nno-" + htmlInfo.NodeNo, ".", htmlInfo.OwnCssProps);
					}

					if (htmlInfo.BeforeCssProps.Count > 0)
					{
						RenderCssStyle(builder, "nno-" + htmlInfo.NodeNo + ":before", ".", htmlInfo.BeforeCssProps);
					}

					RenderElementCss(ctx, htmlInfo.ChildNodes, builder);
				}
			}
		}

		private static void RenderCssStyle(StringBuilder builder, string styleName, string styleNamePrefix, IDictionary<string, string> styleProperties)
		{
			builder
				.Append(Environment.NewLine)
				.Append(styleNamePrefix)
				.Append(styleName)
				.Append(" {")
				.Append(Environment.NewLine);
			RenderCssStyleProperties(styleProperties, builder);
			builder.Append(" }");
		}

		private static void RenderCssStyleProperties(IDictionary<string, string> cssProperties, StringBuilder builder)
		{
			foreach (var prop in cssProperties)
			{
				builder
					.Append(prop.Key)
					.Append(": ")
					.Append(prop.Value)
					.Append(";")
					.Append(Environment.NewLine);
			}
		}

		//public static string HandleFontFamilyStyle(IEnumerable<XElement> styles, string styleName)
		//{
		//	if (string.IsNullOrWhiteSpace(styleName))
		//	{
		//		return null;
		//	}

		//	var fontStyle = FindStyleElementByNameAttr(styleName, StyleTypes.font_face, styles);
		//	var fontFamily = OdtContentHelper.GetOdtElementAttrValOrNull(fontStyle, "font-family", OdtXmlNs.SvgCompatible);
		//	var fontFamilyGeneric = OdtContentHelper.GetOdtElementAttrValOrNull(fontStyle, "font-family-generic", OdtXmlNs.Style);
		//	return $"\"{fontFamily?.Replace("\"", "").Replace("'", "")}\", \"{fontFamilyGeneric?.Replace("\"", "").Replace("'", "")}\""; // TODO
		//}

		public static List<OdtStyle> GetOdtStyles(IEnumerable<XElement> odtStyles)
		{
			var styles = new List<OdtStyle>();

			foreach (var odtStyle in odtStyles)
			{
				if (OdtContentHelper.TryGetXAttrValue(odtStyle, "name", OdtXmlNs.Style, out string odtStyleName)
				&& !styles.Any(p => p.Style?.Equals(odtStyleName, StrCompICIC) == false))
				{
					var style = new OdtStyle(odtStyle);
					styles.Add(style);
					GetStyleProperties(style, odtStyle.Elements());
				}
			}

			return styles;
		}

		private static void GetStyleProperties(OdtStyle style, IEnumerable<XElement> styleProps)
		{
			foreach (var styleProp in styleProps)
			{
				if (styleProp.Name.LocalName.Equals("tab-stops", StrCompICIC))
				{
					foreach (var tabStopProp in styleProp.Elements())
					{
						OdtContentHelper.TryGetXAttrValue(tabStopProp, "position", OdtXmlNs.Style, out string position);
						OdtContentHelper.TryGetXAttrValue(tabStopProp, "type", OdtXmlNs.Style, out string type);
						style.TabStops.Add((type, position));
					}
				}
				else if (SupportedListStyles.Contains(styleProp.Name.LocalName, StringComparer.InvariantCultureIgnoreCase))
				{
					var listLevel = new OdtStyle(styleProp);
					style.ListLevels.Add(listLevel);
					GetPropertyAttributes(styleProp, listLevel);
					GetStyleProperties(listLevel, styleProp.Elements());
				}
				else
				{
					GetPropertyAttributes(styleProp, style);
					GetStyleProperties(style, styleProp.Elements());
				}
			}
		}

		private static void GetPropertyAttributes(XElement odtProp, OdtStyle style)
		{
			foreach (var odtAttr in odtProp.Attributes())
			{
				style.Attrs.Add(new OdtStyleAttr(odtAttr, odtProp.Name.LocalName));
			}
		}

		public static XElement FindStyleElementByNameAttr(string attrName, string styleLocalName, IEnumerable<XElement> odtStyles)
		{
			return odtStyles
				.FirstOrDefault(p =>
				p.Name.LocalName.Equals(styleLocalName, StrCompICIC)
				&& (p.Attribute(XName.Get("name", OdtXmlNs.Style))?.Value.Equals(attrName, StrCompICIC) == true));
		}
	}
}
