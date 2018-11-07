﻿/*
DDDN.OdtToHtml.Transformation.OdtTrans
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

namespace DDDN.OdtToHtml.Transformation
{
	public static class OdtTrans
	{
		private static readonly StringComparer StrCompICIC = StringComparer.InvariantCultureIgnoreCase;

		public static readonly List<string> TextNodeParent = new List<string>()
		{
			"p", "h", "span", "a"
		};

		public static readonly List<OdtTransTagToTag> TagToTag = new List<OdtTransTagToTag>()
		{
			{ new OdtTransTagToTag {
				OdtName = "image",
				HtmlName = "img",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["width"] = "100%",
					["height"] = "auto" } } },
			{ new OdtTransTagToTag {
				OdtName = "tab",
				HtmlName = "span" } },
			{ new OdtTransTagToTag {
				OdtName = "h",
				HtmlName = "h",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-top"] = "0",
					["margin-right"] = "0",
					["margin-bottom"] = "0",
					["margin-left"] = "0",
					["padding-top"] = "0",
					["padding-right"] = "0",
					["padding-bottom"] = "0",
					["padding-left"] = "0" } } },
			{ new OdtTransTagToTag {
				OdtName = "p",
				HtmlName = "p",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-top"] = "0",
					["margin-right"] = "0",
					["margin-bottom"] = "0",
					["margin-left"] = "0",
					["padding-top"] = "0",
					["padding-right"] = "0",
					["padding-bottom"] = "0",
					["padding-left"] = "0" } } },
			{ new OdtTransTagToTag {
				OdtName = "span",
				HtmlName = "span",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-top"] = "0",
					["margin-right"] = "0",
					["margin-bottom"] = "0",
					["margin-left"] = "0",
					["padding-top"] = "0",
					["padding-right"] = "0",
					["padding-bottom"] = "0",
					["padding-left"] = "0" } } },
			{ new OdtTransTagToTag { OdtName = "paragraph", HtmlName = "p" } },
			{ new OdtTransTagToTag { OdtName = "a", HtmlName = "a" } },
			{ new OdtTransTagToTag { OdtName = "text-box", HtmlName = "div" } },
			{ new OdtTransTagToTag {
				OdtName = "table",
				HtmlName = "table",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-top"] = "0.5em",
					["margin-right"] = "0",
					["margin-bottom"] = "0.5em",
					["margin-left"] = "0",
					["padding-top"] = "0",
					["padding-right"] = "0",
					["padding-bottom"] = "0",
					["padding-left"] = "0" } } },
			{ new OdtTransTagToTag { OdtName = "table-columns", HtmlName = "tr" } },
			{ new OdtTransTagToTag { OdtName = "table-column", HtmlName = "th" } } ,
			{ new OdtTransTagToTag { OdtName = "table-row", HtmlName = "tr" } } ,
			{ new OdtTransTagToTag { OdtName = "table-header-rows", HtmlName = "tr" } } ,
			{ new OdtTransTagToTag {
				OdtName = "table-cell",
				HtmlName = "td",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["height"] = "1rem" } } },
			{ new OdtTransTagToTag {
				OdtName = "list",
				HtmlName = "ul",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["list-style-type"] = "none",
					["margin-top"] = "0",
					["margin-right"] = "0",
					["margin-bottom"] = "0",
					["margin-left"] = "0",
					["padding-top"] = "0",
					["padding-right"] = "0",
					["padding-bottom"] = "0",
					["padding-left"] = "0" } } },
			{ new OdtTransTagToTag {
				OdtName = "list-item",
				HtmlName = "li",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-top"] = "0",
					["margin-right"] = "0",
					["margin-bottom"] = "0",
					["margin-left"] = "0",
					["padding-top"] = "0",
					["padding-right"] = "0",
					["padding-bottom"] = "0",
					["padding-left"] = "0" } } },
			{ new OdtTransTagToTag {
				OdtName = "list-header",
				HtmlName = "li",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-top"] = "0",
					["margin-right"] = "0",
					["margin-bottom"] = "0",
					["margin-left"] = "0",
					["padding-top"] = "0",
					["padding-right"] = "0",
					["padding-bottom"] = "0",
					["padding-left"] = "0" } } },
			{ new OdtTransTagToTag {
				OdtName = "bookmark-start",
				HtmlName = "span" } },
			{ new OdtTransTagToTag {
				OdtName = "bookmark",
				HtmlName = "span" } },
		};

		public static readonly Dictionary<string, string> OdtAttrToHtmlAttr = new Dictionary<string, string>(StrCompICIC)
		{
			["table-cell.number-columns-spanned"] = "colspan",
			["table-cell.number-rows-spanned"] = "rowspan",
			["a.href"] = "href",
			["a.target-frame-name"] = "target",
			["bookmark-start.name"] = "id",
			["bookmark.name"] = "id"
		};

		public static readonly List<string> OdtAttr = new List<string>()
		{
			"continue-numbering", "outline-level"
		};

		public static readonly List<OdtTransStyleToStyle> StyleToStyle = new List<OdtTransStyleToStyle>()
		{
			// width/height
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "table-properties.width" },
					CssPropName = "width",
					AsPercentageTo = OdtTransStyleToStyle.RelativeTo.Width,
					OverridableBy = new List<string> { "table-properties.rel-width" } }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "table-properties.rel-width" },
					CssPropName = "width",
					AsPercentageTo = OdtTransStyleToStyle.RelativeTo.Width }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "table-properties.height" },
					CssPropName = "height" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "table-column-properties.column-width" },
					CssPropName = "width",
					AsPercentageTo = OdtTransStyleToStyle.RelativeTo.Width }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "table-column-properties.column-height" },
					CssPropName = "height" }
			},
			 //align
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "table-cell-properties.vertical-align" },
					CssPropName = "vertical-align" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "table-properties.align" },
					CssPropName = "margin",
					ValueToValue = new List<OdtTransValueToValue> {
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["align"] = "center" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["margin"] = "auto" }
						},
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["align"] = "margins" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["margin"] = "auto" }
						},
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["align"] = "start" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["margin-left"] = "0" }
						},
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["align"] = "end" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["margin-right"] = "0" }
						},
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["align"] = "left" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["margin-left"] = "0" }
						},
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["align"] = "right" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["margin-right"] = "0" }
						}
					}
				}
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.text-align" },
					CssPropName = "text-align",
					ValueToValue = new List<OdtTransValueToValue> {
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-align"] = "start" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-align"] = "left" }
						},
						 new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-align"] = "end" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-align"] = "right" }
						}
					}
				}
			},
			// color
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.color" },
					CssPropName = "color" }
			},
			// background
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.background-color", "paragraph-properties.background-color", "table-cell-properties.background-color" },
					CssPropName = "background-color" }
			},
			// fonts
			{  new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.font-name" },
					CssPropName = "font-family" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.font-size" },
					CssPropName = "font-size" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.font-style" },
					CssPropName = "font-style" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.font-weight" },
					CssPropName = "font-weight" }
			},
			// line
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.line-height" },
					CssPropName = "line-height" }
			},
			// margin
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.margin-top", "table-properties.margin-top" },
					CssPropName = "margin-top" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.margin-right", "table-properties.margin-right" },
					CssPropName = "margin-right" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.margin-bottom", "table-properties.margin-bottom" },
					CssPropName = "margin-bottom" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.margin-left", "table-properties.margin-left" },
					CssPropName = "margin-left" }
			},
			// padding
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.padding-top", "table-cell-properties.padding-top" },
					CssPropName = "padding-top" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.padding-right", "table-cell-properties.padding-right" },
					CssPropName = "padding-right" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.padding-bottom", "table-cell-properties.padding-bottom" },
					CssPropName = "padding-bottom" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.padding-left", "table-cell-properties.padding-left" },
					CssPropName = "padding-left" }
			},
			// border
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.border", "table-cell-properties.border" },
					CssPropName = "border" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.border-top", "table-cell-properties.border-top" },
					CssPropName = "border-top" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.border-right", "table-cell-properties.border-right" },
					CssPropName = "border-right" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.border-bottom", "table-cell-properties.border-bottom" },
					CssPropName = "border-bottom" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "paragraph-properties.border-left", "table-cell-properties.border-left" },
					CssPropName = "border-left" }
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "table-properties.border-model" },
					CssPropName = "border-spacing",
					ValueToValue = new List<OdtTransValueToValue> {
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["border-model"] = "collapsing" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["border-spacing"] = "0" }
						}
					}
				}
			},
			// text
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.text-line-through-style" },
					CssPropName = "text-decoration-style",
					ValueToValue = new List<OdtTransValueToValue> {
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "solid" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "solid", ["text-decoration-line"] = "line-through" }
						},
						 new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "dotted" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "dotted", ["text-decoration-line"] = "line-through" }
						},
							new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "dashed", ["text-decoration-line"] = "line-through" }
						},
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "wave" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "wavy", ["text-decoration-line"] = "line-through" }
						},
						 new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "dot-dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "dashed", ["text-decoration-line"] = "line-through" }
						},
							new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "dot-dot-dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "dotted", ["text-decoration-line"] = "line-through" }
						}
					}
				}
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.text-underline-style" },
					CssPropName = "text-decoration-style",
					ValueToValue = new List<OdtTransValueToValue> {
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "solid" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "solid", ["text-decoration-line"] = "underline" }
						},
						 new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "dotted" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "dotted", ["text-decoration-line"] = "underline" }
						},
							new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "dashed", ["text-decoration-line"] = "underline" }
						},
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "wave" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "wavy", ["text-decoration-line"] = "underline" }
						},
						 new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "dot-dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "dashed", ["text-decoration-line"] = "underline" }
						},
							new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "dot-dot-dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-style"] = "dotted", ["text-decoration-line"] = "underline" }
						}
					}
				}
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.width" },
					CssPropName = "max-width"
					}
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.text-line-through-color" },
					CssPropName = "text-decoration-color",
					ValueToValue = new List<OdtTransValueToValue> {
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-color"] = "font-color" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-color"] = "inherit" }
						}
					}
				}
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.text-underline-color" },
					CssPropName = "text-decoration-color",
					ValueToValue = new List<OdtTransValueToValue> {
						new OdtTransValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-color"] = "font-color" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-color"] = "inherit" }
						}
					}
				}
			},
			{ new OdtTransStyleToStyle {
					OdtAttrNames = new List<string> { "text-properties.text-transform" },
					CssPropName = "text-transform"
				}
			},

			// TODO
			//{ new OdtStyleToStyle {
			//		OdtAttrName = "",
			//		OdtAttrNames = new List<string> { "table-cell-properties.glyph-orientation-vertical" },
			//		CssPropName = "writing-mode",
			//		ValueToValue = new List<OdtValueToValue> {
			//			new OdtValueToValue() {
			//				CssProp = new Dictionary<string, string>(StrCompICIC) { ["glyph-orientation-vertical"] = "0" },
			//				OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["writing-mode"] = "vertical-rl" }
			//			}
			//		}
			//	}
			//}
		};
	}
}