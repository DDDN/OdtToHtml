/*
DDDN.OdtToHtml.OdtTrans
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

namespace DDDN.OdtToHtml
{
	public static class OdtTrans
	{
		private static readonly StringComparer StrCompICIC = StringComparer.InvariantCultureIgnoreCase;

		public static readonly List<string> TextNodeParent = new List<string>()
		{
			"p", "h", "span"
		};

		public static readonly List<OdtTagToHtml> TagToTag = new List<OdtTagToHtml>()
		{
			{ new OdtTagToHtml {
				OdtName = "image",
				HtmlName = "img",
				DefaultProperty = new Dictionary<string, string>(StrCompICIC)
				{
					["width"] = "100%",
					["height"] = "auto" } } },
			{ new OdtTagToHtml {
				OdtName = "tab",
				HtmlName = "span" } },
			{ new OdtTagToHtml {
				OdtName = "h",
				HtmlName = "p",
				DefaultProperty = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-top"] = "0",
					["margin-bottom"] = "0" } } },
			{ new OdtTagToHtml {
				OdtName = "p",
				HtmlName = "p",
				DefaultProperty = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-top"] = "0",
					["margin-bottom"] = "0" } } },
			{ new OdtTagToHtml { OdtName = "span", HtmlName = "span" } },
			{ new OdtTagToHtml { OdtName = "paragraph", HtmlName = "p" } },
			{ new OdtTagToHtml { OdtName = "s", HtmlName = "span" } },
			{ new OdtTagToHtml { OdtName = "a", HtmlName = "a" } },
			{ new OdtTagToHtml { OdtName = "text-box", HtmlName = "div" } },
			{ new OdtTagToHtml {
				OdtName = "table",
				HtmlName = "table",
				DefaultProperty = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-top"] = "0.5em",
					["margin-bottom"] = "0.5em" } } },
			{ new OdtTagToHtml { OdtName = "table-columns", HtmlName = "tr" } },
			{ new OdtTagToHtml { OdtName = "table-column", HtmlName = "th" } } ,
			{ new OdtTagToHtml { OdtName = "table-row", HtmlName = "tr" } } ,
			{ new OdtTagToHtml { OdtName = "table-header-rows", HtmlName = "tr" } } ,
			{ new OdtTagToHtml {
				OdtName = "table-cell",
				HtmlName = "td",
				DefaultProperty = new Dictionary<string, string>(StrCompICIC)
				{
					["min-height"] = "1em",
					["min-width"] = "1em" } } },
			{ new OdtTagToHtml {
				OdtName = "list",
				HtmlName = "ul",
				DefaultProperty = new Dictionary<string, string>(StrCompICIC)
				{
					["list-style-type"] = "none",
					["margin-left"] = "0",
					["margin-right"] = "0",
					["padding-left"] = "0",
					["padding-right"] = "0" } } },
			{ new OdtTagToHtml {
				OdtName = "list-item",
				HtmlName = "li",
				DefaultProperty = new Dictionary<string, string>(StrCompICIC)
				{
					["margin-left"] = "0",
					["margin-right"] = "0",
					["padding-left"] = "0",
					["padding-right"] = "0" } } },
		};

		public static readonly Dictionary<string, string> AttrNameToAttrName = new Dictionary<string, string>(StrCompICIC)
		{
			["p.style-name"] = "class",
			["span.style-name"] = "class",
			["h.style-name"] = "class",
			["s.style-name"] = "class",
			["table-column.style-name"] = "class",
			["table-row.style-name"] = "class",
			["table-cell.number-columns-spanned"] = "colspan",
			["table-cell.number-rows-spanned"] = "rowspan",
			["table-cell.style-name"] = "class",
			["list.style-name"] = "class",
			["table.style-name"] = "class",
			["a.href"] = "href",
			["a.target-frame-name"] = "target"
		};

		public static readonly List<OdtStyleToStyle> StyleToStyle = new List<OdtStyleToStyle>()
		{
			// width/height
			{ new OdtStyleToStyle {
					OdtAttrName = "width",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "width",
					AsPercentage = OdtStyleToStyle.RelativeTo.Width }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "height",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "height",
					AsPercentage = OdtStyleToStyle.RelativeTo.Height }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "column-width",
					StyleTypes = new List<string> { "table-column-properties" },
					CssPropName = "width",
					AsPercentage = OdtStyleToStyle.RelativeTo.Width }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "column-height",
					StyleTypes = new List<string> { "table-column-properties" },
					CssPropName = "height",
					AsPercentage = OdtStyleToStyle.RelativeTo.Height }
			},
			// align
			{ new OdtStyleToStyle {
					OdtAttrName = "vertical-align",
					StyleTypes = new List<string> { "table-cell-properties" },
					CssPropName = "vertical-align" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "align",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "margin",
					ValueToValue = new List<OdtValueToValue> {
						new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["align"] = "center" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["margin"] = "auto" }
						},
						 new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["align"] = "start" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["margin"] = "0" }
						},
							new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["align"] = "end" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["margin"] = "0" }
						}
					}
				}
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "text-align",
					StyleTypes = new List<string> { "paragraph-properties" },
					CssPropName = "text-align",
					ValueToValue = new List<OdtValueToValue> {
						new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-align"] = "start" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-align"] = "left" }
						},
						 new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-align"] = "end" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-align"] = "right" }
						}
					}
				}
			},
			// color
			{ new OdtStyleToStyle {
					OdtAttrName = "color",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "color" }
			},
			// background
			{ new OdtStyleToStyle {
					OdtAttrName = "background-color",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" }, // "text-properties"
					CssPropName = "background-color" }
			},
			// fonts
			{  new OdtStyleToStyle {
					OdtAttrName = "font-name",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-family" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "font-size",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-size" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "font-style",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-style" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "font-weight",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-weight" }
			},
			// line
			{ new OdtStyleToStyle {
					OdtAttrName = "line-height",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "line-height" }
			},
			// margin
			{ new OdtStyleToStyle {
					OdtAttrName = "margin-top",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-top" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "margin-right",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-right" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "margin-bottom",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-bottom" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "margin-left",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-left" }
			},
			// padding
			{ new OdtStyleToStyle {
					OdtAttrName = "padding-top",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-top" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "padding-right",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-right" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "padding-bottom",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-bottom" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "padding-left",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-left" }
			},
			// border
			{ new OdtStyleToStyle {
					OdtAttrName = "border",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "border-top",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-top" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "border-right",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-right" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "border-bottom",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-bottom" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "border-left",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-left" }
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "border-model",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "border-spacing",
					ValueToValue = new List<OdtValueToValue> {
						new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["border-model"] = "collapsing" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["border-spacing"] = "0" }
						}
					}
				}
			},
			// text
			{ new OdtStyleToStyle {
					OdtAttrName = "text-line-through-style",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "text-decoration",
					NameToValue = "line-through",
					ValueToValue = new List<OdtValueToValue> {
						new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "solid" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "solid" }
						},
						 new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "dotted" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "dotted" }
						},
							new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "dashed" }
						},
						new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "wave" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "wavy" }
						},
						 new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "dot-dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "dashed" }
						},
							new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-style"] = "dot-dot-dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "dotted" }
						}
					}
				}
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "text-underline-style",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "text-decoration",
					NameToValue = "underline",
					ValueToValue = new List<OdtValueToValue> {
						new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "solid" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "solid" }
						},
						 new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "dotted" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "dotted" }
						},
							new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "dashed" }
						},
						new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "wave" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "wavy" }
						},
						 new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "dot-dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "dashed" }
						},
							new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-underline-style"] = "dot-dot-dash" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration"] = "dotted" }
						}
					}
				}
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "width",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "max-width",
					}
			},
			{ new OdtStyleToStyle {
					OdtAttrName = "text-line-through-color",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "text-decoration-color",
					ValueToValue = new List<OdtValueToValue> {
						new OdtValueToValue() {
							OdtStyleAttr = new Dictionary<string, string>(StrCompICIC) { ["text-line-through-color"] = "font-color" },
							CssProp = new Dictionary<string, string>(StrCompICIC) { ["text-decoration-color"] = "inherit" }
						}
					}
				}
			},
			//{ new OdtStyleToStyle {
			//		OdtAttrName = "glyph-orientation-vertical",
			//		StyleTypes = new List<string> { "table-cell-properties" },
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