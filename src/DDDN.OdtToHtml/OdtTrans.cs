/*
DDDN.OdtToHtml.OdtTrans
Copyright(C) 2017 Lukasz Jaskiewicz(lukasz @jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System.Collections.Generic;

namespace DDDN.OdtToHtml
{
	public static class OdtTrans
	{
		public static readonly List<OdtTagToHtml> OdtTagNameToHtmlTagName = new List<OdtTagToHtml>()
		{
			{ new OdtTagToHtml { OdtName = "h", HtmlName = "p" } },
			{ new OdtTagToHtml { OdtName = "p", HtmlName = "p" } },
			{ new OdtTagToHtml { OdtName = "span", HtmlName = "span" } },
			{ new OdtTagToHtml { OdtName = "paragraph", HtmlName = "p" } },
			{ new OdtTagToHtml {
				OdtName = "graphic",
				HtmlName = "img",
				DefaultProperty = new Dictionary<string, string>
				{
					["width"] = "100%",
					["height"] = "auto"
				} } },
			{ new OdtTagToHtml {
				OdtName = "image",
				HtmlName = "img",
				DefaultProperty = new Dictionary<string, string>
				{
					["width"] = "100%",
					["height"] = "auto"
				} } },
			{ new OdtTagToHtml { OdtName = "s", HtmlName = "span" } },
			{ new OdtTagToHtml { OdtName = "a", HtmlName = "a" } },
			{ new OdtTagToHtml { OdtName = "frame", HtmlName = "div" } },
			{ new OdtTagToHtml { OdtName = "text-box", HtmlName = "div" } },
			{ new OdtTagToHtml { OdtName = "table", HtmlName = "table" } },
			{ new OdtTagToHtml { OdtName = "table-columns", HtmlName = "tr" } },
			{ new OdtTagToHtml { OdtName = "table-column", HtmlName = "th" } } ,
			{ new OdtTagToHtml { OdtName = "table-row", HtmlName = "tr" } } ,
			{ new OdtTagToHtml { OdtName = "table-cell", HtmlName = "td" } } ,
			{ new OdtTagToHtml { OdtName = "list", HtmlName = "ul" } } ,
			{ new OdtTagToHtml { OdtName = "list-item", HtmlName = "li" } } ,
			{ new OdtTagToHtml { OdtName = "automatic-styles", HtmlName = "style" } }
		};

		public static readonly Dictionary<string, string> OdtNodeAttrToHtmlNodeAttr = new Dictionary<string, string>()
		{
			["p.style-name"] = "class",
			["span.style-name"] = "class",
			["h.style-name"] = "class",
			["s.style-name"] = "class",
			["table-column.style-name"] = "class",
			["table-row.style-name"] = "class",
			["table-cell.style-name"] = "class",
			["list.style-name"] = "class",
			["frame.style-name"] = "class",
			["table.style-name"] = "class",
			["a.href"] = "href",
			["image.href"] = "src",
			["a.target-frame-name"] = "target"
		};

		public static readonly Dictionary<string, string> OdtEleAttrToCssProp = new Dictionary<string, string>()
		{
			["frame.width"] = "max-width",
			["frame.height"] = "max-height",
		};

		public static readonly List<OdtStyleAttrToCssProperty> OdtStyleToCssStyle = new List<OdtStyleAttrToCssProperty>()
		{
			// width/height
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "width",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "max-width" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "rel-width",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "width" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "column-width",
					StyleTypes = new List<string> { "table-column-properties" },
					CssPropName = "max-width" }
			},
			// align
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "vertical-align",
					StyleTypes = new List<string> { "table-cell-properties" },
					CssPropName = "vertical-align" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "align",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "align" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "text-align",
					StyleTypes = new List<string> { "paragraph-properties" },
					CssPropName = "text-align",
					ValueToValue = new Dictionary<string, string>() {
						["start"] = "left",
						["end"] = "right" } }
			},
			// color
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "color",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "color" }
			},
			// background
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "background-color",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" }, // "text-properties"
					CssPropName = "background-color" }
			},
			// fonts
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "font-name",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-family" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "font-size",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-size" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "font-style",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-style" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "font-weight",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-weight" }
			},
			// line
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "line-height",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "line-height" }
			},
			// margin
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "margin-top",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-top" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "margin-right",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-right" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "margin-bottom",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-bottom" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "margin-left",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-left" }
			},
			// padding
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "padding-top",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-top" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "padding-right",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-right" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "padding-bottom",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-bottom" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "padding-left",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-left" }
			},
			// border
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "border",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "border-top",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-top" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "border-right",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-right" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "border-bottom",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-bottom" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "border-left",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-left" }
			},
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "border-model",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "border-spacing",
					ValueToValue = new Dictionary<string, string>() {
						["collapsing"] = "0" } }
			},
			// text
			{ new OdtStyleAttrToCssProperty {
					OdtAttrName = "text-line-through-style",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "text-decoration",
					NameToValue = "line-through",
					ValueToValue = new Dictionary<string, string>
					{
						["solid"] = "solid",
						["dotted"] = "dotted",
						["dash"] = "dashed",
						["wave"] = "wavy",
						["dot-dash"] = "dashed",
						["dot-dot-dash"] = "dotted"
					} }
			},
			{ new OdtStyleAttrToCssProperty {
					OdtAttrName = "text-underline-style",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "text-decoration",
					NameToValue = "underline",
					ValueToValue = new Dictionary<string, string>
					{
						["solid"] = "solid",
						["dotted"] = "dotted",
						["dash"] = "dashed",
						["wave"] = "wavy",
						["dot-dash"] = "dashed",
						["dot-dot-dash"] = "dotted"
					} }
			},
			{ new OdtStyleAttrToCssProperty {
					OdtAttrName = "text-line-through-color",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "text-decoration-color",
					ValueToValue = new Dictionary<string, string>
					{
						["font-color"] = "inherit"
					} }
			},
			// writing
			{  new OdtStyleAttrToCssProperty {
					OdtAttrName = "writing-mode",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "writing-mode",
					ValueToValue = new Dictionary<string, string>() {
						["lr"] = "horizontal-tb",
						["lr-tb"] = "horizontal-tb",
						["rl"] = "horizontal-tb",
						["tb"] = "vertical-lr",
						["tb-rl"] = "vertical-rl" } }
			}
	};

		public static readonly List<string> LevelParent = new List<string>
		{
			{  "list" }
		};
	}
}