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
			{ new OdtTagToHtml { OdtName = "tab", HtmlName = "p" } },
			{ new OdtTagToHtml { OdtName = "span", HtmlName = "span" } },
			{ new OdtTagToHtml { OdtName = "paragraph", HtmlName = "p" } },
			{ new OdtTagToHtml { OdtName = "graphic", HtmlName = "img" } },
			{ new OdtTagToHtml { OdtName = "image", HtmlName = "img" } },
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

		public static readonly Dictionary<string, string> OdtNodeAttrToHtmlNodeInlineStyle = new Dictionary<string, string>()
		{
			["frame.width"] = "max-width",
			["frame.height"] = "max-height",
		};

		public static readonly List<OdtStyleToCss> OdtStyleToCssStyle = new List<OdtStyleToCss>()
		{
			// width/height
			{  new OdtStyleToCss {
					OdtName = "width", CssName = "max-width",
					OdtTypes = new List<string> { "table-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "rel-width", CssName = "width",
					OdtTypes = new List<string> { "table-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "column-width", CssName = "max-width",
					OdtTypes = new List<string> { "table-column-properties" } }
			},
			// align
			{  new OdtStyleToCss {
					OdtName = "vertical-align", CssName = "vertical-align",
					OdtTypes = new List<string> { "table-cell-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "align", CssName = "align",
					OdtTypes = new List<string> { "table-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "text-align", CssName = "text-align",
					OdtTypes = new List<string> { "paragraph-properties" } }
			},
			// color
			{  new OdtStyleToCss {
					OdtName = "color", CssName = "color",
					OdtTypes = new List<string> { "text-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "background-color", CssName = "background-color",
					OdtTypes = new List<string> { "paragraph-properties", "table-cell-properties" } } // "text-properties"
			},
			// fonts
			{  new OdtStyleToCss {
					OdtName = "font-name", CssName = "font-family",
					OdtTypes = new List<string> { "text-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "font-size", CssName = "font-size",
					OdtTypes = new List<string> { "text-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "font-style", CssName = "font-style",
					OdtTypes = new List<string> { "text-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "line-height", CssName = "line-height",
					OdtTypes = new List<string> { "text-properties" } }
			},
			// margin
			{  new OdtStyleToCss {
					OdtName = "margin-top", CssName = "margin-top",
					OdtTypes = new List<string> { "paragraph-properties", "table-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "margin-right", CssName = "margin-right",
					OdtTypes = new List<string> { "paragraph-properties", "table-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "margin-bottom", CssName = "margin-bottom",
					OdtTypes = new List<string> { "paragraph-properties", "table-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "margin-left", CssName = "margin-left",
					OdtTypes = new List<string> { "paragraph-properties", "table-properties" } }
			},
			// padding
			{  new OdtStyleToCss {
					OdtName = "padding-top", CssName = "padding-top",
					OdtTypes = new List<string> { "paragraph-properties", "table-cell-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "padding-right", CssName = "padding-right",
					OdtTypes = new List<string> { "paragraph-properties", "table-cell-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "padding-bottom", CssName = "padding-bottom",
					OdtTypes = new List<string> { "paragraph-properties", "table-cell-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "padding-left", CssName = "padding-left",
					OdtTypes = new List<string> { "paragraph-properties", "table-cell-properties" } }
			},
			// border
			{  new OdtStyleToCss {
					OdtName = "border-top", CssName = "border-top",
					OdtTypes = new List<string> { "paragraph-properties", "table-cell-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "border-right", CssName = "border-right",
					OdtTypes = new List<string> { "paragraph-properties", "table-cell-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "border-bottom", CssName = "border-bottom",
					OdtTypes = new List<string> { "paragraph-properties", "table-cell-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "border-left", CssName = "border-left",
					OdtTypes = new List<string> { "paragraph-properties", "table-cell-properties" } }
			},
			{  new OdtStyleToCss {
					OdtName = "border-model", CssName = "border-spacing",
					OdtTypes = new List<string> { "table-properties" },
					Values = new Dictionary<string, string>() {
						["collapsing"] = "0"} }
			},
			// other
			{  new OdtStyleToCss {
					OdtName = "writing-mode", CssName = "writing-mode",
					OdtTypes = new List<string> { "table-properties" },
					Values = new Dictionary<string, string>() {
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