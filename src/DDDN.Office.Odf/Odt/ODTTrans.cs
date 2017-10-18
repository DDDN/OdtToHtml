/*
DDDN.Office.Odf.Odt.OdtTrans
Copyright(C) 2017 Lukasz Jaskiewicz(lukasz @jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System.Collections.Generic;

namespace DDDN.Office.Odf.Odt
{
	public static class OdtTrans
	{
		public static readonly List<OdfTagToHtml> Tags = new List<OdfTagToHtml>()
		{
			{
				new OdfTagToHtml
				{
					OdfName = "h",
					HtmlName = "p"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "p",
					HtmlName = "p"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "tab",
					HtmlName = "p"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "span",
					HtmlName = "span"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "paragraph",
					HtmlName = "p"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "graphic",
					HtmlName = "img"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "image",
					HtmlName = "img"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "s",
					HtmlName = "span"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "a",
					HtmlName = "a"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "frame",
					HtmlName = "div"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "text-box",
					HtmlName = "div"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "table",
					HtmlName = "table"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "table-columns",
					HtmlName = "tr"
				}
			},
			{
				new OdfTagToHtml
				{
					OdfName = "table-column",
					HtmlName = "th"
				}
			}
			,
			{
				new OdfTagToHtml
				{
					OdfName = "table-row",
					HtmlName = "tr"
				}
			}
			,
			{
				new OdfTagToHtml
				{
					OdfName = "table-cell",
					HtmlName = "td"
				}
			}
			,
			{
				new OdfTagToHtml
				{
					OdfName = "list",
					HtmlName = "ul"
				}
			}
			,
			{
				new OdfTagToHtml
				{
					OdfName = "list-item",
					HtmlName = "li"
				}
			}
			,
			{
				new OdfTagToHtml
				{
					OdfName = "automatic-styles",
					HtmlName = "style"
				}
			}
			,
			{
				new OdfTagToHtml
				{
					OdfName = "space-before",
					HtmlName = "margin-left"
				}
			}
		};

		public static readonly Dictionary<string, string> Attrs = new Dictionary<string, string>()
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

		public static readonly Dictionary<string, string> StyleAttr = new Dictionary<string, string>()
		{
			["frame.width"] = "max-width",
			["frame.height"] = "max-height",
		};

		public static readonly List<OdfStyleToCss> Css = new List<OdfStyleToCss>()
		{
			{
				 new OdfStyleToCss
				 {
					  OdfName = "border-model",
					  CssName = "border-spacing",
					  Values = new Dictionary<string, string>()
					  {
							["collapsing"] = "0"
					  }
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "writing-mode",
					  CssName = "writing-mode",
					  Values = new Dictionary<string, string>()
					  {
							["lr"] = "horizontal-tb",
							["lr-tb"] = "horizontal-tb",
							["rl"] = "horizontal-tb",
							["tb"] = "vertical-lr",
							["tb-rl"] = "vertical-rl"
					  }
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "hyphenate",
					  CssName = "hyphens",
					  Values = new Dictionary<string, string>()
					  {
							["false"] = "none"
					  }
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "font-name",
					  CssName = "font-family"
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "page-width",
					  CssName = "width"
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "rel-width",
					  CssName = "width"
				 }
			},
			//{
			//	 new OdfStyleToCss
			//	 {
			//		  OdfName = "tab-stop-distance",
			//		  CssName = "margin"
			//	 }
			//},
			{ new OdfStyleToCss { OdfName = "page-height" } },
			{ new OdfStyleToCss { OdfName = "num-format" } },
			{ new OdfStyleToCss { OdfName = "print-orientation" } },
			{ new OdfStyleToCss { OdfName = "keep-with-next" } },
			{ new OdfStyleToCss { OdfName = "keep-together" } },
			{ new OdfStyleToCss { OdfName = "widows" } },
			{ new OdfStyleToCss { OdfName = "language" } },
			{ new OdfStyleToCss { OdfName = "country" } },
			{ new OdfStyleToCss { OdfName = "display-name" } },
			{ new OdfStyleToCss { OdfName = "orphans" } },
			{ new OdfStyleToCss { OdfName = "fill" } },
			{ new OdfStyleToCss { OdfName = "fill-color" } },
			{ new OdfStyleToCss { OdfName = "line-break" } },
			{ new OdfStyleToCss { OdfName = "punctuation-wrap" } },
			{ new OdfStyleToCss { OdfName = "number-lines" } },
			{ new OdfStyleToCss { OdfName = "text-autospace" } },
			{ new OdfStyleToCss { OdfName = "snap-to-layout-grid" } },
			{ new OdfStyleToCss { OdfName = "text-autospace" } },
			{ new OdfStyleToCss { OdfName = "tab-stop-distance" } },
			{ new OdfStyleToCss { OdfName = "use-window-font-color" } },
			{ new OdfStyleToCss { OdfName = "letter-kerning" } },
			{ new OdfStyleToCss { OdfName = "text-scale" } },
			{ new OdfStyleToCss { OdfName = "text-position" } },
			{ new OdfStyleToCss { OdfName = "font-relief" } },
			{ new OdfStyleToCss { OdfName = "horizontal-pos" } },
			{ new OdfStyleToCss { OdfName = "horizontal-rel" } },
			{ new OdfStyleToCss { OdfName = "vertical-pos" } },
			{ new OdfStyleToCss { OdfName = "vertical-rel" } },
			{ new OdfStyleToCss { OdfName = "break-before" } },
			{ new OdfStyleToCss { OdfName = "font-size-complex" } },
			{ new OdfStyleToCss { OdfName = "font-size-asian" } },
			{ new OdfStyleToCss { OdfName = "font-weight-complex" } },
			{ new OdfStyleToCss { OdfName = "font-weight-asian" } },
			{ new OdfStyleToCss { OdfName = "font-name-complex" } },
			{ new OdfStyleToCss { OdfName = "font-name-asian" } },
			{ new OdfStyleToCss { OdfName = "font-style-asian" } },
			{ new OdfStyleToCss { OdfName = "font-style-complex" } },
			{ new OdfStyleToCss { OdfName = "language-asian" } },
			{ new OdfStyleToCss { OdfName = "language-complex" } },
			{ new OdfStyleToCss { OdfName = "country-asian" } },
			{ new OdfStyleToCss { OdfName = "country-complex" } },
		};
	}
}
