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

using System;
using System.Collections.Generic;

namespace DDDN.OdtToHtml
{
	public static class OdtTrans
	{
		public static readonly Dictionary<string, string> TagToTag = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase)
		{
			["h"] = "p",
			["p"] = "p",
			["span"] = "span",
			["paragraph"] = "p",
			["s"] = "span",
			["a"] = "a",
			["text-box"] = "div",
			["table"] = "table",
			["table-columns"] = "tr",
			["table-column"] = "th",
			["table-row"] = "tr",
			["table-cell"] = "td",
			["list"] = "ul",
			["list-item"] = "li",
			["automatic-styles"] = "style"
		};

		public static readonly Dictionary<string, string> AttrNameToAttrName = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase)
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

		public static readonly List<OdtStyleToStyle> StyleToStyle = new List<OdtStyleToStyle>()
		{
			// width/height
			{  new OdtStyleToStyle {
					OdtAttrName = "width",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "width",
					AsPercentage = OdtStyleToStyle.RelativeTo.Width }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "height",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "height",
					AsPercentage = OdtStyleToStyle.RelativeTo.Height }
			},
			//{  new OdtStyleToStyle {
			//		OdtAttrName = "rel-width",
			//		StyleTypes = new List<string> { "table-properties" },
			//		CssPropName = "width" }
			//},
			{  new OdtStyleToStyle {
					OdtAttrName = "column-width",
					StyleTypes = new List<string> { "table-column-properties" },
					CssPropName = "width",
					AsPercentage = OdtStyleToStyle.RelativeTo.Width }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "column-height",
					StyleTypes = new List<string> { "table-column-properties" },
					CssPropName = "height",
					AsPercentage = OdtStyleToStyle.RelativeTo.Height }
			},
			// align
			{  new OdtStyleToStyle {
					OdtAttrName = "vertical-align",
					StyleTypes = new List<string> { "table-cell-properties" },
					CssPropName = "vertical-align" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "align",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "margin",
					ValueToValue = new Dictionary<string, string>() {
						["center"] = "auto",
						["start"] = "0",
						["end"] = "0" } }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "text-align",
					StyleTypes = new List<string> { "paragraph-properties" },
					CssPropName = "text-align",
					ValueToValue = new Dictionary<string, string>() {
						["start"] = "left",
						["end"] = "right" } }
			},
			// color
			{  new OdtStyleToStyle {
					OdtAttrName = "color",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "color" }
			},
			// background
			{  new OdtStyleToStyle {
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
			{  new OdtStyleToStyle {
					OdtAttrName = "font-size",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-size" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "font-style",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-style" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "font-weight",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "font-weight" }
			},
			// line
			{  new OdtStyleToStyle {
					OdtAttrName = "line-height",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "line-height" }
			},
			// margin
			{  new OdtStyleToStyle {
					OdtAttrName = "margin-top",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-top" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "margin-right",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-right" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "margin-bottom",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-bottom" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "margin-left",
					StyleTypes = new List<string> { "paragraph-properties", "table-properties" },
					CssPropName = "margin-left" }
			},
			// padding
			{  new OdtStyleToStyle {
					OdtAttrName = "padding-top",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-top" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "padding-right",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-right" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "padding-bottom",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-bottom" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "padding-left",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "padding-left" }
			},
			// border
			{  new OdtStyleToStyle {
					OdtAttrName = "border",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "border-top",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-top" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "border-right",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-right" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "border-bottom",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-bottom" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "border-left",
					StyleTypes = new List<string> { "paragraph-properties", "table-cell-properties" },
					CssPropName = "border-left" }
			},
			{  new OdtStyleToStyle {
					OdtAttrName = "border-model",
					StyleTypes = new List<string> { "table-properties" },
					CssPropName = "border-spacing",
					ValueToValue = new Dictionary<string, string>() {
						["collapsing"] = "0" } }
			},
			// text
			{ new OdtStyleToStyle {
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
			{ new OdtStyleToStyle {
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
			{ new OdtStyleToStyle {
					OdtAttrName = "text-line-through-color",
					StyleTypes = new List<string> { "text-properties" },
					CssPropName = "text-decoration-color",
					ValueToValue = new Dictionary<string, string>
					{
						["font-color"] = "inherit"
					} }
			},
			//// writing
			//{  new OdtStyleToStyle {
			//		OdtAttrName = "writing-mode",
			//		StyleTypes = new List<string> { "table-properties" },
			//		CssPropName = "writing-mode",
			//		ValueToValue = new Dictionary<string, string>() {
			//			["lr"] = "horizontal-tb",
			//			["lr-tb"] = "horizontal-tb",
			//			["rl"] = "horizontal-tb",
			//			["tb"] = "vertical-lr",
			//			["tb-rl"] = "vertical-rl" } }
			//}
	};

		public static readonly List<string> LevelParent = new List<string>
		{
			{  "list" }
		};
	}
}