/*
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
				OdtTag = "image",
				HtmlTag = "img",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["width"] = "100%",
					["height"] = "auto" } } },
			{ new OdtTransTagToTag {
				OdtTag = "tab",
				HtmlTag = "span" } },
			{ new OdtTransTagToTag {
				OdtTag = "h",
				HtmlTag = "h",
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
				OdtTag = "p",
				HtmlTag = "p",
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
				OdtTag = "span",
				HtmlTag = "span",
				DefaultValue = " ",
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
			{ new OdtTransTagToTag { OdtTag = "paragraph", HtmlTag = "p" } },
			{ new OdtTransTagToTag { OdtTag = "a", HtmlTag = "a" } },
			{ new OdtTransTagToTag { OdtTag = "text-box", HtmlTag = "div" } },
			{ new OdtTransTagToTag {
				OdtTag = "table",
				HtmlTag = "table",
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
			{ new OdtTransTagToTag { OdtTag = "table-columns", HtmlTag = "tr" } },
			{ new OdtTransTagToTag { OdtTag = "table-column", HtmlTag = "th" } } ,
			{ new OdtTransTagToTag { OdtTag = "table-row", HtmlTag = "tr" } } ,
			{ new OdtTransTagToTag { OdtTag = "table-header-rows", HtmlTag = "tr" } } ,
			{ new OdtTransTagToTag {
				OdtTag = "table-cell",
				HtmlTag = "td",
				DefaultCssProperties = new Dictionary<string, string>(StrCompICIC)
				{
					["height"] = "1rem" } } },
			{ new OdtTransTagToTag {
				OdtTag = "list",
				HtmlTag = "ul",
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
				OdtTag = "list-item",
				HtmlTag = "li",
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
				OdtTag = "list-header",
				HtmlTag = "li",
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
				OdtTag = "bookmark-start",
				HtmlTag = "span" } },
			{ new OdtTransTagToTag {
				OdtTag = "bookmark",
				HtmlTag = "span" } },
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
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "table-properties" },
					LocalName = "width",
					AsPercentageTo = OdtTransStyleToStyle.RelativeToPage.Width,
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("rel-width", null) },
						new List<(string, string)>() { ("width", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "table-properties" },
					LocalName = "width",
					AsPercentageTo = OdtTransStyleToStyle.RelativeToPage.Width,
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("rel-width", null) },
						new List<(string, string)>() { ("width", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "table-properties" },
					LocalName = "height",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("height", null) },
						new List<(string, string)>() { ("height", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "table-column-properties" },
					LocalName = "width",
					AsPercentageTo = OdtTransStyleToStyle.RelativeToPage.Width,
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("column-width", null) },
						new List<(string, string)>() { ("width", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "table-column-properties" },
					LocalName = "height",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("column-height", null) },
						new List<(string, string)>() { ("height", null) }) } } },
			//align
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "table-cell-properties" },
					LocalName = "vertical-align",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("vertical-align", null) },
						new List<(string, string)>() { ("vertical-align", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "table-properties" },
					LocalName = "margin",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("align", "center") },
						new List<(string, string)>() { ("margin", "auto") } ),
						(new List<(string, string)>() { ("align", "margins") },
						new List<(string, string)>() { ("margin", "auto") } ),
						(new List<(string, string)>() { ("align", "start") },
						new List<(string, string)>() { ("margin-left", "0") } ),
						(new List<(string, string)>() { ("align", "end") },
						new List<(string, string)>() { ("margin-right", "0") } ),
						(new List<(string, string)>() { ("align", "left") },
						new List<(string, string)>() { ("margin-left", "0") } ),
						(new List<(string, string)>() { ("align", "right") },
						new List<(string, string)>() { ("margin-right", "0") }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties" },
					LocalName = "text-align",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("text-align", "start") },
						new List<(string, string)>() { ("text-align", "left") }),
						(new List<(string, string)>() { ("text-align", "end") },
						new List<(string, string)>() { ("text-align", "right") }) } } },
			// color
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties" },
					LocalName = "color",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("color", null) },
						new List<(string, string)>() { ("color", null) }) } } },
			// background
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties", "paragraph-properties", "table-cell-properties" },
					LocalName = "background-color",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("background-color", null) },
						new List<(string, string)>() { ("background-color", null) }) } } },
			// fonts
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "font-face" },
					LocalName = "font-family",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("font-family", null), ("font-family-generic", null) },
						new List<(string, string)>() { ("font-family", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties" },
					LocalName = "font-size",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("font-size", null) },
						new List<(string, string)>() { ("font-size", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties" },
					LocalName = "font-style",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("font-style", null) },
						new List<(string, string)>() { ("font-style", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties" },
					LocalName = "font-weight",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("font-weight", null) },
						new List<(string, string)>() { ("font-weight", null) }) } } },
			// line
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties.line-height" },
					LocalName = "line-height",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("line-height", null) },
						new List<(string, string)>() { ("line-height", null) }) } } },
			// margin
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-properties" },
					LocalName = "margin-top",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("margin-top", null) },
						new List<(string, string)>() { ("margin-top", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-properties" },
					LocalName = "margin-right",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("margin-right", null) },
						new List<(string, string)>() { ("margin-right", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-properties" },
					LocalName = "margin-bottom",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("margin-bottom", null) },
						new List<(string, string)>() { ("margin-bottom", null) }) } } },
			{ new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-properties" },
					LocalName = "margin-left",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("margin-left", null) },
						new List<(string, string)>() { ("margin-left", null) }) } } },
			// padding
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-cell-properties" },
					LocalName = "padding-top",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("padding-top", null) },
						new List<(string, string)>() { ("padding-top", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-cell-properties" },
					LocalName = "padding-right",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("padding-right", null) },
						new List<(string, string)>() { ("padding-right", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-cell-properties" },
					LocalName = "padding-bottom",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("padding-bottom", null) },
						new List<(string, string)>() { ("padding-bottom", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-cell-properties" },
					LocalName = "padding-left",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("padding-left", null) },
						new List<(string, string)>() { ("padding-left", null) }) } } },
			// border
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-cell-properties" },
					LocalName = "border",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("border", null) },
						new List<(string, string)>() { ("border", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-cell-properties" },
					LocalName = "border-top",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("border-top", null) },
						new List<(string, string)>() { ("border-top", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-cell-properties" },
					LocalName = "border-right",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("border-right", null) },
						new List<(string, string)>() { ("border-right", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-cell-properties" },
					LocalName = "border-bottom",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("border-bottom", null) },
						new List<(string, string)>() { ("border-bottom", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "paragraph-properties", "table-cell-properties" },
					LocalName = "border-left",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("border-left", null) },
						new List<(string, string)>() { ("border-left", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "table-properties.border-model" },
					LocalName = "border-spacing",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("border-model", null) },
						new List<(string, string)>() { ("border-spacing", null) }),
						(new List<(string, string)>() { ("border-model", "collapsing") },
						new List<(string, string)>() { ("border-spacing", "0") }) } } },
			// text
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties" },
					LocalName = "text-decoration-style",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("text-line-through-style", "solid") },
						new List<(string, string)>() { ("text-decoration-style", "solid"), ("text-decoration-line", "line-through") } ),
						(new List<(string, string)>() { ("text-line-through-style", "dotted") },
						new List<(string, string)>() { ("text-decoration-style", "dotted"), ("text-decoration-line", "line-through") } ),
						(new List<(string, string)>() { ("text-line-through-style", "dash") },
						new List<(string, string)>() { ("text-decoration-style", "dashed"), ("text-decoration-line", "line-through") } ),
						(new List<(string, string)>() { ("text-line-through-style", "wave") },
						new List<(string, string)>() { ("text-decoration-style", "wavy"), ("text-decoration-line", "line-through") } ),
						(new List<(string, string)>() { ("text-line-through-style", "dot-dash") },
						new List<(string, string)>() { ("text-decoration-style", "dashed"), ("text-decoration-line", "line-through") } ),
						(new List<(string, string)>() { ("text-line-through-style", "dot-dot-dash") },
						new List<(string, string)>() { ("text-decoration-style", "dotted"), ("text-decoration-line", "line-through") } ) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties.text-underline-style" },
					LocalName = "text-decoration-style",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("text-underline-style", "solid") },
						new List<(string, string)>() { ("text-decoration-style", "solid"), ("text-decoration-line", "underline") } ),
						(new List<(string, string)>() { ("text-underline-style", "dotted") },
						new List<(string, string)>() { ("text-decoration-style", "dotted"), ("text-decoration-line", "underline") } ),
						(new List<(string, string)>() { ("text-underline-style", "dash") },
						new List<(string, string)>() { ("text-decoration-style", "dashed"), ("text-decoration-line", "underline") } ),
						(new List<(string, string)>() { ("text-underline-style", "wave") },
						new List<(string, string)>() { ("text-decoration-style", "wavy"), ("text-decoration-line", "underline") } ),
						(new List<(string, string)>() { ("text-underline-style", "dot-dash") },
						new List<(string, string)>() { ("text-decoration-style", "dashed"), ("text-decoration-line", "underline") } ),
						(new List<(string, string)>() { ("text-underline-style", "dot-dot-dash") },
						new List<(string, string)>() { ("text-decoration-style", "dotted"), ("text-decoration-line", "underline") } ) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties" },
					LocalName = "max-width",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("width", null) },
						new List<(string, string)>() { ("max-width", null) }) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties" },
					LocalName = "text-decoration-color",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("text-line-through-color", null) },
						new List<(string, string)>() { ("text-decoration-color", null) } ),
						(new List<(string, string)>() { ("text-line-through-color", "font-color") },
						new List<(string, string)>() { ("text-decoration-color", "inherit") } ) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties" },
					LocalName = "text-decoration-color",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("text-underline-color", null) },
						new List<(string, string)>() { ("text-decoration-color", null) } ),
						(new List<(string, string)>() { ("text-underline-color", "font-color") },
						new List<(string, string)>() { ("text-decoration-color", "inherit") } ) } } },
			{
				new OdtTransStyleToStyle {
					PropNames = new List<string> { "text-properties.text-transform" },
					LocalName = "text-transform",
					Values = new List<(List<(string, string)>, List<(string, string)>)>() {
						(new List<(string, string)>() { ("text-transform", null) },
						new List<(string, string)>() { ("text-transform", null) }) } } }

			// TODO table-cell-properties.glyph-orientation-vertical -> "writing-mode / 0 -> vertical-rl
		};
	}
}