/*
DDDN.OdtToHtml.Transformation.OdtTrans
Copyright(C) 2017-2019 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using static DDDN.OdtToHtml.Transformation.OdtTransStyleToStyle;

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
			{ new OdtTransStyleToStyle("width", new string[] { "table-properties" }, RelativeToPage.Width)
                .Add(new OdtStyleValues().Add("rel-width", null), new OdtStyleValues().Add("width", null))
            },
            { new OdtTransStyleToStyle("height", new string[] { "table-properties" })
                .Add(new OdtStyleValues().Add("height", null), new OdtStyleValues().Add("height", null))
            },
            { new OdtTransStyleToStyle("column-width", new string[] { "table-column-properties" })
                .Add(new OdtStyleValues().Add("column-width", null), new OdtStyleValues().Add("width", null))
            },
            { new OdtTransStyleToStyle("column-height", new string[] { "table-column-properties" })
                .Add(new OdtStyleValues().Add("column-height", null), new OdtStyleValues().Add("height", null))
            },                                                                            
            //align
            { new OdtTransStyleToStyle("vertical-align", new string[] { "table-cell-properties" })
                .Add(new OdtStyleValues().Add("vertical-align", null), new OdtStyleValues().Add("vertical-align", null))
            },
            { new OdtTransStyleToStyle("align", new string[] { "table-properties" })
                .Add(new OdtStyleValues().Add("align", "center"), new OdtStyleValues().Add("margin", "auto"))
                .Add(new OdtStyleValues().Add("align", "margins"), new OdtStyleValues().Add("margin", "auto"))
                .Add(new OdtStyleValues().Add("align", "start"), new OdtStyleValues().Add("margin-left", "0"))
                .Add(new OdtStyleValues().Add("align", "end"), new OdtStyleValues().Add("margin-right", "0"))
                .Add(new OdtStyleValues().Add("align", "left"), new OdtStyleValues().Add("margin-left", "0"))
                .Add(new OdtStyleValues().Add("align", "right"), new OdtStyleValues().Add("margin-right", "0"))
            },
            { new OdtTransStyleToStyle("text-align", new string[] { "paragraph-properties" })
                .Add(new OdtStyleValues().Add("text-align", "start"), new OdtStyleValues().Add("text-align", "left"))
                .Add(new OdtStyleValues().Add("text-align", "end"), new OdtStyleValues().Add("text-align", "right"))
                .Add(new OdtStyleValues().Add("align", "start"), new OdtStyleValues().Add("margin-left", "0"))
                .Add(new OdtStyleValues().Add("align", "end"), new OdtStyleValues().Add("margin-right", "0"))
                .Add(new OdtStyleValues().Add("align", "left"), new OdtStyleValues().Add("margin-left", "0"))
                .Add(new OdtStyleValues().Add("align", "right"), new OdtStyleValues().Add("margin-right", "0"))
            },
            // color
            { new OdtTransStyleToStyle("color", new string[] { "text-properties" })
                .Add(new OdtStyleValues().Add("color", null), new OdtStyleValues().Add("color", null))
            },
            // background
            { new OdtTransStyleToStyle("background-color", new string[] { "text-properties", "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("background-color", null) , new OdtStyleValues().Add("background-color", null))
            },
            // fonts
            { new OdtTransStyleToStyle("font-family", new string[] { "font-face" })
                .Add(new OdtStyleValues().Add("font-family", null).Add("font-family-generic", null), new OdtStyleValues().Add("font-family", null))
            },
            { new OdtTransStyleToStyle("font-size", new string[] { "text-properties" })
                .Add(new OdtStyleValues().Add("font-size", null), new OdtStyleValues().Add("font-size", null))
            },
            { new OdtTransStyleToStyle("font-style", new string[] { "text-properties" })
                .Add(new OdtStyleValues().Add("font-style", null), new OdtStyleValues().Add("font-style", null))
            },
            { new OdtTransStyleToStyle("font-weight", new string[] { "text-properties" })
                .Add(new OdtStyleValues().Add("font-weight", null), new OdtStyleValues().Add("font-weight", null))
            },
            // line
            { new OdtTransStyleToStyle("line-height", new string[] { "text-properties.line-height" })
                .Add(new OdtStyleValues().Add("line-height", null), new OdtStyleValues().Add("line-height", null))
            },
             // margin
            { new OdtTransStyleToStyle("margin-top", new string[] { "paragraph-properties", "table-properties" })
                .Add(new OdtStyleValues().Add("margin-top", null), new OdtStyleValues().Add("margin-top", null))
            },
            { new OdtTransStyleToStyle("margin-right", new string[] { "paragraph-properties", "table-properties" })
                .Add(new OdtStyleValues().Add("margin-right", null), new OdtStyleValues().Add("margin-right", null))
            },
            { new OdtTransStyleToStyle("margin-bottom", new string[] { "paragraph-properties", "table-properties" })
                .Add(new OdtStyleValues().Add("margin-bottom", null), new OdtStyleValues().Add("margin-bottom", null))
            },
            { new OdtTransStyleToStyle("margin-left", new string[] { "paragraph-properties", "table-properties" })
                .Add(new OdtStyleValues().Add("margin-left", null), new OdtStyleValues().Add("margin-left", null))
            },
            // padding
            { new OdtTransStyleToStyle("padding-top", new string[] { "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("padding-top", null), new OdtStyleValues().Add("padding-top", null))
            },
            { new OdtTransStyleToStyle("padding-right", new string[] { "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("padding-right", null), new OdtStyleValues().Add("padding-right", null))
            },
            { new OdtTransStyleToStyle("padding-bottom", new string[] { "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("padding-bottom", null), new OdtStyleValues().Add("padding-bottom", null))
            },
            { new OdtTransStyleToStyle("padding-left", new string[] { "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("padding-left", null), new OdtStyleValues().Add("padding-left", null))
            },
            // border
            { new OdtTransStyleToStyle("border", new string[] { "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("border", null), new OdtStyleValues().Add("border", null))
            },
            { new OdtTransStyleToStyle("border-top", new string[] { "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("border-top", null), new OdtStyleValues().Add("border-top", null))
            },
            { new OdtTransStyleToStyle("border-right", new string[] { "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("border-right", null), new OdtStyleValues().Add("border-right", null))
            },
            { new OdtTransStyleToStyle("border-left", new string[] { "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("border-left", null), new OdtStyleValues().Add("border-left", null))
            },
            { new OdtTransStyleToStyle("border-bottom", new string[] { "paragraph-properties", "table-cell-properties" })
                .Add(new OdtStyleValues().Add("border-bottom", null), new OdtStyleValues().Add("border-bottom", null))
            },
            { new OdtTransStyleToStyle("border-spacing", new string[] { "table-properties.border-model" })
                .Add(new OdtStyleValues().Add("border-model", null).Add("border-spacing", null), new OdtStyleValues().Add("border-model", "collapsing").Add("border-spacing", "0"))
            },
            // text
            { new OdtTransStyleToStyle("text-decoration-style", new string[] { "text-properties" })
                .Add(new OdtStyleValues().Add("text-line-through-style", "solid"),
                    new OdtStyleValues().Add("text-decoration-style", "solid").Add("text-decoration-line", "line-through"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "dotted"),
                    new OdtStyleValues().Add("text-decoration-style", "dotted").Add("text-decoration-line", "line-through"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "dash"),
                    new OdtStyleValues().Add("text-decoration-style", "dashed").Add("text-decoration-line", "line-through"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "wave"),
                    new OdtStyleValues().Add("text-decoration-style", "wavy").Add("text-decoration-line", "line-through"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "dot-dash"),
                    new OdtStyleValues().Add("text-decoration-style", "dashed").Add("text-decoration-line", "line-through"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "dot-dot-dash"),
                    new OdtStyleValues().Add("text-decoration-style", "dotted").Add("text-decoration-line", "line-through"))
            },
            { new OdtTransStyleToStyle("text-decoration-style", new string[] { "text-properties.text-underline-style" })
                .Add(new OdtStyleValues().Add("text-line-through-style", "solid"),
                    new OdtStyleValues().Add("text-decoration-style", "solid").Add("text-decoration-line", "underline"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "dotted"),
                    new OdtStyleValues().Add("text-decoration-style", "dotted").Add("text-decoration-line", "underline"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "dash"),
                    new OdtStyleValues().Add("text-decoration-style", "dashed").Add("text-decoration-line", "underline"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "wave"),
                    new OdtStyleValues().Add("text-decoration-style", "wavy").Add("text-decoration-line", "underline"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "dot-dash"),
                    new OdtStyleValues().Add("text-decoration-style", "dashed").Add("text-decoration-line", "underline"))
                .Add(new OdtStyleValues().Add("text-line-through-style", "dot-dot-dash"),
                    new OdtStyleValues().Add("text-decoration-style", "dotted").Add("text-decoration-line", "underline"))
            },
            { new OdtTransStyleToStyle("max-width", new string[] { "text-properties" })
                .Add(new OdtStyleValues().Add("width", null), new OdtStyleValues().Add("max-width", null))
            },
            { new OdtTransStyleToStyle("text-decoration-color", new string[] { "text-properties" })
                .Add(new OdtStyleValues().Add("text-line-through-color", null), new OdtStyleValues().Add("text-decoration-color", null))
                .Add(new OdtStyleValues().Add("text-line-through-color", "font-color"), new OdtStyleValues().Add("text-decoration-color", "inherit"))
            },
            { new OdtTransStyleToStyle("text-decoration-color", new string[] { "text-properties" })
                .Add(new OdtStyleValues().Add("text-underline-color", null), new OdtStyleValues().Add("text-decoration-color", null))
                .Add(new OdtStyleValues().Add("text-underline-color", "font-color"), new OdtStyleValues().Add("text-decoration-color", "inherit"))
            },
            { new OdtTransStyleToStyle("text-transform", new string[] { "text-properties.text-transform" })
                .Add(new OdtStyleValues().Add("text-transform", null), new OdtStyleValues().Add("text-transform", null))
            },
        };
        // TODO table-cell-properties.glyph-orientation-vertical -> "writing-mode / 0 -> vertical-rl
    };
}