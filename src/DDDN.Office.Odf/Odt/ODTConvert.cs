/*
* DDDN.Office.Odf.Odt.ODTConvert
* 
* Copyright(C) 2017 Lukasz Jaskiewicz
* Author: Lukasz Jaskiewicz (lukasz@jaskiewicz.de, devdone@outlook.com)
*
* This program is free software; you can redistribute it and/or modify it under the terms of the
* GNU General Public License as published by the Free Software Foundation; version 2 of the License.
*
* This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
* warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License along with this program; if not, write
* to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DDDN.Office.Odf.Odt
{
    public class ODTConvert : IODTConvert
    {
        private XDocument ContentXDoc;
        private XDocument StylesXDoc;

        private static readonly Dictionary<string, string> HtmlTagsTrans = new Dictionary<string, string>()
        {
            ["text"] = "article",
            ["h"] = "p",
            ["p"] = "p",
            ["span"] = "span",
            ["paragraph"] = "p",
            ["graphic"] = "img",
            ["s"] = "span",
            ["a"] = "a",
            ["table"] = "table",
            ["table-columns"] = "tr",
            ["table-column"] = "th",
            ["table-row"] = "tr",
            ["table-cell"] = "td",
            ["list"] = "ul",
            ["list-item"] = "li",
            ["automatic-styles"] = "style"
        };

        private static readonly Dictionary<string, string> HtmlAttrTrans = new Dictionary<string, string>()
        {
            ["style-name"] = "class",
            ["href"] = "href",
            ["target-frame-name"] = "target"
        };

        private static readonly List<OdfStyleToCss> CssTrans = new List<OdfStyleToCss>()
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

        public ODTConvert(IODTFile odtFile)
        {
            if (odtFile == null)
            {
                throw new ArgumentNullException(nameof(odtFile));
            }

            ContentXDoc = odtFile.GetZipArchiveEntryAsXDocument("content.xml");
            StylesXDoc = odtFile.GetZipArchiveEntryAsXDocument("styles.xml");
        }

        public ODTConvertData Convert()
        {
            var data = new ODTConvertData
            {
                Html = GetHtml(),
                Css = GetCss(),
                FirstHeaderText = GetFirstHeaderText(),
                FirstParagraphHtml = GetFirstParagraphHtml()
            };

            return data;
        }

        private string GetFirstParagraphHtml()
        {
            var firstParagraph = ContentXDoc.Root
                .Elements(XName.Get("body", ODFXmlNamespaces.Office))
                .Elements(XName.Get("text", ODFXmlNamespaces.Office))
                .Elements(XName.Get("p", ODFXmlNamespaces.Text))
                .FirstOrDefault();

            if (firstParagraph != default(XElement))
            {
                var htmlEle = new XElement(HtmlTagsTrans[firstParagraph.Name.LocalName]);
                ContentNodesWalker(firstParagraph.Nodes(), htmlEle);
                var html = htmlEle.ToString(SaveOptions.DisableFormatting);
                return html;
            }
            else
            {
                return string.Empty;
            }
        }

        private string GetCss()
        {
            List<IOdfStyle> Styles = new List<IOdfStyle>();

            //var fontFaceDeclarations = ContentXDoc.Root
            //     .Elements(XName.Get("font-face-decls", ODFXmlNamespaces.Office))
            //     .Elements()
            //     .Where(p => p.Name.Equals(XName.Get("font-face", ODFXmlNamespaces.Style)));
            //StylesWalker(fontFaceDeclarations, Styles);

            var automaticStyles = ContentXDoc.Root
                 .Elements(XName.Get("automatic-styles", ODFXmlNamespaces.Office))
                 .Elements()
                 .Where(p => p.Name.Equals(XName.Get("style", ODFXmlNamespaces.Style)));
            StylesWalker(automaticStyles, Styles);

            var defaultStyles = StylesXDoc.Root
                 .Elements(XName.Get("styles", ODFXmlNamespaces.Office))
                 .Elements()
                 .Where(p => p.Name.Equals(XName.Get("default-style", ODFXmlNamespaces.Style)));
            StylesWalker(defaultStyles, Styles);

            var styles = StylesXDoc.Root
                 .Elements(XName.Get("styles", ODFXmlNamespaces.Office))
                 .Elements()
                 .Where(p => p.Name.Equals(XName.Get("style", ODFXmlNamespaces.Style)));
            StylesWalker(styles, Styles);

            var css = RenderCss(Styles);
            return css;
        }

        private string RenderCss(List<IOdfStyle> styles)
        {
            var builder = new StringBuilder(2048);

            foreach (var style in styles)
            {
                if (style.Type.Equals("default-style", StringComparison.InvariantCultureIgnoreCase)
                     || string.IsNullOrWhiteSpace(style.Name))
                {
                    builder.Append($"{Environment.NewLine}article > {HtmlTagsTrans[style.Family]} {{{Environment.NewLine}");
                }
                else
                {
                    builder.Append($"{Environment.NewLine}.{style.Name} {{{Environment.NewLine}");
                }

                foreach (var attr in style.Attrs)
                {
                    TransformStyleAttr(builder, style, attr);
                }

                foreach (var props in style.PropAttrs.Values)
                {
                    foreach (var propAttr in props)
                    {
                        TransformStyleAttr(builder, style, propAttr);
                    }
                }

                builder.Append("}");
            }

            return builder.ToString();
        }

        private static void TransformStyleAttr(StringBuilder builder, IOdfStyle style, IOdfStyleAttr attr)
        {
            var trans = CssTrans
                .Where(p => p.OdfName.Equals(attr.Name))
                .FirstOrDefault();

            if (trans != default(OdfStyleToCss))
            {
                if (string.IsNullOrWhiteSpace(trans.CssName))
                {
                    return;
                }
                else
                {
                    var attrName = trans.CssName;
                    var attrVal = "";

                    if (trans.Values != null && trans.Values.ContainsKey(attr.Value))
                    {
                        attrVal = trans.Values[attr.Value];
                    }
                    else
                    {
                        attrVal = attr.Value;
                    }

                    builder.Append($"{attrName}: {attrVal};{Environment.NewLine}");
                }
            }
            else
            {
                builder.Append($"{attr.Name}: {attr.Value};{Environment.NewLine}");
            }
        }

        private void StylesWalker(IEnumerable<XElement> elements, List<IOdfStyle> styles)
        {
            foreach (var ele in elements)
            {
                IOdfStyle style = new OdfStyle(ele);
                styles.Add(style);
                StylePropertyWalker(ele.Elements(), style);
            }
        }

        private void StylePropertyWalker(IEnumerable<XElement> elements, IOdfStyle style)
        {
            foreach (var ele in elements.Where(p => p.Name.LocalName.EndsWith("-properties")))
            {
                style.AddPropertyAttributes(ele);
                StylePropertyWalker(ele.Elements(), style);
            }
        }

        private string GetHtml()
        {
            var contentEle = ContentXDoc.Root
                    .Elements(XName.Get("body", ODFXmlNamespaces.Office))
                    .Elements(XName.Get("text", ODFXmlNamespaces.Office))
                    .First();

            var htmlEle = new XElement(HtmlTagsTrans[contentEle.Name.LocalName]);
            ContentNodesWalker(contentEle.Nodes(), htmlEle);

            var html = htmlEle.ToString(SaveOptions.DisableFormatting);
            return html;
        }

        private string GetFirstHeaderText()
        {
            var firstHeader = ContentXDoc.Root
                .Elements(XName.Get("body", ODFXmlNamespaces.Office))
                .Elements(XName.Get("text", ODFXmlNamespaces.Office))
                .Elements(XName.Get("h", ODFXmlNamespaces.Text))
                .FirstOrDefault();

            if (firstHeader != default(XElement))
            {
                return ODTReader.GetValue(firstHeader);
            }
            else
            {
                return String.Empty;
            }
        }

        private void ContentNodesWalker(IEnumerable<XNode> odNode, XElement htmlElement)
        {
            var childHtmlEle = htmlElement;

            foreach (var node in odNode)
            {
                if (node.NodeType == XmlNodeType.Text)
                {
                    var textNode = node as XText;
                    childHtmlEle.SetValue(childHtmlEle.Value + textNode.Value);
                }
                else if (node.NodeType == XmlNodeType.Element)
                {
                    var elementNode = node as XElement;

                    if (elementNode.Name.Equals(XName.Get("s", ODFXmlNamespaces.Text)))
                    {
                        AddNbsp(elementNode, htmlElement);
                    }
                    else if (HtmlTagsTrans.TryGetValue(elementNode.Name.LocalName, out string htmlTag))
                    {
                        childHtmlEle = new XElement(htmlTag);
                        CopyAttributes(elementNode, childHtmlEle);
                        htmlElement.Add(childHtmlEle);
                        ContentNodesWalker(elementNode.Nodes(), childHtmlEle);
                    }
                }
            }
        }

        private static void AddNbsp(XElement odElement, XElement htmlElement)
        {
            var spacesValue = odElement.Attribute(XName.Get("c", ODFXmlNamespaces.Text))?.Value;
            int.TryParse(spacesValue, out int spacesCount);

            if (spacesCount == 0)
            {
                spacesCount++;
            }

            for (int i = 0; i < spacesCount; i++)
            {
                htmlElement.SetValue(htmlElement.Value + "&nbsp;");
            }
        }

        private static void CopyAttributes(XElement odElement, XElement htmlElement)
        {
            if (odElement.HasAttributes)
            {
                foreach (var attr in odElement.Attributes())
                {
                    if (HtmlAttrTrans.TryGetValue(attr.Name.LocalName, out string htmlAttrName))
                    {
                        var htmlAttr = new XAttribute(htmlAttrName, attr.Value);
                        htmlElement.Add(htmlAttr);
                    }
                }
            }
        }
    }
}
