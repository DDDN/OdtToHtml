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
        private IODTFile ODTFile;

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

        private static readonly Dictionary<string, OdfStyleToCss> CssTrans = new Dictionary<string, OdfStyleToCss>()
        {
            ["border-model"] = new OdfStyleToCss
            {
                OdfName = "",
                CssName = "",
                Values = new Dictionary<string, string>()
                {
                    ["d"] = ""
                }
            }
        };


        public ODTConvert(IODTFile odtFile)
        {
            ODTFile = odtFile ?? throw new System.ArgumentNullException(nameof(odtFile));
        }

        public ODTConvertData Convert()
        {
            var data = new ODTConvertData();
            GetHtml(data);
            GetCss(data);
            return data;
        }

        private void GetCss(ODTConvertData data)
        {
            XDocument stylesXDoc = ODTFile.GetZipArchiveEntryAsXDocument("styles.xml");
            XDocument contentXDoc = ODTFile.GetZipArchiveEntryAsXDocument("content.xml");
            List<IOdfStyle> Styles = new List<IOdfStyle>();

            var fontFaceDeclarations = contentXDoc.Root
                 .Elements(XName.Get("font-face-decls", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
                 .Elements()
                 .Where(p => p.Name.Equals(XName.Get("font-face", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")));
            StylesWalker(fontFaceDeclarations, Styles);

            var automaticStyles = contentXDoc.Root
                 .Elements(XName.Get("automatic-styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
                 .Elements()
                 .Where(p => p.Name.Equals(XName.Get("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")));
            StylesWalker(automaticStyles, Styles);

            var defaultStyles = stylesXDoc.Root
                 .Elements(XName.Get("styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
                 .Elements()
                 .Where(p => p.Name.Equals(XName.Get("default-style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")));
            StylesWalker(defaultStyles, Styles);

            var styles = stylesXDoc.Root
                 .Elements(XName.Get("styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
                 .Elements()
                 .Where(p => p.Name.Equals(XName.Get("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")));
            StylesWalker(styles, Styles);

            data.Css = RenderCss(Styles);
        }

        private string RenderCss(List<IOdfStyle> styles)
        {
            var builder = new StringBuilder(1024);

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
                    builder.Append($"{attr.Name}: {attr.Value};{Environment.NewLine}");
                }

                foreach (var props in style.PropAttrs.Values)
                {
                    foreach (var propAttr in props)
                    {
                        builder.Append($"{propAttr.Name}: {propAttr.Value};{Environment.NewLine}");
                    }
                }

                builder.Append("}");
            }

            return builder.ToString();
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

        private void GetHtml(ODTConvertData data)
        {
            XDocument contentXDoc = ODTFile.GetZipArchiveEntryAsXDocument("content.xml");

            var contentEle = contentXDoc.Root
                    .Elements(XName.Get("body", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
                    .Elements(XName.Get("text", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
                    .First();

            var htmlEle = new XElement(HtmlTagsTrans[contentEle.Name.LocalName]);
            ContentNodesWalker(contentEle.Nodes(), htmlEle);

            data.Html = htmlEle.ToString(SaveOptions.DisableFormatting);
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

                    if (elementNode.Name.Equals(XName.Get("s", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")))
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
            var spacesValue = odElement.Attribute(XName.Get("c", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"))?.Value;
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

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    ODTFile.Dispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~ODTConvert() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}
