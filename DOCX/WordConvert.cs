/*
* DDDN.Net.OpenXML.WordConvert
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

using DDDN.Net.Html;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DDDN.Office.DOCX
{
    public class WordConvert : IWordConvert, IDisposable
    {
        private WordprocessingDocument Doc { get; set; }
        private Dictionary<string, WStyleInfo> WStyleInfos { get; set; } = new Dictionary<string, WStyleInfo>();
        private List<WParagraphInfo> WParagraphInfos { get; set; } = new List<WParagraphInfo>();
        private static readonly string StyleNamePrefix = "Word";
        /// <summary>
        /// CLass Constructor
        /// </summary>
        /// <param name="docStream">DOCX file stream.</param>
        public WordConvert(FileStream docStream)
        {
            if (docStream == null)
            {
                throw new ArgumentNullException(nameof(docStream));
            }

            Doc = WordprocessingDocument.Open(docStream, false);
            GetWDocInfo();
        }
        /// <summary>
        /// Renders the whole document body in HTML.
        /// </summary>
        /// <returns></returns>
        public string GetHTML(string rootHtmlTagName)
        {
            IHtmlNode rootHtmlTag = new HtmlNode(rootHtmlTagName);

            foreach (var pInfo in WParagraphInfos)
            {
                IHtmlNode pNode = new HtmlNode(HtmlTag.P);
                pNode.AddClass("WordNormal");
                pNode.AddClass($"{StyleNamePrefix}{pInfo.Id}");
                rootHtmlTag.AddChild(pNode);

                foreach (var rInfo in pInfo.Runs)
                {
                    var runNode = new HtmlNode(HtmlTag.Span, rInfo.Text);

                    if (rInfo.FontColor != null)
                    {
                        runNode.AddStyleProperty(CssProperty.color, $"#{rInfo.FontColor}");
                    }
                    if (rInfo.FontSize != null)
                    {
                        runNode.AddStyleProperty(CssProperty.font_size, $"{rInfo.FontSize}px");
                    }

                    pNode.AddChild(runNode);
                }
            }

            var htmlB = new StringBuilder(2048);
            rootHtmlTag.RenderHtml(htmlB);
            var html = htmlB.ToString();
            return html;
        }
        /// <summary>
        /// Renders the CSS that should be linked/added into the html header.
        /// </summary>
        /// <returns>The css to be linked/added to the header.</returns>
        public string GetCSS()
        {
            var styles = Doc.MainDocumentPart.StyleDefinitionsPart?.Styles;

            if (styles == null)
            {
                return string.Empty;
            }

            var cssDef = new Dictionary<string, Dictionary<string, string>>();

            foreach (var style in styles.OfType<Style>())
            {
                var styleInfo = new WStyleInfo()
                {
                    StyleId = style.StyleId,
                    StyleType = style.Type,
                    BasedOnStyle = style.BasedOn
                };

                if (style.StyleRunProperties != null)
                {
                    styleInfo.RunProperty.FontColor = style.StyleRunProperties.OfType<Color>()?.First().Val;
                    styleInfo.RunProperty.FontSize = style.StyleRunProperties.OfType<FontSize>()?.First().Val;
                }

                WStyleInfos.Add(styleInfo.StyleId, styleInfo);
            }

            var cssB = new StringBuilder(2048);
            return cssB.ToString();
        }
        /// <summary>
        /// Parse the needed infromation from the DOCX strong typed classes, that represent the XML structure.
        /// </summary>
        private void GetWDocInfo()
        {
            foreach (var para in Doc.MainDocumentPart.Document.Body.Elements<Paragraph>())
            {
                var pInfo = new WParagraphInfo()
                {
                    Id = para.ParagraphProperties?.ParagraphStyleId?.Val
                };

                foreach (var run in para.OfType<Run>())
                {
                    var runInfo = new WRunInfo()
                    {
                        Text = run.InnerText,
                        FontSize = run.RunProperties.FontSize?.Val,
                        FontColor = run.RunProperties.Color?.Val
                    };

                    pInfo.Runs.Add(runInfo);
                }

                WParagraphInfos.Add(pInfo);
            }
        }

        #region IDisposable Support
        private bool disposed = false;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                Doc.Dispose();
            }

            disposed = true;
        }
        #endregion
    }
}
