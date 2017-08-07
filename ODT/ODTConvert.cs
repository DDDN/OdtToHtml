/*
* DDDN.Office.ODT.ODTConvert
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
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace DDDN.Office.ODT
{
    public class ODTConvert : IODTConvert, IDisposable
    {
        private ZipArchive ODTZipArchive;

        private Dictionary<string, (string, bool)> Tags = new Dictionary<string, (string TagName, bool TakeValue)>()
        {
            ["text"] = ("article", false),
            ["p"] = ("p", true),
            ["span"] = ("span", false),
            ["table"] = ("table", false),
            ["table-column"] = ("th", false),
            ["table-row"] = ("tr", false),
            ["table-cell"] = ("td", false),
            ["list"] = ("ul", false),
            ["list-item"] = ("li", false)
        };

        public ODTConvert(ZipArchive odtDocumentUnzipped)
        {
            ODTZipArchive = odtDocumentUnzipped ?? throw new System.ArgumentNullException(nameof(odtDocumentUnzipped));
        }

        public string GetCss()
        {
            throw new System.NotImplementedException();
        }

        public string GetHtml()
        {
            var contentEntry = ODTZipArchive.Entries
                .Where(p => p.Name.Equals("content.xml", StringComparison.InvariantCultureIgnoreCase))
                .FirstOrDefault();
            XDocument contentXDoc = null;

            using (var contentStream = contentEntry.Open())
            {
                contentXDoc = XDocument.Load(contentStream);
            }

            var contentEle = contentXDoc.Root
                .Elements(XName.Get("body", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
                .Elements(XName.Get("text", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
                .First();

            if (TryCreateHtmlElement(contentEle, out XElement htmlEle))
            {
                ElementWalker(contentEle.Elements(), htmlEle);
            }

            return htmlEle.ToString();
        }

        private void ElementWalker(IEnumerable<XElement> contentElements, XElement htmlElement)
        {
            foreach (var ele in contentElements)
            {
                if (TryCreateHtmlElement(ele, out XElement htmlEle))
                {
                    htmlElement.Add(htmlEle);
                    ElementWalker(ele.Elements(), htmlEle);
                }
            }
        }

        private bool TryCreateHtmlElement(XElement odElement, out XElement htmlElement)
        {
            var htmlTag = (TagName: String.Empty, TakeValue: false);

            if (Tags.TryGetValue(odElement.Name.LocalName, out htmlTag))
            {
                if (htmlTag.TakeValue)
                {
                    htmlElement = new XElement(htmlTag.TagName, odElement.Value);
                }
                else
                {
                    htmlElement = new XElement(htmlTag.TagName);
                }

                return true;
            }
            else
            {
                htmlElement = null;
                return false;
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
                    ODTZipArchive.Dispose();
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
