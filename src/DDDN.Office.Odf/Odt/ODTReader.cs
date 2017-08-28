/*
* DDDN.Office.Odf.Odt.ODTReader
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
using System.Xml;
using System.Xml.Linq;

namespace DDDN.Office.Odf.Odt
{
    public static class ODTReader
    {
        public static string GetValue(XElement xElement)
        {
            return WalkTheNodes(xElement.Nodes());
        }

        public static string GetFirstHeaderText(IODTFile odtFile)
        {
            return "";
        }

        public static Dictionary<string, string> GetTranslations(string cultureNameFromFileName, IODTFile odtFile)
        {
            var translations = new Dictionary<string, string>();

            XDocument contentXDoc = odtFile.GetZipArchiveEntryAsXDocument("content.xml");

            var contentEle = contentXDoc.Root
                .Elements(XName.Get("body", ODFXmlNamespaces.Office))
                .Elements(XName.Get("text", ODFXmlNamespaces.Office))
                .First();

            foreach (var table in contentEle.Elements()
                .Where(p => p.Name.LocalName.Equals("table", StringComparison.CurrentCultureIgnoreCase)))
            {
                foreach (var row in table.Elements()
                .Where(p => p.Name.LocalName.Equals("table-row", StringComparison.CurrentCultureIgnoreCase))
                .Skip(1))
                {
                    var cells = row.Elements()
                        .Where(p => p.Name.LocalName.Equals("table-cell", StringComparison.CurrentCultureIgnoreCase));

                    if (cells.Any())
                    {
                        var translationKey = ODTReader.GetValue(cells.First());
                        var translation = ODTReader.GetValue(cells.Skip(1).First());
                        translations.Add($"{translationKey}.{cultureNameFromFileName}", translation);
                    }
                }
            }

            return translations;
        }

        private static string WalkTheNodes(IEnumerable<XNode> nodes)
        {
            if (nodes == null)
            {
                throw new ArgumentNullException(nameof(nodes));
            }

            string val = string.Empty;

            foreach (var node in nodes)
            {
                if (node.NodeType == XmlNodeType.Text)
                {
                    var textNode = node as XText;
                    val += textNode.Value;
                }
                else if (node.NodeType == XmlNodeType.Element)
                {
                    var elementNode = node as XElement;

                    if (elementNode.Name.Equals(XName.Get("s", ODFXmlNamespaces.Text)))
                    {
                        val += " ";
                    }
                    else
                    {
                        val += WalkTheNodes(elementNode.Nodes());
                    }
                }
            }

            return val;
        }
    }
}
