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

using System.Text;
using System.Xml;
using System.Xml.Xsl;

namespace DDDN.Office.ODT
{
    public class ODTConvert : IODTConvert
    {
        public string GetHtml(XmlDocument contentXml, XmlReader htmlXslt)
        {
            var htmlBuilder = new StringBuilder(4096);

            using (XmlWriter htmlWriter = XmlWriter.Create(htmlBuilder))
            {
                XslCompiledTransform xslt = new XslCompiledTransform();
                xslt.Load(htmlXslt);

                xslt.Transform(contentXml, htmlWriter);
            }

            return htmlBuilder.ToString();
        }

        public string GetCss(XmlDocument stylesXml, XmlReader cssXslt)
        {
            var styleBuilder = new StringBuilder(4096);

            using (XmlWriter htmlWriter = XmlWriter.Create(styleBuilder))
            {
                XslCompiledTransform xslt = new XslCompiledTransform();
                xslt.Load(cssXslt);

                xslt.Transform(stylesXml, htmlWriter);
            }

            return styleBuilder.ToString();
        }
    }
}
