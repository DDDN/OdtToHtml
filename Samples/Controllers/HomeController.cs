/*
* DDDN.Office.ODT.Samples.HomeController
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

using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace DDDN.Office.ODT.Samples
{
    public class HomeController : Controller
    {
        IHostingEnvironment _hostingEnvironment;

        public HomeController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment ?? throw new System.ArgumentNullException(nameof(hostingEnvironment));
        }

        public IActionResult Index()
        {
            var htmlXsltFilePath =
                Path.Combine(new string[] {
                _hostingEnvironment.ContentRootPath,
                @"\bin\Debug\netcoreapp2.0\DDDN.Office.ODT.Xslt",
                "ODT2HTML_1.0.xslt" });
            var htmlXsltFileInfo = _hostingEnvironment.ContentRootFileProvider.GetFileInfo(htmlXsltFilePath);

            var cssXsltFilePath =
                Path.Combine(new string[] {
                _hostingEnvironment.ContentRootPath,
                @"\bin\Debug\netcoreapp2.0\DDDN.Office.ODT.Xslt",
                "ODT2CSS_1.0.xslt" });
            var cssXsltFileInfo = _hostingEnvironment.ContentRootFileProvider.GetFileInfo(cssXsltFilePath);

            var odtFilePath = Path.Combine(new string[] {
                _hostingEnvironment.ContentRootPath,
                @"\bin\Debug\netcoreapp2.0\DDDN.Office.ODT.Samples.ODT",
                "sample1.odt" });
            var odtFileInfo = _hostingEnvironment.ContentRootFileProvider.GetFileInfo(odtFilePath);
            var odtFileInfoZipStream = ZipFile.OpenRead(odtFileInfo.PhysicalPath);

            var contentEntry = odtFileInfoZipStream.Entries.Where(
                p => p.Name.Equals("content.xml", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();

            var stylesEntry = odtFileInfoZipStream.Entries.Where(
                p => p.Name.Equals("styles.xml", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();


            var odtConverter = new ODTConvert();

            var htmlXsltDoc = new XmlDocument();
            htmlXsltDoc.Load(htmlXsltFileInfo.CreateReadStream());
            var contentXmlDoc = new XmlDocument();
            contentXmlDoc.Load(contentEntry.Open());
            var html = odtConverter.GetHtml(contentXmlDoc, htmlXsltDoc);
            ViewData["ArticleHtml"] = html;

            var cssXsltDoc = new XmlDocument();
            cssXsltDoc.Load(cssXsltFileInfo.CreateReadStream());
            var stylesXmlDoc = new XmlDocument();
            stylesXmlDoc.Load(stylesEntry.Open());
            var css = odtConverter.GetCss(stylesXmlDoc, cssXsltDoc);
            ViewData["ArticleCss"] = css;

            return View();
        }
    }
}
