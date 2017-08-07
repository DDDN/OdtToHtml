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
using System.IO.Compression;

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
            var odtFileInfo = _hostingEnvironment.WebRootFileProvider.GetFileInfo("odt\\Sample1.odt");
            using (var odtZipArchive = ZipFile.OpenRead(odtFileInfo.PhysicalPath))
            {
                using (var odtCon = new ODTConvert(odtZipArchive))
                {
                    var html = odtCon.GetHtml();
                    ViewData["ArticleHtml"] = html;
                }
            }

            return View();
        }
    }
}
