using DDDN.Office.Odf.Odt;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace DDDN.Office.Odf.Samples
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

			using (IODTFile odtFile = new ODTFile(odtFileInfo.PhysicalPath))
			{
				using (IODTConvert odtCon = new ODTConvert(odtFile))
				{
					var html = odtCon.GetHtml();
					ViewData["ArticleHtml"] = html;
					var css = odtCon.GetCss();
					ViewData["ArticleCss"] = css;
				}
			}

			return View();
		}
	}
}
