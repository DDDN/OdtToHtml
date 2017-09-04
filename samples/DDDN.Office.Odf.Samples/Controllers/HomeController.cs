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
			var odtFileInfo = _hostingEnvironment.WebRootFileProvider.GetFileInfo("odt\\Sample1_images.odt");

			using (IODTFile odtFile = new ODTFile(odtFileInfo.PhysicalPath))
			{
				var odtCon = new ODTConvert(odtFile);
				var convertData = odtCon.Convert();
				ViewData["ArticleHtml"] = convertData.Html;
				ViewData["ArticleCss"] = convertData.Css;
			}

			return View();
		}
	}
}
