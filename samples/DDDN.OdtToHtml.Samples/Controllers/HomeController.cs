using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace DDDN.OdtToHtml.Samples
{
	public class HomeController : Controller
	{
		private readonly IHostingEnvironment _hostingEnvironment;

		public HomeController(IHostingEnvironment hostingEnvironment)
		{
			_hostingEnvironment = hostingEnvironment ?? throw new System.ArgumentNullException(nameof(hostingEnvironment));
		}

		public IActionResult Index()
		{
			return View(_hostingEnvironment.WebRootFileProvider.GetDirectoryContents("odt"));
		}

		public IActionResult Open(string id)
		{
			const string contentSubDirname = "content";
			OdtConvertData convertData = null;
			var odtFileInfo = _hostingEnvironment.WebRootFileProvider.GetFileInfo(Path.Combine("odt", id));

			using (IOdtFile odtFile = new OdtFile(odtFileInfo.PhysicalPath))
			{
				var odtCon = new OdtConvert(odtFile);
				convertData = odtCon.Convert(new OdtConvertSettings
				{
					RootElementTagName = "article",
					RootElementId = "artid",
					RootElementClassNames = "artclass",
					LinkUrlPrefix = $"/{contentSubDirname}"
				});
			}

			ViewData["ArticleHtml"] = convertData.Html;
			ViewData["ArticleCss"] = convertData.Css;

			foreach (var articleContent in convertData.EmbedContent)
			{
				var contentLink = Path.Combine(_hostingEnvironment.WebRootPath, contentSubDirname, articleContent.Name);
				System.IO.File.WriteAllBytes(contentLink, articleContent.Data);
			}

			return View();
		}
	}
}
