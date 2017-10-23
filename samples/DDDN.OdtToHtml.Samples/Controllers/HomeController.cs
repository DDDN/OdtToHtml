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
			var odtFileInfo = _hostingEnvironment.WebRootFileProvider.GetFileInfo(Path.Combine("odt", id));

			using (IOdtFile odtFile = new OdtFile(odtFileInfo.PhysicalPath))
			{
				var odtCon = new OdtConvert(odtFile);
				var convertData = odtCon.Convert(new OdtConvertSettings
				{
					RootElementTagName = "article",
					RootElementId = "odtarticle",
					LinkUrlPrefix = "/content"
				});

				ViewData["ArticleHtml"] = convertData.Html;
				ViewData["ArticleCss"] = convertData.Css;

				foreach (var content in convertData.EmbedContent)
				{
					System.IO.File.WriteAllBytes(Path.Combine(_hostingEnvironment.WebRootPath, "content", content.Name), content.Content);
				}
			}

			return View();
		}
	}
}
