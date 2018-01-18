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
			// subdirectory name where the ODT document content files like images will be stored
			const string contentSubDirname = "content";

			OdtConvertedData convertData = null;

			// generating path to the ODT document that will be converted
			var odtFileInfo = _hostingEnvironment.WebRootFileProvider.GetFileInfo(Path.Combine("odt", id));

			// open the ODT file from the file system
			using (IOdtFile odtFile = new OdtFile(odtFileInfo.PhysicalPath))
			{
				var odtCon = new OdtConvert();

				// CALL TO THE CONVERT METHOD
				convertData = odtCon.Convert(odtFile, new OdtConvertSettings
				{
					// root HTML tag info that will contains the converted HTML
					RootElementTagName = "article",
					RootElementId = "artid",
					RootElementClassNames = "artclass",
					// here you can provide a prefix for all content links to match your environment requirements
					LinkUrlPrefix = $"/{contentSubDirname}"
				});
			}

			// move the generated HTML/CSS to the razor web page
			ViewData["ArticleHtml"] = convertData.Html;
			ViewData["ArticleCss"] = convertData.Css;

			// write the content (like images) of the ODT document to the file system to make it available to the web browser
			foreach (var articleContent in convertData.EmbedContent)
			{
				var contentLink = Path.Combine(_hostingEnvironment.WebRootPath, contentSubDirname, articleContent.LinkName);
				System.IO.File.WriteAllBytes(contentLink, articleContent.Data);
			}

			return View();
		}
	}
}
