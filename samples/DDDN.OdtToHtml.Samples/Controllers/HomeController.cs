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

			// the structure which converted data which the Convert method returns
			OdtConvertedData convertedData = null;

			// generating path to the ODT document that will be converted
			var odtFileInfo = _hostingEnvironment.WebRootFileProvider.GetFileInfo(Path.Combine("odt", id));

			// settings for the Convert method
			var odtConvertSettings = new OdtConvertSettings
			{
				RootElementTagName = "article", // root HTML tag info that will contains the converted HTML
				RootElementId = "article_id", // the id attribute value of the root HTML tag
				RootElementClassNames = "article_class", // the class attribute value of the root HTML tag
				LinkUrlPrefix = $"/{contentSubDirname}", // here you can provide a prefix for all content links to match your environment requirements
				DefaultTabSize = "2rem" // the default value for tabs (not tab stops)
			};

			// open the ODT document from the file system and call the Convert method to get the HTML/CSS
			using (IOdtFile odtFile = new OdtFile(odtFileInfo.PhysicalPath))
			{
				convertedData = new OdtConvert().Convert(odtFile, odtConvertSettings);
			}

			ViewData["ArticleHtml"] = convertedData.Html; // move the generated HTML to the razor view
			ViewData["ArticleCss"] = convertedData.Css; // move the generated CSS to the razor view
			var usedFontFamilies = convertedData.UsedFontFamilies; // all font families used in CSS/HTML useful for font links
			var pageInfo = convertedData.PageInfo; // contains the document sheet's dimensions and margins
			var documentFirstHeader = convertedData.DocumentFirstHeader; // the "text only" content of the first document header for preview purposes
			var documentFirstParagraph = convertedData.DocumentFirstParagraph; // the "text only" content of the first document paragraph for preview purposes

			// write the content (like images) of the ODT document to the file system to make it available to the web browser
			foreach (var articleContent in convertedData.EmbedContent)
			{
				var contentLink = Path.Combine(_hostingEnvironment.WebRootPath, contentSubDirname, articleContent.LinkName);
				System.IO.File.WriteAllBytes(contentLink, articleContent.Data);
			}

			return View();
		}
	}
}
