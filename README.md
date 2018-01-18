# OdtToHtml
OdtToHtml is a .NET Core/Standard library for converting ODT documents (Open Document Text Format) into responsive HTML/CSS.

## How to use
### Nuget package
Simply add the [DDDN.OdtToHtml](https://www.nuget.org/packages/DDDN.OdtToHtml/) nuget package reference to your project.
### Sample application
Please have a look at the ASP.NET Core 2.0 [DDDN.OdtToHtml.Samples](https://github.com/DDDN/OdtToHtml/tree/dev/samples/DDDN.OdtToHtml.Samples) sample application.
### Sample code
This example of an ASP.NET sample method shows how to call the `Convert` method in order to get the HTML/CSS generated from an ODT document.
```C#
public IActionResult Convert(string id)
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
```
## Development
OdtToHtml has been developed as a part of the [CrossBlog ASP.NET Core blog engine](https:\\github.com/DDDN/CrossBlog). With CrossBlog you can write your blog using the office/word processing application of your choice and of course using the ODT document format.
But it can of course also be used otherwise.

## Dependencies
There are no  package or code dependencies other than the .NETStandard 2.0 libraries.

## Missing features
- all canvas/shapes features
- outline functionality (only lists at the moment)
- other features that you don't necessary need for writing blog posts

## License
The source code is mainly licensed under GNU General Public License, version 2. Please refer to the source code files for detailed licensing information for each file.
