

# DDDN.OdtToHtml
OdtToHtml is a .NET Standard/Core library for converting ODT documents (Open Document Text Format) into responsive HTML/CSS.

## How to use
### Nuget package
Simply add the [DDDN.OdtToHtml](https://www.nuget.org/packages/DDDN.OdtToHtml/) nuget package reference to your project.
### Build requirements
You can download the code and compile the library using Visual studio 2017 and the .NET Core 2.x SDK.
### Sample application
Please have a look at the [DDDN.OdtToHtml.Samples](https://github.com/DDDN/OdtToHtml/tree/dev/samples/DDDN.OdtToHtml.Samples) ASP.NET Core 2.0 sample application.
### Sample code (MVC action method)
The following example shows how to call the `Convert` method within a MVC controller action method, in order to get the HTML/CSS from an ODT document saved somewhere in the www root folder.
```C#
// id is the ODT document file name without the suffix
public IActionResult Open(string id)
{
	// subdirectory name where the ODT document content files like images will be stored
	const string contentSubDirname = "content";

	// structure returned by the Convert method with the converted data
	OdtConvertedData convertedData = null;

	// generating path to the ODT document that will be converted
	var odtFileInfo = _hostingEnvironment.WebRootFileProvider.GetFileInfo(Path.Combine("odt", id));

	// settings for the Convert method
	var odtConvertSettings = new OdtConvertSettings
	{
		RootElementTagName = "article", // root HTML tag that will contain the converted HTML
		RootElementId = "article_id", // the id attribute value of the root HTML tag (optional)
		RootElementClassNames = "article_class", // the class attribute value of the root HTML tag (optional)
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
```
### The HTML part (MVC Razor view)
Using the ViewData object, you can pass the converted HTML and CSS to the MVC view:
```C#
// controller action
ViewData["ArticleHtml"] = convertedData.Html;
ViewData["ArticleCss"] = convertedData.Css;
```
this is how the MCV view can looks like:
```C#
@section Styles
{
	@Html.Raw(WebUtility.HtmlDecode((string)ViewData["ArticleCss"]))
}

@Html.Raw(WebUtility.HtmlDecode((string)ViewData["ArticleHtml"]))
```
and this is an example of the the _Layout.cshtml part:
```HTML
<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<title>@ViewData["Title"] - DDDN.OdtToHtml.Samples</title>
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<link rel="stylesheet" href="~/css/site.css" asp-append-version="true" />
	<style>
		@RenderSection("styles", required: false)
	</style>
</head>
<body>
	<div class="container">
		<header>
			<a asp-controller="Home" asp-action="Index">DDDN.OdtToHtml.Samples - Home</a>
		</header>
		<main>
			@RenderBody()
		</main>
	</div>
	<script src="~/js/site.js" asp-append-version="true"></script>
	@RenderSection("scripts", required: false)
</body>
</html>
```

## Development
OdtToHtml has been developed as a part of the [CrossBlog ASP.NET Core blog engine](https:\\github.com/DDDN/CrossBlog). With CrossBlog you can write your blog using the office/word processing application of your choice and of course using the ODT document format. But it can of course also be used otherwise.

## Dependencies
There are no  package or code dependencies other than the .NETStandard 2.0 libraries.

## Missing features
- all canvas/shapes features
- outline functionality (only lists at the moment)
- TOC formatting
- correct text flow around images and other content

Contact me if something important is missing or doesn't work properly.

## License
The source code is mostly licensed under GNU General Public License, version 2 of the license. Please refer to the source code file headers for detailed licensing information.
## Version history
**0.25.1.211 (alpha)**
- improved rendering of list styles and list fonts
- improved text decorations
- added anchor links for the document TOC
- the conversion now also returns a CSS property value list with all used font families

**0.2.18.1 (alpha)**
- first alpha release
