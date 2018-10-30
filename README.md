

# DDDN.OdtToHtml
OdtToHtml is a .NET Standard/Core library for converting ODT documents (Open Document Text Format) into responsive HTML/CSS.

It should be understand as a more comfortable way of writing HTML articles. You can use all the editing/correcting features of Microsoft Office Word, LibreOffice or other editors for a faster writing and reviewing your text and content.

OdtToHtml has been developed as a part of the [CrossBlog ASP.NET Core blog engine](https:\\github.com/DDDN/CrossBlog), but it can also be used otherwise/standalone.

## How to get OdtToHtml
### Nuget package
Simply add the [DDDN.OdtToHtml](https://www.nuget.org/packages/DDDN.OdtToHtml/) nuget package reference to your project using Visual Studio or by downloading it manually.
### Get it from here
You can pull the code form Github and compile the solution using Visual studio 2017 and the .NET Core 2.x SDK. The solution contains a sample web application with some example ODT documents that can be converted to HTML with a single click.

## How to use the library
Please have a look at the [DDDN.OdtToHtml.Samples](https://github.com/DDDN/OdtToHtml/tree/dev/samples/DDDN.OdtToHtml.Samples) ASP.NET Core 2.0 sample web application that contains the following code:
### Convert call inside a MVC controller action method
The following example shows how to call the `Convert` method within a MVC controller action method in order to get the HTML/CSS from an ODT document saved somewhere under the wwwroot folder.
```C#
// id is the ODT document file name without the suffix
public IActionResult Open(string odtDocFileName)
{
	// holds the wwwroot subdirectory name where the ODT document content like images will be saved
	const string contentSubDirname = "content";

	// structure will contain all the data returned form the ODT to HTML conversion
	OdtConvertedData convertedData = null;

	// builds the path to the ODT document. "odt" is a wwwroot subfolder
	var odtFileInfo = _hostingEnvironment.WebRootFileProvider.GetFileInfo(Path.Combine("odt", odtDocFileName));

	// providing settings for the ODT to HTML conversion call
	var odtConvertSettings = new OdtConvertSettings
	{
		RootElementTagName = "article", // root HTML tag that will contain the converted HTML
		RootElementId = "article_id", // the id attribute value of the root HTML tag (optional)
		RootElementClassNames = "article_class", // the class attribute value of the root HTML tag (optional)
		LinkUrlPrefix = $"/{contentSubDirname}", // here you can provide a partial path directly under the wwwroot folder content (images) link generation
		DefaultTabSize = "2rem" // the default value for tabs (not tab stops)
	};

	// open the ODT document on the file system and call the Convert method to convert the document to HTML
	using (IOdtFile odtFile = new OdtFile(odtFileInfo.PhysicalPath))
	{
		convertedData = new OdtConvert().Convert(odtFile, odtConvertSettings);
	}
	
	// write the content (like images) of the ODT document to the file system to make it available to the web browser requests
	foreach (var articleContent in convertedData.EmbedContent)
	{
		var contentLink = Path.Combine(_hostingEnvironment.WebRootPath, contentSubDirname, articleContent.LinkName);
		System.IO.File.WriteAllBytes(contentLink, articleContent.Data);
	}

	ViewData["ArticleHtml"] = convertedData.Html; // move the generated HTML to the razor view
	ViewData["ArticleCss"] = convertedData.Css; // move the generated CSS to the razor view
	
	// additional usefull data returned for conversion
	var usedFontFamilies = convertedData.UsedFontFamilies; // all font families used, useful for font links
	var pageInfo = convertedData.PageInfo; // contains the document sheet's dimensions and margins
	var documentFirstHeader = convertedData.DocumentFirstHeader; // the "text only" content of the first document header for preview/teaser purposes
	var documentFirstParagraph = convertedData.DocumentFirstParagraph; // the "text only" content of the first document paragraph for preview/teaser purposes

	return View();
}

// That’s all! You have successfully created a web page from an office document! :)

```
### The Razor View part
This is how the MCV Razor view that uses a _Layout.cshtml can looks like:
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
			<a asp-controller="Home" asp-action="Index">Your converted ODT document</a>
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

## OdtToHtml library Dependencies
There are no package or code dependencies other than the .NETStandard 2.0 libraries.

## Missing features
Some of the features of the ODT format (some listed below) are not implemented by the converter because they are not that usefull for displaying HTML pages.
Please feel free to contact me if something important is missing or doesn't work properly.

- all canvas/shapes features
- outline functionality (only lists at the moment)
- TOC formatting partially works 
- correct text flow around images and other content

## License
The source code is mostly licensed under GNU General Public License, version 2 of the license. Please refer to the source code file headers for detailed licensing information.

## Version history
**0.27.10.301 (alpha)**
 - improved list and overall rendering
 
**0.26.2.111 (alpha)**
 - bug fixes
 
**0.26.2.102 (alpha)**
 - improved rendering of lists/tables
 - minor rendering improvements
 
**0.25.1.211 (alpha)**
 - improved rendering of list styles and list fonts
 - improved text decorations
 - added anchor links for the document TOC
 - the conversion now also returns a CSS property value list with all used font families
 
**0.2.18.1 (alpha)**
 - first alpha release

