/*
DDDN.Office.Odf.Odt.ODTConvert
Copyright(C) 2017 Lukasz Jaskiewicz(lukasz @jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DDDN.Office.Odf.Odt
{
	public class ODTConvert : IODTConvert
	{
		private IODTFile OdtFile;
		private string LinkUrlPrefix;
		private XDocument ContentXDoc;
		private XDocument StylesXDoc;
		private List<IOdfStyle> Styles;
		private List<ODFEmbedContent> EmbedContent = new List<ODFEmbedContent>();


		private static readonly Dictionary<string, string> HtmlTagsTrans = new Dictionary<string, string>()
		{
			["text"] = "article",
			["h"] = "p",
			["p"] = "p",
			["span"] = "span",
			["paragraph"] = "p",
			["graphic"] = "img",
			["image"] = "img",
			["s"] = "span",
			["a"] = "a",
			["frame"] = "div",
			["text-box"] = "div",
			["table"] = "table",
			["table-columns"] = "tr",
			["table-column"] = "th",
			["table-row"] = "tr",
			["table-cell"] = "td",
			["list"] = "ul",
			["list-item"] = "li",
			["automatic-styles"] = "style"
		};

		private static readonly Dictionary<string, string> HtmlAttrTrans = new Dictionary<string, string>()
		{
			["p.style-name"] = "class",
			["span.style-name"] = "class",
			["h.style-name"] = "class",
			["s.style-name"] = "class",
			["table-column.style-name"] = "class",
			["table-row.style-name"] = "class",
			["table-cell.style-name"] = "class",
			["list.style-name"] = "class",
			["frame.style-name"] = "class",
			["table.style-name"] = "class",
			["a.href"] = "href",
			["image.href"] = "src",
			["a.target-frame-name"] = "target"
		};

		private static readonly Dictionary<string, string> StyleAttr = new Dictionary<string, string>()
		{
			["frame.width"] = "width",
			["frame.height"] = "height",
		};

		private static readonly List<OdfStyleToCss> CssTrans = new List<OdfStyleToCss>()
		{
			{
				 new OdfStyleToCss
				 {
					  OdfName = "border-model",
					  CssName = "border-spacing",
					  Values = new Dictionary<string, string>()
					  {
							["collapsing"] = "0"
					  }
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "writing-mode",
					  CssName = "writing-mode",
					  Values = new Dictionary<string, string>()
					  {
							["lr"] = "horizontal-tb",
							["lr-tb"] = "horizontal-tb",
							["rl"] = "horizontal-tb",
							["tb"] = "vertical-lr",
							["tb-rl"] = "vertical-rl"
					  }
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "hyphenate",
					  CssName = "hyphens",
					  Values = new Dictionary<string, string>()
					  {
							["false"] = "none"
					  }
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "font-name",
					  CssName = "font-family"
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "page-width",
					  CssName = "width"
				 }
			},
			{
				 new OdfStyleToCss
				 {
					  OdfName = "rel-width",
					  CssName = "width"
				 }
			},
			{ new OdfStyleToCss { OdfName = "page-height" } },
			{ new OdfStyleToCss { OdfName = "num-format" } },
			{ new OdfStyleToCss { OdfName = "print-orientation" } },
			{ new OdfStyleToCss { OdfName = "keep-with-next" } },
			{ new OdfStyleToCss { OdfName = "keep-together" } },
			{ new OdfStyleToCss { OdfName = "widows" } },
			{ new OdfStyleToCss { OdfName = "language" } },
			{ new OdfStyleToCss { OdfName = "country" } },
			{ new OdfStyleToCss { OdfName = "display-name" } },
			{ new OdfStyleToCss { OdfName = "orphans" } },
			{ new OdfStyleToCss { OdfName = "fill" } },
			{ new OdfStyleToCss { OdfName = "fill-color" } },
			{ new OdfStyleToCss { OdfName = "line-break" } },
			{ new OdfStyleToCss { OdfName = "punctuation-wrap" } },
			{ new OdfStyleToCss { OdfName = "number-lines" } },
			{ new OdfStyleToCss { OdfName = "text-autospace" } },
			{ new OdfStyleToCss { OdfName = "snap-to-layout-grid" } },
			{ new OdfStyleToCss { OdfName = "text-autospace" } },
			{ new OdfStyleToCss { OdfName = "tab-stop-distance" } },
			{ new OdfStyleToCss { OdfName = "use-window-font-color" } },
			{ new OdfStyleToCss { OdfName = "letter-kerning" } },
			{ new OdfStyleToCss { OdfName = "text-scale" } },
			{ new OdfStyleToCss { OdfName = "text-position" } },
			{ new OdfStyleToCss { OdfName = "font-relief" } },
			{ new OdfStyleToCss { OdfName = "horizontal-pos" } },
			{ new OdfStyleToCss { OdfName = "horizontal-rel" } },
			{ new OdfStyleToCss { OdfName = "vertical-pos" } },
			{ new OdfStyleToCss { OdfName = "vertical-rel" } },
			{ new OdfStyleToCss { OdfName = "break-before" } },
			{ new OdfStyleToCss { OdfName = "font-size-complex" } },
			{ new OdfStyleToCss { OdfName = "font-size-asian" } },
			{ new OdfStyleToCss { OdfName = "font-weight-complex" } },
			{ new OdfStyleToCss { OdfName = "font-weight-asian" } },
			{ new OdfStyleToCss { OdfName = "font-name-complex" } },
			{ new OdfStyleToCss { OdfName = "font-name-asian" } },
			{ new OdfStyleToCss { OdfName = "font-style-asian" } },
			{ new OdfStyleToCss { OdfName = "font-style-complex" } },
			{ new OdfStyleToCss { OdfName = "language-asian" } },
			{ new OdfStyleToCss { OdfName = "language-complex" } },
			{ new OdfStyleToCss { OdfName = "country-asian" } },
			{ new OdfStyleToCss { OdfName = "country-complex" } },
		};

		public ODTConvert(IODTFile odtFile, string linkUrlPrefix = "")
		{
			OdtFile = odtFile ?? throw new ArgumentNullException(nameof(odtFile));
			LinkUrlPrefix = linkUrlPrefix;
		}

		public ODTConvertData Convert(ODTConvertSettings settings)
		{
			ContentXDoc = OdtFile.GetZipArchiveEntryAsXDocument("content.xml");
			StylesXDoc = OdtFile.GetZipArchiveEntryAsXDocument("styles.xml");
			GetOdfStyles();
			var pageCssClassName = GetPageLayoutStyleName();
			var html = GetHtml(settings, pageCssClassName);
			var css = RenderCss();

			var data = new ODTConvertData
			{
				Css = css,
				Html = html,
				PageCssClassName = pageCssClassName,
				FirstHeaderText = GetFirstHeaderText(),
				FirstParagraphHtml = GetFirstParagraphHtml(),
				EmbedContent = EmbedContent
			};

			return data;
		}

		private string GetFirstParagraphHtml()
		{
			var firstParagraph = ContentXDoc.Root
				 .Elements(XName.Get("body", ODFXmlNamespaces.Office))
				 .Elements(XName.Get("text", ODFXmlNamespaces.Office))
				 .Elements(XName.Get("p", ODFXmlNamespaces.Text))
				 .FirstOrDefault();

			if (firstParagraph != default(XElement))
			{
				var htmlEle = new XElement(HtmlTagsTrans[firstParagraph.Name.LocalName]);
				HtmlNodesWalker(firstParagraph.Nodes(), htmlEle);
				var html = htmlEle.ToString(SaveOptions.DisableFormatting);
				return html;
			}
			else
			{
				return string.Empty;
			}
		}

		private void GetOdfStyles()
		{
			Styles = new List<IOdfStyle>();

			//var fontFaceDeclarations = ContentXDoc.Root
			//     .Elements(XName.Get("font-face-decls", ODFXmlNamespaces.Office))
			//     .Elements()
			//     .Where(p => p.Name.Equals(XName.Get("font-face", ODFXmlNamespaces.Style)));
			//StylesWalker(fontFaceDeclarations, Styles);

			var automaticStyles = ContentXDoc.Root
				  .Elements(XName.Get("automatic-styles", ODFXmlNamespaces.Office))
				  .Elements()
				  .Where(p => p.Name.Equals(XName.Get("style", ODFXmlNamespaces.Style)));
			StylesWalker(automaticStyles, Styles);

			var defaultStyles = StylesXDoc.Root
				  .Elements(XName.Get("styles", ODFXmlNamespaces.Office))
				  .Elements()
				  .Where(p => p.Name.Equals(XName.Get("default-style", ODFXmlNamespaces.Style)));
			StylesWalker(defaultStyles, Styles);

			var styles = StylesXDoc.Root
				  .Elements(XName.Get("styles", ODFXmlNamespaces.Office))
				  .Elements()
				  .Where(p => p.Name.Equals(XName.Get("style", ODFXmlNamespaces.Style)));
			StylesWalker(styles, Styles);

			var pageLayout = StylesXDoc.Root
						 .Elements(XName.Get("automatic-styles", ODFXmlNamespaces.Office))
						 .Elements()
						 .Where(p => p.Name.Equals(XName.Get("page-layout", ODFXmlNamespaces.Style)));
			StylesWalker(pageLayout, Styles);
		}

		private string RenderCss()
		{
			var builder = new StringBuilder(2048);

			foreach (var style in Styles)
			{
				if (style.Type.Equals("default-style", StringComparison.InvariantCultureIgnoreCase)
					  || string.IsNullOrWhiteSpace(style.Name))
				{
					builder.Append($"{Environment.NewLine}article > {HtmlTagsTrans[style.Family]} {{{Environment.NewLine}");
				}
				else
				{
					builder.Append($"{Environment.NewLine}.{style.Name} {{{Environment.NewLine}");
				}

				foreach (var attr in style.Attrs)
				{
					TransformStyleAttr(builder, style, attr);
				}

				foreach (var props in style.PropAttrs.Values)
				{
					foreach (var propAttr in props)
					{
						TransformStyleAttr(builder, style, propAttr);
					}
				}

				builder.Append("}");
			}

			return builder.ToString();
		}

		private static void TransformStyleAttr(StringBuilder builder, IOdfStyle style, IOdfStyleAttr attr)
		{
			var trans = CssTrans
				 .Where(p => p.OdfName.Equals(attr.Name))
				 .FirstOrDefault();

			if (trans != default(OdfStyleToCss))
			{
				if (string.IsNullOrWhiteSpace(trans.CssName))
				{
					return;
				}
				else
				{
					var attrName = trans.CssName;
					var attrVal = "";

					if (trans.Values != null && trans.Values.ContainsKey(attr.Value))
					{
						attrVal = trans.Values[attr.Value];
					}
					else
					{
						attrVal = attr.Value;
					}

					builder.Append($"{attrName}: {attrVal};{Environment.NewLine}");
				}
			}
			else
			{
				builder.Append($"{attr.Name}: {attr.Value};{Environment.NewLine}");
			}
		}

		private void StylesWalker(IEnumerable<XElement> elements, List<IOdfStyle> styles)
		{
			foreach (var ele in elements)
			{
				IOdfStyle style = new OdfStyle(ele);
				styles.Add(style);
				StylePropertyWalker(ele.Elements(), style);
			}
		}

		private void StylePropertyWalker(IEnumerable<XElement> elements, IOdfStyle style)
		{
			foreach (var ele in elements.Where(p => p.Name.LocalName.EndsWith("-properties")))
			{
				style.AddPropertyAttributes(ele);
				StylePropertyWalker(ele.Elements(), style);
			}
		}

		private string GetHtml(ODTConvertSettings settings, string pageCssClassName)
		{
			var contentEle = ContentXDoc.Root
					  .Elements(XName.Get("body", ODFXmlNamespaces.Office))
					  .Elements(XName.Get("text", ODFXmlNamespaces.Office))
					  .First();


			XElement htmlEle = null;

			if (settings.FluidWidth)
			{
				htmlEle = new XElement(HtmlTagsTrans[contentEle.Name.LocalName]);
			}
			else
			{
				htmlEle = new XElement(HtmlTagsTrans[contentEle.Name.LocalName], new XAttribute("class", pageCssClassName));
			}

			HtmlNodesWalker(contentEle.Nodes(), htmlEle);

			var html = htmlEle.ToString(SaveOptions.DisableFormatting);
			return html;
		}

		private string GetPageLayoutStyleName()
		{
			IOdfStyle pageLayoutStyle = Styles.Where(p => p.Type.Equals("page-layout")).FirstOrDefault();

			if (pageLayoutStyle != default(IOdfStyle))
			{
				foreach (var propAttr in pageLayoutStyle.PropAttrs)
				{
					foreach (var attrs in propAttr.Value)
					{
						attrs.Name = attrs.Name.Replace("margin", "padding");
					}
				}

				return pageLayoutStyle.Name;
			}
			else
			{
				return null;
			}
		}

		private string GetFirstHeaderText()
		{
			var firstHeader = ContentXDoc.Root
				 .Elements(XName.Get("body", ODFXmlNamespaces.Office))
				 .Elements(XName.Get("text", ODFXmlNamespaces.Office))
				 .Elements(XName.Get("h", ODFXmlNamespaces.Text))
				 .FirstOrDefault();

			if (firstHeader != default(XElement))
			{
				return ODTReader.GetValue(firstHeader);
			}
			else
			{
				return String.Empty;
			}
		}

		private void HtmlNodesWalker(IEnumerable<XNode> odNode, XElement htmlElement)
		{
			var childHtmlEle = htmlElement;

			foreach (var node in odNode)
			{
				if (node.NodeType == XmlNodeType.Text)
				{
					var textNode = node as XText;
					childHtmlEle.SetValue(childHtmlEle.Value + textNode.Value);
				}
				else if (node.NodeType == XmlNodeType.Element)
				{
					var elementNode = node as XElement;

					if (elementNode.Name.Equals(XName.Get("s", ODFXmlNamespaces.Text)))
					{
						AddNbsp(elementNode, htmlElement);
					}
					else if (HtmlTagsTrans.TryGetValue(elementNode.Name.LocalName, out string htmlTag))
					{
						childHtmlEle = new XElement(htmlTag);
						CopyAttributes(elementNode, childHtmlEle);
						AddStyles(htmlTag, childHtmlEle);
						htmlElement.Add(childHtmlEle);
						HtmlNodesWalker(elementNode.Nodes(), childHtmlEle);
					}
				}
			}
		}

		private void AddStyles(string htmlTag, XElement childHtmlEle)
		{
			if (htmlTag.Equals("img", StringComparison.InvariantCultureIgnoreCase))
			{
				var htmlStyleAttr = childHtmlEle.Attributes().Where(p => p.Name.LocalName.Equals("style")).FirstOrDefault();

				if (htmlStyleAttr == default(XAttribute))
				{
					htmlStyleAttr = new XAttribute("style", "width:100%;height:100%");
					childHtmlEle.Add(htmlStyleAttr);
				}
				else
				{
					var styles = htmlStyleAttr.Value.Split(';');

					for (int i = 0; i < styles.Length; i++)
					{
						styles[i] = styles[i].Trim();
						var nameVal = styles[i].Split(':');
						nameVal[0] = nameVal[0].Trim();

						if (nameVal[0].Equals("width", StringComparison.InvariantCultureIgnoreCase))
						{
							styles[i] = "width:100%";
						}
						else if (nameVal[0].Equals("height", StringComparison.InvariantCultureIgnoreCase))
						{
							styles[i] = "height:100%";
						}
					}

					htmlStyleAttr.Value = "";

					foreach (var style in styles)
					{
						htmlStyleAttr.Value += $"{style};";
					}
				}
			}
		}

		private static void AddNbsp(XElement odElement, XElement htmlElement)
		{
			var spacesValue = odElement.Attribute(XName.Get("c", ODFXmlNamespaces.Text))?.Value;
			int.TryParse(spacesValue, out int spacesCount);

			if (spacesCount == 0)
			{
				spacesCount++;
			}

			for (int i = 0; i < spacesCount; i++)
			{
				htmlElement.SetValue(htmlElement.Value + "&nbsp;");
			}
		}

		private void CopyAttributes(XElement odElement, XElement htmlElement)
		{
			if (odElement.HasAttributes)
			{
				foreach (var attr in odElement.Attributes())
				{
					var attrName = $"{odElement.Name.LocalName}.{attr.Name.LocalName}";

					if (HtmlAttrTrans.TryGetValue(attrName, out string htmlAttrName))
					{
						var attrVal = attr.Value;

						if (htmlAttrName.Equals("src", StringComparison.InvariantCultureIgnoreCase))
						{
							attrVal = GetEmbedContent(attr.Value);
						}

						var htmlAttr = new XAttribute(htmlAttrName, attrVal);
						htmlElement.Add(htmlAttr);
					}
					else if (StyleAttr.TryGetValue(attrName, out string styleAttrName))
					{
						var styleAttr = htmlElement.Attributes().Where(p => p.Name.LocalName.Equals("style")).FirstOrDefault();
						var attrVal = $"{styleAttrName}:{attr.Value};";

						if (styleAttr == default(XAttribute))
						{
							var htmlAttr = new XAttribute("style", attrVal);
							htmlElement.Add(htmlAttr);
						}
						else
						{
							styleAttr.Value += attrVal;
						}
					}
				}
			}
		}

		private string GetEmbedContent(string link)
		{
			var id = Guid.NewGuid();
			var name = $"{LinkUrlPrefix}/{id.ToString()}/{link.Replace('/', '_')}";
			var fileContent = OdtFile.GetZipArchiveEntryFileContent(link);

			var content = new ODFEmbedContent
			{
				Id = id,
				Name = name,
				Content = fileContent
			};

			EmbedContent.Add(content);

			return name;
		}
	}
}
