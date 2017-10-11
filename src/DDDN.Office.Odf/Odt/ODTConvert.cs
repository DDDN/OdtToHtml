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
		private readonly IODTFile OdtFile;
		private readonly XDocument ContentXDoc;
		private readonly XDocument StylesXDoc;
		private readonly List<ODFEmbedContent> EmbedContent = new List<ODFEmbedContent>();
		private List<IOdfStyle> Styles;

		public ODTConvert(IODTFile odtFile)
		{
			OdtFile = odtFile ?? throw new ArgumentNullException(nameof(odtFile));

			ContentXDoc = OdtFile.GetZipArchiveEntryAsXDocument("content.xml");
			StylesXDoc = OdtFile.GetZipArchiveEntryAsXDocument("styles.xml");
		}

		public ODTConvertData Convert(ODTConvertSettings convertSettings)
		{
			GetOdfStyles();
			var (width, height) = GetPageInfo();

			return new ODTConvertData
			{
				Css = RenderCss(convertSettings),
				Html = GetHtml(convertSettings),
				PageHeight = width,
				PageWidth = height,
				DocumentFirstHeader = GetFirstHeaderText(),
				DocumentFirstParagraph = GetFirstParagraphText(),
				EmbedContent = EmbedContent
			};
		}

		private string GetFirstParagraphText()
		{
			var paragraphs = ContentXDoc.Root
				 .Elements(XName.Get("body", ODFXmlNamespaces.Office))
				 .Elements(XName.Get("text", ODFXmlNamespaces.Office))
				 .Elements(XName.Get("p", ODFXmlNamespaces.Text));

			foreach (var p in paragraphs)
			{
				var inner = ODTReader.GetValue(p);
				if (!string.IsNullOrWhiteSpace(inner))
				{
					return inner;
				}
			}

			return string.Empty;
		}

		private void GetOdfStyles()
		{
			Styles = new List<IOdfStyle>();

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

		private string RenderCss(ODTConvertSettings convertSettings)
		{
			var builder = new StringBuilder(8192);

			foreach (var style in Styles)
			{
				if (style.Type.Equals("default-style", StringComparison.InvariantCultureIgnoreCase)
					  && ODTTrans.Tags.Exists(p => p.OdfName.Equals(style.Family, StringComparison.InvariantCultureIgnoreCase)))
				{
					builder
						.Append(Environment.NewLine)
						.Append($"{convertSettings.RootHtmlTag} ")
						.Append(ODTTrans.Tags.Find(p => p.OdfName.Equals(style.Family, StringComparison.InvariantCultureIgnoreCase)).HtmlName)
						.Append(" {")
						.Append(Environment.NewLine);
					StyleAttrWalker(builder, style);
					builder.Append("}");
				}
				else if (!string.IsNullOrWhiteSpace(style.Name))
				{
					builder.Append(Environment.NewLine)
						.Append(".")
						.Append(style.Name)
						.Append(" {")
						.Append(Environment.NewLine);
					StyleAttrWalker(builder, style);
					builder.Append("}");
				}
			}

			return builder.ToString();
		}

		private void StyleAttrWalker(StringBuilder builder, IOdfStyle style)
		{
			if (!string.IsNullOrWhiteSpace(style.ParentStyleName))
			{
				var parentStyle = Styles
					.Find(p => 0 == string.Compare(p.Name, style.ParentStyleName, StringComparison.CurrentCultureIgnoreCase));

				if (parentStyle != default(IOdfStyle))
				{
					StyleAttrWalker(builder, parentStyle);
				}
			}

			foreach (var attr in style.Attrs)
			{
				TransformStyleAttr(builder, attr);
			}

			foreach (var props in style.PropAttrs.Values)
			{
				foreach (var propAttr in props)
				{
					TransformStyleAttr(builder, propAttr);
				}
			}
		}

		private static void TransformStyleAttr(StringBuilder builder, IOdfStyleAttr attr)
		{
			var trans = ODTTrans.Css.Find(p => p.OdfName.Equals(attr.Name, StringComparison.InvariantCultureIgnoreCase));

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

					if (trans.Values?.ContainsKey(attr.Value) == true)
					{
						attrVal = trans.Values[attr.Value];
					}
					else
					{
						attrVal = attr.Value;
					}

					builder
						.Append(attrName)
						.Append(": ")
						.Append(attrVal)
						.Append(";")
						.Append(Environment.NewLine);
				}
			}
			else
			{
				builder
					.Append(attr.Name)
					.Append(": ")
					.Append(attr.Value)
					.Append(";")
					.Append(Environment.NewLine);
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

		public void StylePropertyWalker(IEnumerable<XElement> elements, IOdfStyle style)
		{
			foreach (var ele in elements.Where(p => p.Name.LocalName.EndsWith("-properties")))
			{
				style.AddPropertyAttributes(ele);
				StylePropertyWalker(ele.Elements(), style);
			}
		}

		private string GetHtml(ODTConvertSettings convertSettings)
		{
			var documentNodes = ContentXDoc.Root
					  .Elements(XName.Get("body", ODFXmlNamespaces.Office))
					  .Elements(XName.Get("text", ODFXmlNamespaces.Office))
					  .Nodes();

			var rootElement = new XElement
			(
				convertSettings.RootHtmlTag,
				new XAttribute("class", convertSettings.RootHtmlTagClassNames ?? ""),
				new XAttribute("id", convertSettings.RootHtmlTagId ?? "")
			);
			HtmlNodesWalker(documentNodes, rootElement, convertSettings);
			return rootElement.ToString(SaveOptions.DisableFormatting);
		}

		private (string width, string height) GetPageInfo()
		{
			var pageLayoutStyle = Styles.Find(p => p.Type.Equals("page-layout", StringComparison.InvariantCultureIgnoreCase));

			if (pageLayoutStyle == null)
			{
				return (null, null);
			}
			else
			{
				var width = pageLayoutStyle.PropAttrs["page-layout-properties"].FirstOrDefault(p => p.Name.Equals("page-width", StringComparison.InvariantCultureIgnoreCase)).Value;
				var height = pageLayoutStyle.PropAttrs["page-layout-properties"].FirstOrDefault(p => p.Name.Equals("page-height", StringComparison.InvariantCultureIgnoreCase)).Value;
				return (width, height);
			}
		}

		private string GetFirstHeaderText()
		{
			var headers = ContentXDoc.Root
				 .Elements(XName.Get("body", ODFXmlNamespaces.Office))
				 .Elements(XName.Get("text", ODFXmlNamespaces.Office))
				 .Elements(XName.Get("h", ODFXmlNamespaces.Text));

			foreach (var h in headers)
			{
				var inner = ODTReader.GetValue(h);
				if (!string.IsNullOrWhiteSpace(inner))
				{
					return inner;
				}
			}

			return string.Empty;
		}

		private void HtmlNodesWalker(IEnumerable<XNode> odNode, XElement htmlElement, ODTConvertSettings convertSettings)
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
					else if (ODTTrans.Tags.Exists(p => p.OdfName.Equals(elementNode.Name.LocalName, StringComparison.InvariantCultureIgnoreCase)))
					{
						var htmlTag = ODTTrans.Tags.Find(p => p.OdfName.Equals(elementNode.Name.LocalName, StringComparison.InvariantCultureIgnoreCase))?.HtmlName;
						childHtmlEle = new XElement(htmlTag);
						CopyAttributes(elementNode, childHtmlEle, convertSettings);
						AddInlineStyles(htmlTag, childHtmlEle);
						htmlElement.Add(childHtmlEle);
						HtmlNodesWalker(elementNode.Nodes(), childHtmlEle, convertSettings);
					}
				}
			}
		}

		private void AddInlineStyles(string htmlTag, XElement childHtmlEle)
		{
			if (htmlTag.Equals("img", StringComparison.InvariantCultureIgnoreCase))
			{
				var htmlStyleAttr = childHtmlEle.Attributes().FirstOrDefault(p => p.Name.LocalName.Equals("style", StringComparison.InvariantCultureIgnoreCase));

				if (htmlStyleAttr == default(XAttribute))
				{
					htmlStyleAttr = new XAttribute("style", "max-width: 100%; height: auto;");
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
							styles[i] = "max-width:100%";
						}
						else if (nameVal[0].Equals("height", StringComparison.InvariantCultureIgnoreCase))
						{
							styles[i] = "height:auto";
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

		private void CopyAttributes(XElement odElement, XElement htmlElement, ODTConvertSettings convertSettings)
		{
			if (odElement.HasAttributes)
			{
				foreach (var attr in odElement.Attributes())
				{
					var attrName = $"{odElement.Name.LocalName}.{attr.Name.LocalName}";

					if (ODTTrans.Attrs.TryGetValue(attrName, out string htmlAttrName))
					{
						var attrVal = attr.Value;

						if (htmlAttrName.Equals("src", StringComparison.InvariantCultureIgnoreCase))
						{
							attrVal = GetEmbedContent(attr.Value, convertSettings);
						}

						var htmlAttr = new XAttribute(htmlAttrName, attrVal);
						htmlElement.Add(htmlAttr);
					}
					else if (ODTTrans.StyleAttr.TryGetValue(attrName, out string styleAttrName))
					{
						var styleAttr = htmlElement.Attributes().FirstOrDefault(p => p.Name.LocalName.Equals("style", StringComparison.InvariantCultureIgnoreCase));
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

		private string GetEmbedContent(string link, ODTConvertSettings convertSettings)
		{
			var id = Guid.NewGuid();
			var name = $"{convertSettings.LinkUrlPrefix}/{id.ToString()}/{link.Replace('/', '_')}";
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
