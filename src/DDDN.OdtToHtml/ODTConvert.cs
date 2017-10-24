/*
DDDN.OdtToHtml.OdtConvert
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
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DDDN.OdtToHtml
{
	public class OdtConvert : IOdtConvert
	{
		private readonly IOdtFile OdtFile;
		private readonly XDocument ContentXDoc;
		private readonly XDocument StylesXDoc;
		private readonly List<OdtEmbedContent> EmbedContent = new List<OdtEmbedContent>();
		private List<OdtStyle> Styles;

		public OdtConvert(IOdtFile odtFile)
		{
			OdtFile = odtFile ?? throw new ArgumentNullException(nameof(odtFile));

			ContentXDoc = OdtFile.GetZipArchiveEntryAsXDocument("content.xml");
			StylesXDoc = OdtFile.GetZipArchiveEntryAsXDocument("styles.xml");
		}

		public OdtConvertData Convert(OdtConvertSettings convertSettings)
		{
			if (convertSettings == null)
				throw new ArgumentNullException(nameof(convertSettings));

			EmbedContent.Clear();

			GetOdtStyles();
			var (width, height) = GetPageInfo();

			OdtHttpNode htmlTreeRootTag = GetHtml(convertSettings);
			var css = RenderCss(convertSettings);

			var html = OdtHttpNode.RenderHtml(htmlTreeRootTag);

			return new OdtConvertData
			{
				Css = css,
				Html = html,
				PageHeight = width,
				PageWidth = height,
				DocumentFirstHeader = GetFirstHeaderText(),
				DocumentFirstParagraph = GetFirstParagraphText(),
				EmbedContent = EmbedContent
			};
		}

		private void GetOdtStyles()
		{
			Styles = new List<OdtStyle>();

			var stylesAutomaticStyles = StylesXDoc.Root
				  .Elements(XName.Get(OdtStyleType.AutomaticStyles, OdtXmlNamespaces.Office))
				  .Elements();
			OdtStylesWalker(stylesAutomaticStyles, Styles);

			var stylesStyles = StylesXDoc.Root
				  .Elements(XName.Get(OdtStyleType.Styles, OdtXmlNamespaces.Office))
				  .Elements();
			OdtStylesWalker(stylesStyles, Styles);

			var contentAutomaticStyles = ContentXDoc.Root
				  .Elements(XName.Get(OdtStyleType.AutomaticStyles, OdtXmlNamespaces.Office))
				  .Elements();
			OdtStylesWalker(contentAutomaticStyles, Styles);

			var contentStyles = ContentXDoc.Root
				  .Elements(XName.Get(OdtStyleType.Styles, OdtXmlNamespaces.Office))
				  .Elements();
			OdtStylesWalker(contentStyles, Styles);

			TransformLevelsToStyles();
		}

		private void TransformLevelsToStyles()
		{
			var levelStyles = new List<OdtStyle>();

			foreach (var style in Styles.Where(p => p.Levels.Count > 0))
			{
				var levelCount = 1;

				foreach (var level in style.Levels)
				{
					var odtStyle = new OdtStyle()
					{
						Name = $"{style.Name}LVL{levelCount}",
						Attributes = style.Levels[levelCount]
					};

					levelStyles.Add(odtStyle);
					levelCount++;
				}
			}

			Styles.AddRange(levelStyles);
		}

		private void CopyAttributes(
			XElement odElement,
			OdtHttpNode htmlElement,
			OdtConvertSettings convertSettings,
			int level = 0,
			string levelParentClassName = "")
		{
			if (level > 0 && htmlElement.Name.Equals("ul", StringComparison.InvariantCultureIgnoreCase))
			{
				htmlElement.AddAttrValue("class", levelParentClassName);
			}

			foreach (var attr in odElement.Attributes())
			{
				var attrName = $"{odElement.Name.LocalName}.{attr.Name.LocalName}";

				if (OdtTrans.OdtNodeAttrToHtmlNodeAttr.TryGetValue(attrName, out string htmlAttrName))
				{
					var attrVal = attr.Value;

					if (htmlAttrName.Equals("src", StringComparison.InvariantCultureIgnoreCase))
					{
						attrVal = GetEmbedContent(attr.Value, convertSettings);
					}

					htmlElement.AddAttrValue(htmlAttrName, attrVal);
				}
				else if (OdtTrans.OdtNodeAttrToHtmlNodeInlineStyle.TryGetValue(attrName, out string styleAttrName))
				{
					htmlElement.AddAttrValue("style", $"{styleAttrName}:{attr.Value};");
				}
			}

			if (level > 0
				&& htmlElement.Parent.Name.Equals("ul", StringComparison.InvariantCultureIgnoreCase)
				&& htmlElement.Name.Equals("li", StringComparison.InvariantCultureIgnoreCase))
			{
				htmlElement.AddAttrValue("class", $"{levelParentClassName}LVL{level}");
			}
		}

		private void OdtStylesWalker(IEnumerable<XElement> elements, List<OdtStyle> styles)
		{
			foreach (var ele in elements)
			{
				var nameAttr = ele.Attributes()
					.FirstOrDefault(p => p.Name.LocalName.Equals("name", StringComparison.InvariantCultureIgnoreCase))?
					.Value;

				var odtStyle = styles.Find(p => 0 == string.Compare(p.Name, nameAttr, StringComparison.InvariantCultureIgnoreCase));

				if (odtStyle == default(OdtStyle)
					|| ele.Name.LocalName.Equals(OdtStyleType.DefaultStyle, StringComparison.InvariantCultureIgnoreCase))
				{
					odtStyle = new OdtStyle
					{
						Type = ele.Name.LocalName,
					};

					styles.Add(odtStyle);
					AttributeWalker(ele, odtStyle);
					ChildWalker(ele.Elements(), odtStyle);
				}
			}
		}

		private void ChildWalker(IEnumerable<XElement> elements, OdtStyle odtStyle, int level = 0)
		{
			foreach (var ele in elements)
			{
				var levelAttr = ele
					.Attributes()
					.FirstOrDefault(p => p.Name.LocalName.Equals("level", StringComparison.InvariantCultureIgnoreCase));

				if (levelAttr == default(XAttribute))
				{
					AttributeWalker(ele, odtStyle, level);
					ChildWalker(ele.Elements(), odtStyle, level);
				}
				else
				{
					var levelNo = int.Parse(levelAttr.Value);
					AttributeWalker(ele, odtStyle, levelNo);
					ChildWalker(ele.Elements(), odtStyle, levelNo);
				}
			}
		}

		private void AttributeWalker(XElement element, OdtStyle odtStyle, int level = 0)
		{
			foreach (var eleAttr in element.Attributes())
			{
				var newAttr = new OdtStyleAttr()
				{
					OdtType = element.Name.LocalName,
					Name = eleAttr.Name.LocalName,
					Value = eleAttr.Value
				};

				if (level == 0)
				{
					HandleSpecialAttr(eleAttr, odtStyle);
					odtStyle.Attributes.Add(newAttr);
				}
				else
				{
					if (odtStyle.Levels.ContainsKey(level))
					{
						odtStyle.Levels[level].Add(newAttr);
					}
					else
					{
						var attrList = new List<OdtStyleAttr>
						{
							newAttr
						};

						odtStyle.Levels.Add(level, attrList);
					}
				}
			}
		}

		private string RenderCss(OdtConvertSettings convertSettings)
		{
			var builder = new StringBuilder(8192);
			var cssRootIdentifyer = string.IsNullOrWhiteSpace(convertSettings.RootElementId) ? convertSettings.RootElementTagName : $"#{convertSettings.RootElementId}";

			foreach (var style in Styles)
			{
				if (style.Type.Equals(OdtStyleType.DefaultStyle, StringComparison.InvariantCultureIgnoreCase)
					  && OdtTrans.OdtTagNameToHtmlTagName.Exists(p => p.OdtName.Equals(style.Family, StringComparison.InvariantCultureIgnoreCase)))
				{
					builder
						.Append(Environment.NewLine)
						.Append(cssRootIdentifyer)
						.Append(" ")
						.Append(OdtTrans.OdtTagNameToHtmlTagName.Find(p => p.OdtName.Equals(style.Family, StringComparison.InvariantCultureIgnoreCase)).HtmlName)
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

		private void StyleAttrWalker(StringBuilder builder, OdtStyle style)
		{
			if (!string.IsNullOrWhiteSpace(style.ParentStyleName))
			{
				var parentStyle = Styles
					.Find(p => 0 == string.Compare(p.Name, style.ParentStyleName, StringComparison.CurrentCultureIgnoreCase));

				if (parentStyle != default(OdtStyle))
				{
					StyleAttrWalker(builder, parentStyle);
				}
			}

			foreach (var attr in style.Attributes)
			{
				TransformStyleAttr(builder, attr);
			}
		}

		private static void TransformStyleAttr(StringBuilder builder, OdtStyleAttr attr)
		{
			var trans = OdtTrans.OdtStyleToCssStyle
				.Find(p =>
					p.OdtName.Equals(attr.Name, StringComparison.InvariantCultureIgnoreCase)
					&& p.OdtTypes.Contains(attr.OdtType));

			if (trans == null)
			{
				return;
			}

			var attrVal = attr.Value;

			if (trans.Values?.ContainsKey(attr.Value) == true)
			{
				attrVal = trans.Values[attr.Value];
			}

			builder
				.Append(trans.CssName)
				.Append(": ")
				.Append(attrVal)
				.Append(";");
			if (Debugger.IsAttached)
			{
				builder.Append(" /* ")
				.Append(attr?.OdtType)
				.Append(" */ ");
			}
			builder.Append(Environment.NewLine);
		}

		private OdtHttpNode GetHtml(OdtConvertSettings convertSettings)
		{
			var documentNodes = ContentXDoc.Root
					  .Elements(XName.Get("body", OdtXmlNamespaces.Office))
					  .Elements(XName.Get("text", OdtXmlNamespaces.Office))
					  .Nodes();

			var rootNode = new OdtHttpNode(convertSettings.RootElementTagName);
			if (!string.IsNullOrWhiteSpace(convertSettings.RootElementId))
			{
				rootNode.AddAttrValue("id", convertSettings.RootElementId);
			}

			OdtBodyNodesWalker(documentNodes, rootNode, convertSettings);

			return rootNode;
		}

		private void AddDefaultInlineStyles(OdtHttpNode htmlTag)
		{
			if (htmlTag.Name.Equals("img", StringComparison.InvariantCultureIgnoreCase))
			{
				htmlTag.AddAttrValue("style", "max-width: 100%;");
				htmlTag.AddAttrValue("style", "height: auto;");
			}
		}

		private void OdtBodyNodesWalker(
			IEnumerable<XNode> odNodes,
			OdtHttpNode parentHtmlNodeTagName,
			OdtConvertSettings convertSettings,
			int level = 0,
			string levelParentClassName = "")
		{
			foreach (var node in odNodes)
			{
				if (node.NodeType == XmlNodeType.Text)
				{
					parentHtmlNodeTagName.Inner.Add(((XText)node).Value);
				}
				else if (node.NodeType == XmlNodeType.Element)
				{
					var elementNode = node as XElement;

					if (elementNode.Name.Equals(XName.Get("s", OdtXmlNamespaces.Text)))
					{
						AddNbsp(elementNode, parentHtmlNodeTagName);
					}
					else if (OdtTrans
						.OdtTagNameToHtmlTagName
						.Exists(p => p.OdtName.Equals(elementNode.Name.LocalName, StringComparison.InvariantCultureIgnoreCase)))
					{
						var htmlTag = OdtTrans.OdtTagNameToHtmlTagName.Find(p => p
							.OdtName
							.Equals(elementNode.Name.LocalName, StringComparison.InvariantCultureIgnoreCase))?.HtmlName;
						var htmlElement = new OdtHttpNode(htmlTag, level, parentHtmlNodeTagName);
						CopyAttributes(elementNode, htmlElement, convertSettings, level, levelParentClassName);
						AddDefaultInlineStyles(htmlElement);
						var isLevelParentElement = OdtTrans.LevelParent
							.Contains(elementNode.Name.LocalName, StringComparer.InvariantCultureIgnoreCase);

						if (isLevelParentElement)
						{
							if (level == 0)
							{
								levelParentClassName = elementNode
								.Attributes(XName.Get("style-name", OdtXmlNamespaces.Text))
								.FirstOrDefault()?.Value;
							}

							level++;
						}

						OdtBodyNodesWalker(elementNode.Nodes(), htmlElement, convertSettings, level, levelParentClassName);

						if (isLevelParentElement)
						{
							level--;
						}
					}
				}
			}
		}

		private static void AddNbsp(XElement odElement, OdtHttpNode htmlElement)
		{
			var spacesValue = odElement.Attribute(XName.Get("c", OdtXmlNamespaces.Text))?.Value;
			int.TryParse(spacesValue, out int spacesCount);

			if (spacesCount == 0)
			{
				spacesCount++;
			}

			for (int i = 0; i < spacesCount; i++)
			{
				htmlElement.Inner.Add("&nbsp;");
			}
		}

		private string GetEmbedContent(string odtAttrlink, OdtConvertSettings convertSettings)
		{
			var id = Guid.NewGuid();
			var name = $"{id}_{odtAttrlink.Replace('/', '_')}";
			var link = name;

			if (!string.IsNullOrWhiteSpace(convertSettings.LinkUrlPrefix))
			{
				if (convertSettings.LinkUrlPrefix.EndsWith("/", StringComparison.InvariantCultureIgnoreCase))
				{
					link = $"{convertSettings.LinkUrlPrefix}{link}";
				}
				else
				{
					link = $"{convertSettings.LinkUrlPrefix}/{link}";
				}
			}

			var fileContent = OdtFile.GetZipArchiveEntryFileContent(odtAttrlink);

			var content = new OdtEmbedContent
			{
				Id = id,
				Name = name,
				Link = link,
				Content = fileContent
			};

			EmbedContent.Add(content);

			return link;
		}

		private bool HandleSpecialAttr(XAttribute attr, OdtStyle odtStyle)
		{
			if (attr.Name.LocalName.Equals("name"))
			{
				odtStyle.Name = attr.Value;
			}
			else if (attr.Name.LocalName.Equals("parent-style-name"))
			{
				odtStyle.ParentStyleName = attr.Value;
			}
			else if (attr.Name.LocalName.Equals("family"))
			{
				odtStyle.Family = attr.Value;
			}
			else if (attr.Name.LocalName.Equals("list-style-name"))
			{
				odtStyle.ListStyleName = attr.Value;
			}
			else
			{
				return false;
			}

			return true;
		}

		private (string width, string height) GetPageInfo()
		{
			var pageLayoutStyle = Styles.Find(p => p.Type.Equals(OdtStyleType.PageLayout, StringComparison.InvariantCultureIgnoreCase));

			if (pageLayoutStyle == null)
			{
				return (null, null);
			}
			else
			{
				var width = pageLayoutStyle.Attributes.Find(p => p.Name
				.Equals("page-width", StringComparison.InvariantCultureIgnoreCase)
				&& p.OdtType.Equals(OdtStyleType.PageLayoutProperties, StringComparison.InvariantCultureIgnoreCase))
				.Value;
				var height = pageLayoutStyle.Attributes.Find(p => p.Name
				.Equals("page-height", StringComparison.InvariantCultureIgnoreCase)
				&& p.OdtType.Equals(OdtStyleType.PageLayoutProperties, StringComparison.InvariantCultureIgnoreCase))
				.Value;
				return (width, height);
			}
		}

		private string GetFirstHeaderText()
		{
			var headers = ContentXDoc.Root
				 .Elements(XName.Get("body", OdtXmlNamespaces.Office))
				 .Elements(XName.Get("text", OdtXmlNamespaces.Office))
				 .Elements(XName.Get("h", OdtXmlNamespaces.Text));

			foreach (var h in headers)
			{
				var inner = GetOdtElementValue(h);
				if (!string.IsNullOrWhiteSpace(inner))
				{
					return inner;
				}
			}

			return "";
		}

		private string GetFirstParagraphText()
		{
			var paragraphs = ContentXDoc.Root
				 .Elements(XName.Get("body", OdtXmlNamespaces.Office))
				 .Elements(XName.Get("text", OdtXmlNamespaces.Office))
				 .Elements(XName.Get("p", OdtXmlNamespaces.Text));

			foreach (var p in paragraphs)
			{
				var inner = GetOdtElementValue(p);
				if (!string.IsNullOrWhiteSpace(inner))
				{
					return inner;
				}
			}

			return "";
		}

		public static string GetOdtElementValue(XElement xElement)
		{
			return WalkTheNodes(xElement.Nodes());
		}

		private static string WalkTheNodes(IEnumerable<XNode> nodes)
		{
			if (nodes == null)
			{
				throw new ArgumentNullException(nameof(nodes));
			}

			string val = "";

			foreach (var node in nodes)
			{
				if (node.NodeType == XmlNodeType.Text)
				{
					var textNode = node as XText;
					val += textNode.Value;
				}
				else if (node.NodeType == XmlNodeType.Element)
				{
					var elementNode = node as XElement;

					if (elementNode.Name.Equals(XName.Get("s", OdtXmlNamespaces.Text)))
					{
						val += " ";
					}
					else
					{
						val += WalkTheNodes(elementNode.Nodes());
					}
				}
			}

			return val;
		}
	}
}
