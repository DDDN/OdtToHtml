/*
DDDN.Office.Odf.OdfHttpNode
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
using System.Text;

namespace DDDN.Office.Odf
{
	public class OdfHttpNode
	{
		private OdfHttpNode() { }

		public OdfHttpNode(string tagName, int level, OdfHttpNode parent = null) : this(tagName, parent)
		{
			Level = level;
		}

		public OdfHttpNode(string tagName, OdfHttpNode parent = null)
		{
			if (string.IsNullOrWhiteSpace(tagName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(tagName));
			}

			Name = tagName;
			Parent = parent;

			if (Parent != null)
			{
				parent.Childs.Add(this);
			}
		}

		public string Name { get; }
		public List<string> Inner { get; set; } = new List<string>();
		public OdfHttpNode Parent { get; }
		public int Level { get; }
		public List<OdfHttpNode> Childs { get; set; } = new List<OdfHttpNode>();
		public Dictionary<string, List<string>> Attrs { get; } = new Dictionary<string, List<string>>();

		public void AddAttrValue(string attrName, string attrVal)
		{
			if (attrName == null)
				throw new ArgumentNullException(nameof(attrName));

			if (Attrs.ContainsKey(attrName))
			{
				Attrs[attrName].Add(attrVal);
			}
			else
			{
				Attrs[attrName] = new List<string> { attrVal };
			}
		}

		public static string RenderHtml(OdfHttpNode httpTag)
		{
			if (httpTag == null)
				throw new ArgumentNullException(nameof(httpTag));

			var htmlBuilder = new StringBuilder(8192);
			HtmlNodeWalker(htmlBuilder, httpTag);
			return htmlBuilder.ToString();
		}

		private static void HtmlNodeWalker(StringBuilder builder, OdfHttpNode httpTag)
		{
			if (httpTag.Childs.Count > 0 || httpTag.Inner.Count > 0)
			{
				builder.Append("<")
					.Append(httpTag.Name)
					.Append(RenderAtributes(httpTag))
					.Append(">")
					.Append(String.Concat(httpTag.Inner));

				foreach (var child in httpTag.Childs)
				{
					HtmlNodeWalker(builder, child);
				}

				builder.Append("</")
					.Append(httpTag.Name)
					.Append(">");
			}
			else
			{
				builder.Append("<")
					.Append(httpTag.Name)
					.Append(RenderAtributes(httpTag))
					.Append("/>");
			}
		}

		private static string RenderAtributes(OdfHttpNode httpTag)
		{
			var builder = new StringBuilder(96);

			foreach (var attr in httpTag.Attrs)
			{
				builder
					.Append(" ")
					.Append(attr.Key)
					.Append("=\"");
				bool firstName = true;

				foreach (var val in attr.Value)
				{
					if (firstName)
					{
						builder.Append(val);
						firstName = false;
					}
					else
					{
						builder
							.Append(" ")
							.Append(val);
					}
				}

				builder.Append("\"");
			}

			return builder.ToString();
		}
	}
}
