/*
DDDN.Office.Odf.Odt.ODTReader
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
using System.Xml;
using System.Xml.Linq;

namespace DDDN.Office.Odf.Odt
{
	public static class OdtReader
	{
		public static string GetValue(XElement xElement)
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

					if (elementNode.Name.Equals(XName.Get("s", OdfXmlNamespaces.Text)))
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
