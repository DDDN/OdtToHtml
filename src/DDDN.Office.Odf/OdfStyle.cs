/*
DDDN.Office.Odf.OdfStyle
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
using System.Xml.Linq;

namespace DDDN.Office.Odf
{
	public class OdfStyle : IOdfStyle
	{
		public string Name { get; set; }
		public string Type { get; set; }
		public string NamespaceName { get; set; }
		public string ParentStyleName { get; set; }
		public string Family { get; set; }
		public List<OdfStyleAttr> Attrs { get; set; } = new List<OdfStyleAttr>();
		public Dictionary<string, List<OdfStylePropAttr>> PropAttrs { get; set; } = new Dictionary<string, List<OdfStylePropAttr>>();

		public OdfStyle(XElement element)
		{
			if (element == null)
			{
				throw new ArgumentNullException(nameof(element));
			}

			Type = element.Name.LocalName;
			NamespaceName = element.Name.NamespaceName;
			AttrWalker(element);
		}

		public void AddPropertyAttributes(XElement propertyElement)
		{
			if (propertyElement == null)
			{
				throw new ArgumentNullException(nameof(propertyElement));
			}

			PropertyAttrWalker(propertyElement);
		}

		private void AttrWalker(XElement element)
		{
			foreach (var eleAttr in element.Attributes())
			{
				if (!HandleSpecialAttr(eleAttr))
				{
					var attr = new OdfStyleAttr()
					{
						Name = eleAttr.Name.LocalName,
						Value = eleAttr.Value
					};
					Attrs.Add(attr);
				}
			}
		}

		private void PropertyAttrWalker(XElement element)
		{
			foreach (var eleAttr in element.Attributes())
			{
				var propAttr = new OdfStylePropAttr()
				{
					Name = eleAttr.Name.LocalName,
					Value = eleAttr.Value
				};

				if (PropAttrs.ContainsKey(element.Name.LocalName))
				{
					PropAttrs[element.Name.LocalName].Add(propAttr);
				}
				else
				{
					var propAttrList = new List<OdfStylePropAttr> { propAttr };
					PropAttrs.Add(element.Name.LocalName, propAttrList);
				}
			}
		}

		private bool HandleSpecialAttr(XAttribute attr)
		{
			if (attr.Name.LocalName.Equals("name"))
			{
				Name = attr.Value;
			}
			else if (attr.Name.LocalName.Equals("parent-style-name"))
			{
				ParentStyleName = attr.Value;
			}
			else if (attr.Name.LocalName.Equals("family"))
			{
				Family = attr.Value;
			}
			else
			{
				return false;
			}

			return true;
		}
	}
}
