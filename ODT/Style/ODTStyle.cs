/*
* DDDN.Office.ODT.Style.ODTStyle
* 
* Copyright(C) 2017 Lukasz Jaskiewicz
* Author: Lukasz Jaskiewicz (lukasz@jaskiewicz.de, devdone@outlook.com)
*
* This program is free software; you can redistribute it and/or modify it under the terms of the
* GNU General Public License as published by the Free Software Foundation; version 2 of the License.
*
* This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
* warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License along with this program; if not, write
* to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace DDDN.Office.ODT.Style
{
	public class ODTStyle : IODTStyle
	{
		private static readonly Dictionary<string, string> AttrNames = new Dictionary<string, string>()
		{
			["name"] = "Name"
		};

		public string Name { get; set; }
		public string NamespaceName { get; set; }
		public string ParentStyleName { get; set; }
		public string Family { get; set; }
		public XElement ODTElement { get; set; }
		public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>();

		public ODTStyle(XElement element)
		{
			if (element == null)
			{
				throw new ArgumentNullException(nameof(element));
			}

			AttrWalker(element);
		}

		public void AddAttributes(XElement element)
		{
			if (element == null)
			{
				throw new ArgumentNullException(nameof(element));
			}

			AttrWalker(element);
		}

		private void AttrWalker(XElement element)
		{
			NamespaceName = element.Name.NamespaceName;

			foreach (var attr in element.Attributes())
			{
				if (!HandleSpecialAttr(attr))
				{
					Attributes.Add(attr.Name.LocalName, attr.Value);
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

				if (Family.Equals("table", StringComparison.InvariantCultureIgnoreCase))
				{
					Attributes.Add("border-collapse", "collapse");
				}
			}
			else
			{
				return false;
			}

			return true;
		}
	}
}
