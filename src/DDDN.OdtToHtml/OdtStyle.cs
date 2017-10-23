/*
DDDN.OdtToHtml.OdtStyle
Copyright(C) 2017 Lukasz Jaskiewicz(lukasz @jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System.Collections.Generic;

namespace DDDN.OdtToHtml
{
	public class OdtStyle
	{
		public string Name { get; set; } = "";
		public string Type { get; set; } = "";
		public string ParentStyleName { get; set; } = "";
		public string ListStyleName { get; set; } = "";
		public string Family { get; set; } = "";
		public Dictionary<int, List<OdtStyleAttr>> Levels { get; set; } = new Dictionary<int, List<OdtStyleAttr>>();
		public List<OdtStyleAttr> Attributes { get; set; } = new List<OdtStyleAttr>();
	}
}
