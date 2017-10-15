﻿/*
DDDN.Office.Odf.OdfStyle
Copyright(C) 2017 Lukasz Jaskiewicz(lukasz @jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System.Collections.Generic;

namespace DDDN.Office.Odf
{
	public class OdfStyle
	{
		public string Name { get; set; }
		public string Type { get; set; }
		public string NamespaceName { get; set; }
		public string ParentStyleName { get; set; }
		public string Family { get; set; }
		public List<OdfStyleAttr> Attrs { get; set; } = new List<OdfStyleAttr>();
		public Dictionary<string, List<OdfStyleAttr>> PropAttrs { get; set; } = new Dictionary<string, List<OdfStyleAttr>>();
	}
}
