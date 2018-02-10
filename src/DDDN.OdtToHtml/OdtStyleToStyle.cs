﻿/*
DDDN.OdtToHtml.OdtStyleToStyle
Copyright(C) 2017-2018 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System.Collections.Generic;

namespace DDDN.OdtToHtml
{
	public class OdtStyleToStyle
	{
		public enum RelativeTo
		{
			None,
			Width,
			Height
		}

		public List<string> OdtAttrNames { get; set; } = new List<string>();
		public string CssPropName { get; set; }
		public List<string> OverridableBy { get; set; } = new List<string>();
		public RelativeTo AsPercentageTo { get; set; }
		public List<OdtTransValueToValue> ValueToValue { get; set; } = new List<OdtTransValueToValue>();
	}
}
