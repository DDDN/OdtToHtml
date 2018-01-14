﻿/*
DDDN.OdtToHtml.OdtPageInfo
Copyright(C) 2017-2018 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

namespace DDDN.OdtToHtml
{
	public class OdtPageInfo
    {
		public string WidthBrutto { get; set; }
		public string HeightBrutto { get; set; }
		public string WidthNetto { get; set; }
		public string HeightNetto { get; set; }
		public string MarginTop { get; set; }
		public string MarginRight { get; set; }
		public string MarginBottom { get; set; }
		public string MarginLeft { get; set; }
	}
}
