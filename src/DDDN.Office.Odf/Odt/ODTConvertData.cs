/*
* DDDN.Office.Odf.Odt.ODTConvertData
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

using System.Collections.Generic;

namespace DDDN.Office.Odf.Odt
{
	public class ODTConvertData
	{
		public string Html { get; set; }
		public string Css { get; set; }
		public string FirstHeaderText { get; set; }
		public string FirstParagraphHtml { get; set; }
		public Dictionary<string, byte[]> EmbedMedia { get; set; }
	}
}
