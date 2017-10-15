﻿/*
DDDN.Office.Odf.Odt.OdtConvertData
Copyright(C) 2017 Lukasz Jaskiewicz(lukasz @jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System.Collections.Generic;

namespace DDDN.Office.Odf.Odt
{
	public class OdtConvertData
	{
		public string Html { get; set; }
		public string Css { get; set; }
		public string PageWidth { get; set; }
		public string PageHeight { get; set; }
		public string DocumentFirstHeader { get; set; }
		public string DocumentFirstParagraph { get; set; }
		public IEnumerable<OdfEmbedContent> EmbedContent { get; set; }
	}
}
