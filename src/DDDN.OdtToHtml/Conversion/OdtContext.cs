/*
DDDN.OdtToHtml.Conversion.OdtContext
Copyright(C) 2017-2018 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace DDDN.OdtToHtml.Conversion
{
	public class OdtContext
	{
		public IEnumerable<XNode> DocumentNodes { get; set; }
		public IEnumerable<XElement> OdtStyles { get; set; }

		public Dictionary<string, Dictionary<int, OdtListLevel>> OdtListsLevelInfo { get; set; }
			= new Dictionary<string, Dictionary<int, OdtListLevel>>(StringComparer.InvariantCultureIgnoreCase);

		public IEnumerable<OdtEmbedContent> EmbedContent { get; set; }
		public OdtConvertSettings ConvertSettings { get; set; }
		public OdtPageInfoCalc PageInfoCalc { get; set; }
		public OdtPageInfo PageInfo { get; set; }

		public Dictionary<IOdtHtmlNode, List<OdtHtmlInfo>> Lists { get; set; }
			= new Dictionary<IOdtHtmlNode, List<OdtHtmlInfo>>();

		public List<string> UsedFontFamilies { get; set; } = new List<string>();
	}
}
