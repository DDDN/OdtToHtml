/*
DDDN.OdtToHtml.OdtStyleAttr
Copyright(C) 2017-2018 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System.Diagnostics;
using System.Xml.Linq;

namespace DDDN.OdtToHtml
{
	[DebuggerDisplay("{LocalName,nq} ■ {Value,nq} ■ {PropertyName,nq}")]
	public class OdtStyleAttr
	{
		public string LocalName { get; set; }
		public string PropertyName { get; set; }
		public string Value { get; set; }

		private OdtStyleAttr()
		{

		}

		public OdtStyleAttr(XAttribute xAttr, string propName)
		{
			LocalName = xAttr.Name.LocalName;
			PropertyName = propName;
			Value = xAttr.Value;
		}
	}
}
