/*
DDDN.OdtToHtml.OdtHtmlText
Copyright(C) 2017-2018 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Linq;
using System.Xml.Linq;

namespace DDDN.OdtToHtml
{
	internal class OdtHtmlText : IOdtHtmlNode
	{
		public XNode OdtNode { get; }
		public string InnerText { get; }
		public OdtHtmlInfo ParentNode { get; }
		public IOdtHtmlNode PreviousSibling { get; }

		private OdtHtmlText()
		{
		}

		public OdtHtmlText(string innerText, XNode xNode, OdtHtmlInfo parentNode)
		{
			OdtNode = xNode ?? throw new ArgumentNullException(nameof(xNode));
			InnerText = innerText ?? throw new ArgumentNullException(nameof(innerText));
			ParentNode = parentNode ?? throw new ArgumentNullException(nameof(parentNode));
			PreviousSibling = ParentNode.ChildNodes.LastOrDefault();
			ParentNode.ChildNodes.Add(this);
		}
	}
}
