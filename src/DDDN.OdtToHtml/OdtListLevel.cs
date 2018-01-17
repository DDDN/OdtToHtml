/*
DDDN.OdtToHtml.OdtListLevel
Copyright(C) 2017-2018 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Xml.Linq;

namespace DDDN.OdtToHtml
{
	public class OdtListLevel
	{
		public enum ListKind
		{
			None,
			Bullet,
			Number
		}

		public enum NumberKind
		{
			None,
			Numbers,
			Letters,
			RomanLower,
			RomanUpper
		}

		private OdtListLevel()
		{
		}

		public OdtListLevel(string styleName, XElement levelElement, string level)
		{
			if (string.IsNullOrWhiteSpace(styleName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(styleName));
			}

			if (string.IsNullOrWhiteSpace(level))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(level));
			}

			StyleName = styleName;
			Element = levelElement ?? throw new ArgumentNullException(nameof(levelElement));
			Level = level;

			if (levelElement.Name.LocalName.Equals("list-level-style-bullet", StringComparison.InvariantCultureIgnoreCase))
			{
				KindOfList = ListKind.Bullet;
			}
			else if (levelElement.Name.LocalName.Equals("list-level-style-number", StringComparison.InvariantCultureIgnoreCase))
			{
				KindOfList = ListKind.Number;
			}
		}

		public string StyleName { get; }
		public XElement Element { get; }
		public string Level { get; }
		public ListKind KindOfList { get; }
		public string DisplayLevels { get; set; }
		public string BulletChar { get; set; }
		public string NumFormat { get; set; }
		public string NumSuffix { get; set; }
		public string NumPrefix { get; set; }
		public string FontName { get; set; }
		public string SpaceBefore { get; set; }
		public string SpaceBeforePercent { get; set; }
		public string MarginLeft { get; set; }
		public string MarginLeftPercent { get; set; }
		public string TextIndent { get; set; }

		public static NumberKind IsKindOfNumber(OdtListLevel odtListLevel)
		{
			if (odtListLevel.KindOfList != ListKind.Number)
			{
				return NumberKind.None;
			}

			switch (odtListLevel.NumFormat)
			{
				case "1" :
				return NumberKind.Numbers;
				case "A":
					return NumberKind.Letters;
				case "I":
					return NumberKind.RomanUpper;
				case "i":
					return NumberKind.RomanLower;
				default:
					return NumberKind.Numbers;
			}
		}
	}
}
