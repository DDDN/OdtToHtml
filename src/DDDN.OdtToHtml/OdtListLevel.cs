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
			Number,
			Image
		}

		public enum NumberKind
		{
			None,
			Numbers,
			LettersUpper,
			LettersLower,
			RomanUpper,
			RomanLower
		}

		public static class AttrLabelFollowedBy
		{
			public static string Listtab { get; } = "listtab";
			public static string Space { get; } = "space";
			public static string Nothing { get; } = "nothing";
		}

		public static class AttrListTabStopPosition
		{
			public static string LabelWidthAndPosition { get; } = "label-width-and-position";
			public static string LabelAlignment { get; } = "label-alignment";
		}

		public ListKind KindOfList { get; }
		public XElement Element { get; }
		public string StyleName { get; }
		public int Level { get; }
		public string DisplayLevels { get; set; }
		public string BulletChar { get; set; }
		public string NumFormat { get; set; }
		public string NumSuffix { get; set; }
		public string NumPrefix { get; set; }
		public string StyleFontName { get; set; }
		public string TextFontName { get; set; }
		public string LabelFollowedBy { get; set; }
		public string ListLevelPositionAndSpaceMode { get; set; }
		public string PosMarginLeft { get; set; } = "0";
		public string PosFirstLineTextIndent { get; set; } = "0";
		//public string PosMinLabelWidth { get; set; } = "0";
		//public string PosFirstLineIndent { get; set; } = "0";

		private OdtListLevel()
		{
		}

		public OdtListLevel(string styleName, XElement levelElement, int level) : this()
		{
			if (string.IsNullOrWhiteSpace(styleName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(styleName));
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
			else if (levelElement.Name.LocalName.Equals("text:list-level-style-image", StringComparison.InvariantCultureIgnoreCase))
			{
				KindOfList = ListKind.Image;
			}
		}

		public static NumberKind IsKindOfNumber(OdtListLevel odtListLevel)
		{
			if (odtListLevel.KindOfList != ListKind.Number)
			{
				return NumberKind.None;
			}

			switch (odtListLevel.NumFormat)
			{
				case "1":
					return NumberKind.Numbers;
				case "A":
					return NumberKind.LettersUpper;
				case "a":
					return NumberKind.LettersLower;
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
