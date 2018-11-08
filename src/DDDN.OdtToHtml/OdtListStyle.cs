/*
DDDN.OdtToHtml.OdtListStyle
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
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DDDN.OdtToHtml.Conversion;
using static DDDN.OdtToHtml.OdtHtmlInfo;
using static DDDN.OdtToHtml.OdtStyle;

namespace DDDN.OdtToHtml
{
	public class OdtListStyle
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

		public const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;
		private static readonly NumberFormatInfo NumberFormat = new NumberFormatInfo { NumberDecimalSeparator = "." };

		public ListKind KindOfList { get; }
		public XElement Element { get; }
		public string StyleName { get; }
		public int Level { get; }
		public string LevelStyleName { get; set;  }
		public int DisplayLevels { get; set; }
		public string BulletChar { get; set; }
		public string NumFormat { get; set; }
		public string NumSuffix { get; set; }
		public string NumPrefix { get; set; }
		public string NumLetterSync { get; set; }
		public string StyleFontName { get; set; }
		public string TextFontName { get; set; }
		public string LabelFollowedBy { get; set; }
		public string ListLevelPositionAndSpaceMode { get; set; }
		public string PosMarginLeft { get; set; } = "0";
		public string PosFirstLineTextIndent { get; set; } = "0";
		//public string PosMinLabelWidth { get; set; } = "0";
		//public string PosFirstLineIndent { get; set; } = "0";

		private OdtListStyle()
		{
		}

		public OdtListStyle(string styleName, XElement levelElement, int level) : this()
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

		public static (int level, OdtListStyle listLevelInfo) CreateListLevelInfo(
			string styleName,
			XElement listLevelElement,
			IEnumerable<XElement> styles)
		{
			if (styles == null)
			{
				throw new ArgumentNullException(nameof(styles));
			}

			if (string.IsNullOrWhiteSpace(styleName))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(styleName));
			}

			if (listLevelElement == null)
			{
				throw new ArgumentNullException(nameof(listLevelElement));
			}

			var listLevelPropertiesElement = listLevelElement.Element(XName.Get("list-level-properties", OdtXmlNs.Style));
			var listLevelLabelAlignmentElement = listLevelPropertiesElement?.Element(XName.Get("list-level-label-alignment", OdtXmlNs.Style));
			var textPropertiesElement = listLevelElement?.Element(XName.Get("text-properties", OdtXmlNs.Style));
			var level = int.Parse(listLevelElement?.Attribute(XName.Get("level", OdtXmlNs.Text))?.Value);

			var levelStyleName = listLevelElement?.Attribute(XName.Get("style-name", OdtXmlNs.Text))?.Value;
			var levelStyle = OdtStyle.FindStyleElementByNameAttr(levelStyleName, StyleType.style, styles);
			var levelStyleTextPropertiesElement = levelStyle?.Element(XName.Get("text-properties", OdtXmlNs.Style));
			var levelStyleFontName = levelStyleTextPropertiesElement?.Attribute(XName.Get("font-name", OdtXmlNs.Style))?.Value;
			levelStyleFontName = OdtStyle.HandleFontFamilyStyle(styles, levelStyleFontName);

			var listLevelInfo = new OdtListStyle(styleName, listLevelElement, level)
			{
				// "<text:list-level-style-number style:num-letter-sync"
				LevelStyleName = levelStyleName,
				BulletChar = listLevelElement?.Attribute(XName.Get("bullet-char", OdtXmlNs.Text))?.Value,
				DisplayLevels = Convert.ToInt32(listLevelElement?.Attribute(XName.Get("display-levels", OdtXmlNs.Text))?.Value ?? "1"),
				NumFormat = listLevelElement?.Attribute(XName.Get("num-format", OdtXmlNs.Style))?.Value,
				NumPrefix = listLevelElement?.Attribute(XName.Get("num-prefix", OdtXmlNs.Style))?.Value,
				NumSuffix = listLevelElement?.Attribute(XName.Get("num-suffix", OdtXmlNs.Style))?.Value,
				TextFontName = OdtStyle.HandleFontFamilyStyle(styles, textPropertiesElement?.Attribute(XName.Get("font-name", OdtXmlNs.Style))?.Value),
				StyleFontName = levelStyleFontName,
				LabelFollowedBy = listLevelLabelAlignmentElement?.Attribute(XName.Get("label-followed-by", OdtXmlNs.Text))?.Value,
				ListLevelPositionAndSpaceMode = listLevelPropertiesElement?.Attribute(XName.Get("list-level-position-and-space-mode", OdtXmlNs.Text))?.Value ?? OdtListStyle.AttrListTabStopPosition.LabelWidthAndPosition
			};

			// http://docs.oasis-open.org/office/
			var spaceBefore = listLevelPropertiesElement?.Attribute(XName.Get("space-before", OdtXmlNs.Text))?.Value;
			var textIntend = listLevelLabelAlignmentElement?.Attribute(XName.Get("text-indent", OdtXmlNs.XslFoCompatible))?.Value;
			var marginLeft = listLevelLabelAlignmentElement?.Attribute(XName.Get("margin-left", OdtXmlNs.XslFoCompatible))?.Value;
			//var minLabelWidth = listLevelPropertiesElement?.Attribute(XName.Get("min-label-width", OdtXmlNs.Text))?.Value;
			//var listTabStopPosition = listLevelLabelAlignmentElement?.Attribute(XName.Get("list-tab-stop-position", OdtXmlNs.Text))?.Value;
			//var minimumLabelDistance = listLevelLabelAlignmentElement?.Attribute(XName.Get("min-label-distance", OdtXmlNs.Text))?.Value;

			double.TryParse(OdtCssHelper.GetRealNumber(marginLeft), NumberStyles.Any, NumberFormat, out double marginLeftNo);
			double.TryParse(OdtCssHelper.GetRealNumber(textIntend), NumberStyles.Any, NumberFormat, out double textIntendNo);
			double.TryParse(OdtCssHelper.GetRealNumber(spaceBefore), NumberStyles.Any, NumberFormat, out double spaceBeforeNo);
			//double.TryParse(OdtCssHelper.GetRealNumber(minLabelWidth), NumberStyles.Any, NumberFormat, out double minLabelWidthNo);
			//double.TryParse(OdtCssHelper.GetRealNumber(listTabStopPosition), NumberStyles.Any, NumberFormat, out double listTabStopPositionNo);
			//double.TryParse(OdtCssHelper.GetRealNumber(minLabelWidth), NumberStyles.Any, NumberFormat, out double minLabelWidthNo);

			var unit = OdtCssHelper.GetNumberUnit(spaceBefore ?? marginLeft ?? textIntend);
			listLevelInfo.PosMarginLeft = marginLeft ?? "0";
			listLevelInfo.PosFirstLineTextIndent = spaceBefore ?? (marginLeftNo + textIntendNo).ToString(NumberFormat) + unit;

			return (level, listLevelInfo);
		}

		public static Dictionary<string, Dictionary<int, OdtListStyle>> CreateListLevelInfos(IEnumerable<XElement> styles)
		{
			var listStyleInfos = new Dictionary<string, Dictionary<int, OdtListStyle>>(StringComparer.InvariantCultureIgnoreCase);

			foreach (var listStyle in styles.Where(p => p.Name.LocalName.Equals(StyleType.list_style, StrCompICIC)))
			{
				var listStyleName = OdtContentHelper.GetOdtElementAttrValOrNull(listStyle, "name", OdtXmlNs.Style);

				if (listStyleName == null)
				{
					continue;
				}

				listStyleInfos.TryGetValue(listStyleName, out Dictionary<int, OdtListStyle> styleLevelInfos);

				if (styleLevelInfos == null)
				{
					styleLevelInfos = new Dictionary<int, OdtListStyle>();
					listStyleInfos.Add(listStyleName, styleLevelInfos);
					AddStyleListLevels(listStyleName, listStyle, styleLevelInfos, styles);
				}
			}

			return listStyleInfos;
		}

		public static bool TryGetListLevelInfo(OdtContext ctx, OdtListInfo listInfo, out OdtListStyle odtListLevel)
		{
			odtListLevel = null;

			return !string.IsNullOrWhiteSpace(listInfo.RootListInfo?.OdtStyleName)
				&& ctx.OdtListsLevelInfo.TryGetValue(listInfo.RootListInfo.OdtStyleName, out Dictionary<int, OdtListStyle> odtListLevelKeyVal)
				&& odtListLevelKeyVal.TryGetValue(listInfo.ListLevel, out odtListLevel);
		}

		public static string GetNumberLevelContent(int listItemIndex, OdtListStyle.NumberKind numberKind)
		{
			switch (numberKind)
			{
				case NumberKind.Numbers:
					return listItemIndex.ToString();
				case NumberKind.LettersUpper:
					return GetLetters(listItemIndex - 1).ToUpper();
				case NumberKind.LettersLower:
					return GetLetters(listItemIndex - 1).ToLower();
				case NumberKind.RomanUpper:
					return ConvertToRoman(listItemIndex).ToUpper();
				case NumberKind.RomanLower:
					return ConvertToRoman(listItemIndex).ToLower();
				default:
					return listItemIndex.ToString();
			}
		}

		public static bool TryGetListItemIndex(OdtHtmlInfo listItemHtmlInfo, out int index)
		{
			index = 0;

			if (listItemHtmlInfo?.OdtTag.Equals("list-item", StrCompICIC) != true
				|| !listItemHtmlInfo.ParentNode.OdtTag.Equals("list", StrCompICIC))
			{
				return false;
			}

			index = listItemHtmlInfo.ParentNode.ChildNodes.IndexOf(listItemHtmlInfo) + 1;

			var listIndex = listItemHtmlInfo.ListInfo.RootListInfo.ParentNode.ChildNodes.IndexOf(listItemHtmlInfo.ListInfo.RootListInfo) - 1;

			for (int i = listIndex; i >= 0; i--)
			{
				var listInfo = listItemHtmlInfo.ListInfo.RootListInfo.ParentNode.ChildNodes[i] as OdtHtmlInfo;

				if (listInfo.OdtStyleName?.Equals(listItemHtmlInfo.ListInfo.RootListInfo.OdtStyleName, StrCompICIC) == true
					&& (!listInfo.OdtAttrs.TryGetValue("continue-numbering", out string attrValue)
					|| attrValue.Equals("true", StrCompICIC)))
				{
					index += ListTreeWalker(1, listItemHtmlInfo.ListInfo.ListLevel, listInfo, listItemHtmlInfo, listInfo);
				}
			}

			return true;
		}

		public static int ListTreeWalker(
			int currentLevel,
			int targetLevel,
			OdtHtmlInfo currentItemInfo,
			OdtHtmlInfo originalItemInfo,
			OdtHtmlInfo firstLevelListItem)
		{
			var count = 0;

			if (currentItemInfo == null
				|| originalItemInfo == null
				|| firstLevelListItem == null
				|| currentLevel > targetLevel
				|| (!currentItemInfo.OdtTag.Equals("list-item", StrCompICIC)
					&& !currentItemInfo.OdtTag.Equals("list", StrCompICIC)))
			{
				return count;
			}

			if (currentItemInfo.OdtTag.Equals("list-item", StrCompICIC))
			{
				var lists = currentItemInfo.ChildNodes
						.OfType<OdtHtmlInfo>()
						.Where(p => p.OdtTag.Equals("list", StrCompICIC));

				foreach (var item in lists)
				{
					count += ListTreeWalker(currentLevel, targetLevel, item, originalItemInfo, firstLevelListItem);
				}
			}

			if (currentItemInfo.OdtTag.Equals("list", StrCompICIC))
			{
				var listItems = currentItemInfo.ChildNodes
						.OfType<OdtHtmlInfo>()
						.Where(p => p.OdtTag.Equals("list-item", StrCompICIC));

				if (currentLevel == targetLevel)
				{
					return listItems.Count();
				}
				else
				{
					foreach (var item in listItems)
					{
						count += ListTreeWalker(++currentLevel, targetLevel, item, originalItemInfo, firstLevelListItem);
					}
				}
			}

			return count;
		}

		public static void AddStyleListLevels(
			string listStyleName,
			XElement listStyleElement,
			Dictionary<int, OdtListStyle> styleLevelInfos,
			IEnumerable<XElement> styles)
		{
			OdtListStyle parentLevel = null;

			foreach (var styleLevelElement in listStyleElement.Elements())
			{
				var listLevelInfo = CreateListLevelInfo(listStyleName, styleLevelElement, styles);
				styleLevelInfos.Add(listLevelInfo.level, listLevelInfo.listLevelInfo);
				parentLevel = listLevelInfo.listLevelInfo;
			}
		}

		public static string ConvertToRoman(int value)
		{
			int[] decimalValue = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
			string[] romanNumeral = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
			var result = "";
			var residual = value;

			for (var i = 0; i < decimalValue.Length; i++)
			{
				if (residual >= 2000)
				{
					result += "M";
					residual -= 1000;
					i--;
					continue;
				}

				while (decimalValue[i] <= residual)
				{
					result += romanNumeral[i];
					residual -= decimalValue[i];
				}
			}

			return result;
		}

		public static string GetLetters(int value)
		{
			string s = "";

			do
			{
				s = (char)('A' + (value % 26)) + s;
				value /= 26;
			} while (value-- > 0);

			return s;
		}

		public static NumberKind IsKindOfNumber(OdtListStyle odtListLevel)
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
