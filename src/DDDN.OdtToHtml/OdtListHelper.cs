/*
DDDN.OdtToHtml.OdtListHelper
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

namespace DDDN.OdtToHtml
{
	public static class OdtListHelper
	{
		public const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;
		private static readonly NumberFormatInfo NumberFormat = new NumberFormatInfo { NumberDecimalSeparator = "." };

		public static (int level, OdtListLevel listLevelInfo) CreateListLevelInfo(
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

			var fontStyleName = listLevelElement?.Attribute(XName.Get("style-name", OdtXmlNs.Text))?.Value;
			var fontStyle = OdtStyleHelper.FindStyleElementByNameAttr(fontStyleName, "style", styles);
			var fontStyleTextPropertiesElement = fontStyle?.Element(XName.Get("text-properties", OdtXmlNs.Style));
			var fontName = fontStyleTextPropertiesElement?.Attribute(XName.Get("font-name", OdtXmlNs.Style))?.Value;
			fontName = OdtStyleHelper.HandleFontFamilyStyle(styles, fontName);

			var listLevelInfo = new OdtListLevel(styleName, listLevelElement, level)
			{
				// "<text:list-level-style-number style:num-letter-sync"
				BulletChar = listLevelElement?.Attribute(XName.Get("bullet-char", OdtXmlNs.Text))?.Value,
				DisplayLevels = Convert.ToInt32(listLevelElement?.Attribute(XName.Get("display-levels", OdtXmlNs.Text))?.Value ?? "1"),
				NumFormat = listLevelElement?.Attribute(XName.Get("num-format", OdtXmlNs.Style))?.Value,
				NumPrefix = listLevelElement?.Attribute(XName.Get("num-prefix", OdtXmlNs.Style))?.Value,
				NumSuffix = listLevelElement?.Attribute(XName.Get("num-suffix", OdtXmlNs.Style))?.Value,
				TextFontName = OdtStyleHelper.HandleFontFamilyStyle(styles, textPropertiesElement?.Attribute(XName.Get("font-name", OdtXmlNs.Style))?.Value),
				StyleFontName = fontName,
				LabelFollowedBy = listLevelLabelAlignmentElement?.Attribute(XName.Get("label-followed-by", OdtXmlNs.Text))?.Value,
				ListLevelPositionAndSpaceMode = listLevelPropertiesElement?.Attribute(XName.Get("list-level-position-and-space-mode", OdtXmlNs.Text))?.Value ?? OdtListLevel.AttrListTabStopPosition.LabelWidthAndPosition
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

		public static Dictionary<string, Dictionary<int, OdtListLevel>> CreateListLevelInfos(IEnumerable<XElement> styles)
		{
			var listStyleInfos = new Dictionary<string, Dictionary<int, OdtListLevel>>(StringComparer.InvariantCultureIgnoreCase);

			foreach (var listStyle in styles.Where(p => p.Name.LocalName.Equals("list-style", StrCompICIC)))
			{
				var listStyleName = OdtContentHelper.GetOdtElementAttrValOrNull(listStyle, "name", OdtXmlNs.Style);

				if (listStyleName == null)
				{
					continue;
				}

				listStyleInfos.TryGetValue(listStyleName, out Dictionary<int, OdtListLevel> styleLevelInfos);

				if (styleLevelInfos == null)
				{
					styleLevelInfos = new Dictionary<int, OdtListLevel>();
					listStyleInfos.Add(listStyleName, styleLevelInfos);
					AddStyleListLevels(listStyleName, listStyle, styleLevelInfos, styles);
				}
			}

			return listStyleInfos;
		}

		public static bool TryGetListLevelInfo(OdtContext ctx, OdtListInfo listInfo, out OdtListLevel odtListLevel)
		{
			odtListLevel = null;

			return !string.IsNullOrWhiteSpace(listInfo.RootListInfo.OdtCssClassName)
				&& ctx.OdtListsLevelInfo.TryGetValue(listInfo.RootListInfo.OdtCssClassName, out Dictionary<int, OdtListLevel> odtListLevelKeyVal)
				&& odtListLevelKeyVal.TryGetValue(listInfo.ListLevel, out odtListLevel);
		}

		public static string GetNumberLevelContent(int listItemIndex, OdtListLevel.NumberKind numberKind)
		{
			switch (numberKind)
			{
				case OdtListLevel.NumberKind.Numbers:
					return listItemIndex.ToString();
				case OdtListLevel.NumberKind.LettersUpper:
					return GetLetters(listItemIndex - 1).ToUpper();
				case OdtListLevel.NumberKind.LettersLower:
					return GetLetters(listItemIndex - 1).ToLower();
				case OdtListLevel.NumberKind.RomanUpper:
					return ConvertToRoman(listItemIndex).ToUpper();
				case OdtListLevel.NumberKind.RomanLower:
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

				if (listInfo.OdtCssClassName?.Equals(listItemHtmlInfo.ListInfo.RootListInfo.OdtCssClassName, StrCompICIC) == true
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
			Dictionary<int, OdtListLevel> styleLevelInfos,
			IEnumerable<XElement> styles)
		{
			OdtListLevel parentLevel = null;

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

		public static String GetLetters(int value)
		{
			String s = "";

			do
			{
				s = (char)('A' + (value % 26)) + s;
				value /= 26;
			} while (value-- > 0);

			return s;
		}
	}
}
