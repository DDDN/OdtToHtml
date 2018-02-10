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

namespace DDDN.OdtToHtml
{
	public static class OdtListHelper
	{
		public const StringComparison StrCompICIC = StringComparison.InvariantCultureIgnoreCase;
		private static readonly NumberFormatInfo NumberFormat = new NumberFormatInfo { NumberDecimalSeparator = "." };

		public static (int level, OdtListLevel listLevelInfo) CreateListLevelInfo(
			string styleName,
			XElement listLevelElement,
			OdtPageInfoCalc pageInfoCalc,
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
				BulletChar = listLevelElement?.Attribute(XName.Get("bullet-char", OdtXmlNs.Text))?.Value,
				DisplayLevels = listLevelElement?.Attribute(XName.Get("display-levels", OdtXmlNs.Text))?.Value,
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

		public static Dictionary<string, Dictionary<int, OdtListLevel>> CreateListLevelInfos(IEnumerable<XElement> styles, OdtPageInfoCalc pageInfoClac)
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
					AddStyleListLevels(pageInfoClac, listStyleName, listStyle, styleLevelInfos, styles);
				}
			}

			return listStyleInfos;
		}

		public static bool TryGetListLevelInfo(OdtContext ctx, string listStyleName, int listLevel, out OdtListLevel odtListLevel)
		{
			odtListLevel = null;

			if (string.IsNullOrWhiteSpace(listStyleName)
				|| !ctx.OdtListsLevelInfo.TryGetValue(listStyleName, out Dictionary<int, OdtListLevel> listInfo)
				|| !listInfo.TryGetValue(listLevel, out odtListLevel))
			{
				return false;
			}
			else
			{
				return true;
			}
		}

		public static string GetNumberLevelContent(int listItemIndex, OdtListLevel.NumberKind numberKind)
		{
			switch (numberKind)
			{
				case OdtListLevel.NumberKind.Numbers:
					return listItemIndex.ToString();
				case OdtListLevel.NumberKind.LettersUpper:
					return GetLetters(listItemIndex).ToUpper();
				case OdtListLevel.NumberKind.LettersLower:
					return GetLetters(listItemIndex).ToLower();
				case OdtListLevel.NumberKind.RomanUpper:
					return ConvertToRoman(listItemIndex).ToUpper();
				case OdtListLevel.NumberKind.RomanLower:
					return ConvertToRoman(listItemIndex).ToLower();
				default:
					return listItemIndex.ToString();
			}
		}

		public static bool TryGetListItemIndex(OdtHtmlInfo listItemInfo, out int index)
		{
			index = 0;

			if (listItemInfo == null
				|| !listItemInfo.OdtTag.Equals("list-item", StrCompICIC)
				|| !listItemInfo.ParentNode.OdtTag.Equals("list", StrCompICIC))
			{
				return false;
			}

			index = listItemInfo.ParentNode.ChildNodes.IndexOf(listItemInfo) + 1;

			if (TryGetListLevel(listItemInfo, out int targetListLevel, out OdtHtmlInfo rootListInfo))
			{
				var listIndex = rootListInfo.ParentNode.ChildNodes.IndexOf(rootListInfo) - 1;

				for (int i = listIndex; i >= 0; i--)
				{
					var listInfo = rootListInfo.ParentNode.ChildNodes[i] as OdtHtmlInfo;

					if (listInfo.OdtClass?.Equals(rootListInfo.OdtClass, StrCompICIC) == true
						&& (!listInfo.OdtAttrs.TryGetValue("continue-numbering", out string attrValue)
						|| attrValue.Equals("true", StrCompICIC)))
					{
						index += ListTreeWalker(1, targetListLevel, listInfo, listItemInfo, listInfo);
					}
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

		public static bool TryGetListLevel(OdtHtmlInfo odtHtmlInfo, out int listLevel, out OdtHtmlInfo rootListHtmlInfo)
		{
			listLevel = 0;
			rootListHtmlInfo = null;

			var previousHtmlInfo = odtHtmlInfo;

			while (previousHtmlInfo != null)
			{
				if (previousHtmlInfo.OdtTag.Equals("list", StrCompICIC))
				{
					listLevel++;

					if (!String.IsNullOrWhiteSpace(previousHtmlInfo.OdtClass))
					{
						rootListHtmlInfo = previousHtmlInfo;
						return true;
					}
				}

				previousHtmlInfo = previousHtmlInfo.ParentNode;
			}

			return false;
		}

		public static void AddStyleListLevels(
			OdtPageInfoCalc pageInfo,
			string listStyleName,
			XElement listStyleElement,
			Dictionary<int, OdtListLevel> styleLevelInfos,
			IEnumerable<XElement> styles)
		{
			OdtListLevel parentLevel = null;

			foreach (var styleLevelElement in listStyleElement.Elements())
			{
				var listLevelInfo = CreateListLevelInfo(listStyleName, styleLevelElement, pageInfo, styles);
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
