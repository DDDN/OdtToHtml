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
			XElement levelElement,
			Dictionary<int, OdtListLevel> styleLevelInfos,
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

			if (levelElement == null)
			{
				throw new ArgumentNullException(nameof(levelElement));
			}

			var listLevelPropertiesElement = levelElement.Element(XName.Get("list-level-properties", OdtXmlNs.Style));
			var listLevelLabelAlignmentElement = listLevelPropertiesElement?.Element(XName.Get("list-level-label-alignment", OdtXmlNs.Style));
			var textPropertiesElement = levelElement?.Element(XName.Get("text-properties", OdtXmlNs.Style));
			var level = int.Parse(levelElement?.Attribute(XName.Get("level", OdtXmlNs.Text))?.Value);
			styleLevelInfos.TryGetValue(level - 1, out OdtListLevel listInfoParent);

			var listLevelInfo = new OdtListLevel(styleName, levelElement, level)
			{
				BulletChar = levelElement?.Attribute(XName.Get("bullet-char", OdtXmlNs.Text))?.Value,
				DisplayLevels = levelElement?.Attribute(XName.Get("display-levels", OdtXmlNs.Text))?.Value,
				NumFormat = levelElement?.Attribute(XName.Get("num-format", OdtXmlNs.Style))?.Value,
				NumPrefix = levelElement?.Attribute(XName.Get("num-prefix", OdtXmlNs.Style))?.Value,
				NumSuffix = levelElement?.Attribute(XName.Get("num-suffix", OdtXmlNs.Style))?.Value,
				TextFontName = OdtStyleHelper.HandleFontFamilyStyle(styles, textPropertiesElement?.Attribute(XName.Get("font-name", OdtXmlNs.Style))?.Value),
				StyleFontName = OdtStyleHelper.HandleFontFamilyStyle(styles, levelElement?.Attribute(XName.Get("style-name", OdtXmlNs.Text))?.Value)
			};

			// http://docs.oasis-open.org/office/v1.2/part1/cd04/OpenDocument-v1.2-part1-cd04.html

			// old
			var spaceBefore = listLevelPropertiesElement?.Attribute(XName.Get("space-before", OdtXmlNs.Text))?.Value; // the minimum space to allocate to the number or bullet
			var minLabelWidth = listLevelPropertiesElement?.Attribute(XName.Get("min-label-width", OdtXmlNs.Text))?.Value; // the amount to indent before the bullet. This attribute does not appear for the first level of bullet
			var minimumLabelDistance = listLevelLabelAlignmentElement?.Attribute(XName.Get("min-label-distance", OdtXmlNs.Text))?.Value;

			//new
			var listLevelPositionAndSpaceMode = listLevelPropertiesElement?.Attribute(XName.Get("list-level-position-and-space-mode", OdtXmlNs.Text))?.Value; // values: label-alignment, label-width-and-position is default or wehn prop is missing, prop for attr position-and-space-mode
			var labelFollowedBy = listLevelLabelAlignmentElement?.Attribute(XName.Get("label-followed-by", OdtXmlNs.Text))?.Value; // values are listtab, space and nothing
			var listTabStopPosition = listLevelLabelAlignmentElement?.Attribute(XName.Get("list-tab-stop-position", OdtXmlNs.Text))?.Value; // relevant if label-followed-by=listtab

			// The new attributes first-line-indent and indent-at are only relevant, if attribute label-followed-by is defined.
			// As long as the paragraph doesn't specify its own indent attributes first-line and/or left-margin,
			// the new attributes first-line-indent and indent-at of the corresponding1 list level of the applied list style are used.
			var textIntend = listLevelLabelAlignmentElement?.Attribute(XName.Get("text-indent", OdtXmlNs.XslFoCompatible))?.Value; // prop for attr first-line-indent, optional, 0 if missing
			var marginLeft = listLevelLabelAlignmentElement?.Attribute(XName.Get("margin-left", OdtXmlNs.XslFoCompatible))?.Value; // prop for attr indent-at, optional, 0 if missing

			if (!String.IsNullOrWhiteSpace(spaceBefore))
			{
				double.TryParse(OdtCssHelper.GetRealNumber(spaceBefore), NumberStyles.Any, NumberFormat, out double spaceBeforeNo);
				double.TryParse(OdtCssHelper.GetRealNumber(minLabelWidth), NumberStyles.Any, NumberFormat, out double minLabelWidthNo);
				double.TryParse(OdtCssHelper.GetRealNumber(marginLeft), NumberStyles.Any, NumberFormat, out double marginLeftNo);
				double.TryParse(OdtCssHelper.GetRealNumber(textIntend), NumberStyles.Any, NumberFormat, out double textIntendNo);

				listLevelInfo.Calc.SpaceBefore = spaceBeforeNo;

				textIntendNo = marginLeftNo + textIntendNo;
				listLevelInfo.PosTextIndent = textIntendNo.ToString(NumberFormat) + OdtCssHelper.GetNumberUnit(textIntend);

				marginLeftNo = marginLeftNo - spaceBeforeNo - minLabelWidthNo;
				listLevelInfo.PosFirstLineIndent = marginLeftNo.ToString(NumberFormat) + OdtCssHelper.GetNumberUnit(marginLeft);

				listLevelInfo.PosSpaceBefore = (spaceBeforeNo - listInfoParent?.Calc.SpaceBefore ?? 0).ToString(NumberFormat) + OdtCssHelper.GetNumberUnit(spaceBefore);
				listLevelInfo.PosLabelWidth = minLabelWidth;
			}

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

		public static OdtListLevel GetListLevelInfo(OdtContext ctx, string listStyleName, int listLevel)
		{
			if (string.IsNullOrWhiteSpace(listStyleName)
				|| !ctx.OdtListsLevelInfo.TryGetValue(listStyleName, out Dictionary<int, OdtListLevel> listInfo)
				|| !listInfo.TryGetValue(listLevel, out OdtListLevel listLevelInfo))
			{
				return null;
			}
			else
			{
				return listLevelInfo;
			}
		}

		public static string GetListClassName(OdtInfo odtInfo)
		{
			if (odtInfo == null)
			{
				throw new ArgumentNullException(nameof(odtInfo));
			}

			string className = null;

			do
			{
				if (odtInfo.OdtNode.Equals("list", StrCompICIC))
				{
					className = odtInfo.ClassName;
				}

				odtInfo = odtInfo.ParentNode;
			}
			while (odtInfo != null);

			return className;
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

		public static int GetListItemIndex(OdtInfo odtInfo)
		{
			int index = 0;

			var listNode = odtInfo.ParentNode;
			var listStyleName = listNode.ClassName;

			while (listNode != null)
			{
				if (listNode.OdtNode.Equals("list", StrCompICIC)
					&& listNode.ClassName.Equals(listStyleName, StrCompICIC)) // TODO GetListClassName
				{
					OdtInfo itemNode = listNode.ChildNodes?.LastOrDefault();

					while (itemNode != null)
					{
						if (itemNode.OdtNode.Equals("list-item", StrCompICIC))
						{
							index++;
							itemNode = itemNode.PreviousSibling;
						}
					}
				}

				listNode.OdtAttrs.TryGetValue("continue-numbering", out string continueNumbering);

				if (continueNumbering == null || continueNumbering.Equals("true", StrCompICIC))
				{
					listNode = listNode.PreviousSibling;
				}
			}

			return index;
		}

		public static int GetListLevel(OdtInfo odtInfo)
		{
			if (odtInfo == null)
			{
				throw new ArgumentNullException(nameof(odtInfo));
			}

			int listLevel = 0;

			while (odtInfo != null)
			{
				if (odtInfo.OdtNode.Equals("list", StrCompICIC))
				{
					listLevel++;
				}

				odtInfo = odtInfo.ParentNode;
			}

			return listLevel;
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
				var listLevelInfo = CreateListLevelInfo(listStyleName, styleLevelElement, styleLevelInfos, pageInfo, styles);
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

		public static String GetLetters(int valuse)
		{
			String s = "";

			do
			{
				s = (char)('A' + (valuse % 26)) + s;
				valuse /= 26;
			} while (valuse-- > 0);

			return s;
		}
	}
}
