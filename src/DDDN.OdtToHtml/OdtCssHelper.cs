/*
DDDN.OdtToHtml.OdtCssHelper
Copyright(C) 2017-2018 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DDDN.OdtToHtml
{
	public static class OdtCssHelper
	{
		private static readonly NumberFormatInfo NumberFormat = new NumberFormatInfo { NumberDecimalSeparator = "." };

		public static bool IsCssNumberValue(string value)
		{
			return Regex.IsMatch(value, "^[+-]?[0-9]+.?([0-9]+)?(px|em|ex|%|in|cm|mm|pt|pc)$");
		}

		public static string GetCssPercentValueRelativeToPage(string relativeValue, OdtPageInfoCalc pageInfoCalc, OdtStyleToStyle.RelativeTo relaticeTo)
		{
			double pageValue = 0;

			if (string.IsNullOrWhiteSpace(relativeValue)
			|| !IsCssNumberValue(relativeValue))
			{
				return null;
			}

			if (GetNumberUnit(relativeValue).Equals("%"))
			{
				return relativeValue;
			}

			if (relaticeTo == OdtStyleToStyle.RelativeTo.Width)
			{
				pageValue = pageInfoCalc.Width;
			}
			else if (relaticeTo == OdtStyleToStyle.RelativeTo.Height)
			{
				pageValue = pageInfoCalc.Height;
			}
			else
			{
				return null;
			}

			relativeValue = GetRealNumber(relativeValue);

			double.TryParse(relativeValue, NumberStyles.Any, NumberFormat, out double relativeValueNo);

			if (pageValue > 0
				&& relativeValueNo > 0
				&& pageValue > relativeValueNo)
			{
				return ((relativeValueNo / pageValue) * 100).ToString(NumberFormat) + "%";
			}
			else
			{
				return relativeValue;
			}
		}

		public static string GetRealNumber(string value, bool nullOrWhiteSpaceToZero = true)
		{
			if (String.IsNullOrWhiteSpace(value))
			{
				if (nullOrWhiteSpaceToZero)
				{
					value = "0";
				}
				else
				{
					throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(value));
				}
			}

			return Regex.Match(value, "[+-]?([0-9]*[.])?[0-9]+").Value;
		}

		public static string GetNumberUnit(string value, bool emptyIfNull = true)
		{
			if (string.IsNullOrWhiteSpace(value))
			{
				if (emptyIfNull)
				{
					return "";
				}
				else
				{
					throw new ArgumentException("message", nameof(value));
				}
			}

			return Regex.Replace(value, @"[+-]?[\d.]", "");
		}

		public static bool IsCssColorValue(string value)
		{
			return Regex.IsMatch(value, @"^(#[0-9a-f]{3}|#(?:[0-9a-f]{2}){2,4}|(rgb|hsl)a?\((-?\d+%?[,\s]+){2,3}\s*[\d\.]+%?\))$");
		}
	}
}
