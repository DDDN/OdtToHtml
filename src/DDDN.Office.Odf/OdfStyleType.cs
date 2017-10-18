/*
DDDN.Office.Odf.OdfStyleType
Copyright(C) 2017 Lukasz Jaskiewicz(lukasz @jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

namespace DDDN.Office.Odf
{
	public static class OdfStyleType
	{
		public static string AutomaticStyles { get; set; } = "automatic-styles";
		public static string Styles { get; set; } = "styles";
		public static string ListStyle { get; set; } = "list-style";
		public static string DefaultStyle { get; set; } = "default-style";
		public static string Style { get; set; } = "style";
		public static string ListLevelStyleBullet { get; set; } = "list-level-style-bullet";
		public static string ListLevelProperties { get; set; } = "list-level-properties";
		public static string ListLevelLabelAlignment { get; set; } = "list-level-label-alignment";
		public static string ListLevelStyleNumber { get; set; } = "list-level-style-number";
		public static string TableProperties { get; set; } = "table-properties";
		public static string DefaultPageLayout { get; set; } = "default-page-layout";
		public static string PageLayout { get; set; } = "page-layout";
		public static string PageLayoutProperties { get; set; } = "page-layout-properties";
	}
}
