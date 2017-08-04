/*
* DDDN.Net.Html.WStyleInfo
* 
* Copyright(C) 2017 Lukasz Jaskiewicz
* Author: Lukasz Jaskiewicz (lukasz@jaskiewicz.de, devdone@outlook.com)
*
* This program is free software; you can redistribute it and/or modify it under the terms of the
* GNU General Public License as published by the Free Software Foundation; version 2 of the License.
*
* This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
* warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License along with this program; if not, write
* to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using DocumentFormat.OpenXml.Wordprocessing;

namespace DDDN.Office.DOCX
{
    /// <summary>
    /// w:style
    /// </summary>
    public class WStyleInfo
    {
        /// <summary>
        /// w:type
        /// </summary>
        public string StyleType { get; set; }
        /// <summary>
        /// w:styleId
        /// </summary>
        public string StyleId { get; set; }
        /// <summary>
        /// w:basedOn
        /// </summary>
        public BasedOn BasedOnStyle { get; set; }
        /// <summary>
        /// w:rPr
        /// </summary>
        public WStyleRunProperty RunProperty { get; set; } = new WStyleRunProperty();
    }
}
